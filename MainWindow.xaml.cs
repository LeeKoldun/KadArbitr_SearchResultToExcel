using AngleSharp;
using AngleSharp.Dom;
using Aspose.Cells;
using CefSharp;
using CefSharp.OffScreen;
using Refit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace KadArbitr_SearchResultToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string desktopPath, host;
        string? wasm, pr_fp;
        string? input;
        bool RainbowText = false;

        public MainWindow()
        {
            InitializeComponent();

            desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\";

            host = "https://kad.arbitr.ru";
        }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            input = TbxInput.Text.Trim();
            if (input == string.Empty)
            {
                Error("Введите данные ИНН!");
                return;
            }

            BtnSearch.IsEnabled = false;
            BtnText.Text = "Ожидайте...";
            await Task.Delay(10);

            StartProcess();
        }

        private async void StartProcess()
        {
            // Создаем клиента для отправки запросов //
            var client = RestService.For<IPostInfo>(host);

            // Создаем тестовый запрос //
            var testRequest = new PostRequest
            {
                Page = 1,
                Count = 25,
                Courts = { },
                DateFrom = null,
                DateTo = null,
                Sides = new List<PostRequest.Side>
                {
                    new PostRequest.Side
                    {
                        Name = input,
                        Type = -1,
                        ExactMatch = false
                    }
                },
                Judges = { },
                CaseNumbers = { },
                WithVKSInstances = false
            };

            // С помощью CefSharp симулируем заход на сайт и достаем нужные куки для доступа //
            // ЧЕРЕЗ БРАНДМАУЭР НАДО РАЗРЕШИТЬ CefSharp ДОСТУП (Ну, или его можно проигнорировать) //
            using (var browser = new ChromiumWebBrowser(host))
            {
                // Ждём, пока сайт загрузится //
                var initialLoadResponse = await browser.WaitForInitialLoadAsync();

                if (!initialLoadResponse.Success)
                {
                    throw new Exception(string.Format("Page load failed with ErrorCode:{0}, HttpStatusCode:{1}", initialLoadResponse.ErrorCode, initialLoadResponse.HttpStatusCode));
                }

                // Программно нажимаем на кнопку поиска на сайте, чтобы сгенерировались куки для доступа //
                browser.ExecuteScriptAsync("document.getElementsByClassName('b-button-container')[0].click()");

                BtnText.Text = "Крадём куки :^)";
                TextRainbowColor(true);
                // Даём браузеру время загрузить куки //
                await Task.Delay(5000);
                TextRainbowColor(false);

                // Воруем куки :D //
                var cookies = await browser.GetCookieManager().VisitAllCookiesAsync();
                var cookieList = cookies.Select(cookie => $"{cookie.Name}={cookie.Value}");

                wasm = cookieList.SingleOrDefault(c => c.StartsWith("wasm"));
                pr_fp = cookieList.SingleOrDefault(c => c.StartsWith("pr_fp"));
            }

            if (wasm is null || pr_fp is null)
            {
                Error($"Упс! Не удалось собрать нужные куки!\nПопробуйте ещё раз!");
                return;
            }
            // Завершаем сеанс браузера после успешного преступления //
            Cef.Shutdown();

            MessageBox.Show($"Нужные куки получены!\n\n{wasm}\n{pr_fp}", "Куки успешно украдены!", MessageBoxButton.OK, MessageBoxImage.Information);

            string response;
            try
            {
                // Отправляем тестовый запрос //
                response = await client.PostInformation(testRequest, $"{wasm}; {pr_fp}");
                // Берём нужный тег с числом страниц //
                string[] lines = response.Split('\n');
                string? lastLine = lines.SingleOrDefault(l => l.StartsWith("<input type=\"hidden\" id=\"documentsPagesCount\""));
                // 5 - позиция, где находится цифра //
                int pagesCount = int.Parse(lastLine.Split('\"')[5]);

                if (pagesCount == 0)
                {
                    Error($"Ой-ёй!\nДанных нет!");
                    return;
                }

                BtnText.Text = "Загрузка...";
                await Task.Delay(5000);

                // Настраиваем HTML парсер //
                var config = Configuration.Default;
                using var context = BrowsingContext.New(config);

                // Создаем книгу Excel //
                Workbook book = new Workbook();
                List<string[]> cards = new List<string[]>();

                // Загружаем все страницы с интервалом в 5 сек, чтобы не словить защиту от DDOS //
                for (int i = 1; i <= pagesCount; i++)
                {
                    BtnText.Text = $"Загрузка...\nСтраниц: {i}/{pagesCount}";

                    var request = new PostRequest
                    {
                        Page = i,
                        Count = 25,
                        Courts = { },
                        DateFrom = null,
                        DateTo = null,
                        Sides = new List<PostRequest.Side>
                        {
                            new PostRequest.Side
                            {
                                Name = input,
                                Type = -1,
                                ExactMatch = false
                            }
                        },
                        Judges = { },
                        CaseNumbers = { },
                        WithVKSInstances = false
                    };

                    response = await client.PostInformation(request, $"{wasm}; {pr_fp}");

                    // Настраиваем HTML парсер //
                    var htmlConfig = Configuration.Default;
                    using var htmlContext = BrowsingContext.New(htmlConfig);

                    using var doc = await htmlContext.OpenAsync(data => data.Content(response));
                    var rows = doc.QuerySelectorAll("div.b-container");
                    int cardsCount = rows.Count() / 4;

                    for (int j = 0; j < cardsCount; j++)
                    {
                        cards.Add(GetHtmlRow(rows, j));
                    }


                    await Task.Delay(5000);
                }

                // Начинаем конвертировать полученные данные //
                BtnText.Text = "Сохранение данных в формате Excel...";
                await Task.Delay(100);

                string formatedInput = input.Replace("\\", "")
                                            .Replace("/", "")
                                            .Replace(":", "")
                                            .Replace("*", "")
                                            .Replace("?", "")
                                            .Replace("\"", "")
                                            .Replace("<", "")
                                            .Replace(">", "")
                                            .Replace("|", "");
                string tableName = $"КадАрбитр ({formatedInput})";

                // Вставляем заголовки для данных //
                Worksheet sheet = book.Worksheets[0];
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Номер дела:");
                cells["B1"].PutValue("ФИО судьи / Суд:");
                cells["C1"].PutValue("Истцы:");
                cells["D1"].PutValue("Ответчики:");

                Aspose.Cells.Style style = cells[0, 0].GetStyle();
                style.IsTextWrapped = true;

                int rowsCount = cards.Count();

                // Заполняем ячейки данными //
                for (int i = 0; i < rowsCount; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        var cell = cells[i + 1, j];

                        cell.PutValue(cards[i][j]);
                        cell.SetStyle(style);
                    }
                }
                sheet.AutoFitRows();
                sheet.AutoFitColumns();

                book.Save(desktopPath + $"{tableName}.xlsx", SaveFormat.Auto);

                BtnText.Text = $"Успешно!\nСтраниц: {pagesCount}";

                MessageBox.Show($"Данные успешно загружены!\n\nТаблица Excel сохранена на вашем рабочем столе под именем: \"{tableName}\"\n\nДанные находятся на листе Sheet1", "Успех!", MessageBoxButton.OK, MessageBoxImage.Information);
                this.Close();
            }
            catch (ApiException exception)
            {
                Error($"Ошибка!\n\nСкорее всего отсутствует или неверное значение куки \"wasm\" и \"pr_fp\"\n\n{exception}");
                throw;
            }
            catch (Exception)
            {
                Error($"Ошибка!\n\nНу... что-то пошло не так и мы точно не знаем что...");
                throw;
            }
        }

        private void Error(string errorMsg)
        {
            MessageBox.Show(errorMsg, "Ошибка!", MessageBoxButton.OK, MessageBoxImage.Error);

            BtnSearch.IsEnabled = true;
            BtnText.Text = "Найти и конвертировать";
        }

        private async void TextRainbowColor(bool toggle)
        {
            RainbowText = toggle;

            while (RainbowText)
            {
                Random rnd = new Random();

                byte r = (byte)rnd.Next(0, 256);
                byte g = (byte)rnd.Next(0, 256);
                byte b = (byte)rnd.Next(0, 256);

                BtnText.Foreground = new SolidColorBrush(Color.FromRgb(r, g, b));

                await Task.Delay(200);
            }
            if (!RainbowText) BtnText.Foreground = Brushes.Black;
        }

        private string[] GetHtmlRow(IHtmlCollection<IElement> rows, int index)
        {
            /*
                
                Так как строчка не одна - потребуется брать 4 * [индекс строки].
                4 - число столбцов в строке
            
            */
            int rowIndex = 4 * index;
            // Подготовка к конвертации html //
            var card = rows[rowIndex];
            var judges = rows[rowIndex + 1].QuerySelectorAll("div");

            var plaintiffs = rows[rowIndex + 2].QuerySelectorAll("span.js-rolloverHtml");
            int plainCount = plaintiffs.Count();

            var respondents = rows[rowIndex + 3].QuerySelectorAll("span.js-rolloverHtml");
            int respCount = respondents.Count();

            // Процесс конвертации //
            //string date = card.QuerySelector("span").InnerHtml;
            string caseNum = card.QuerySelector("a").InnerHtml.Trim();

            int judCount = judges.Count();

            string judge = judges[0].InnerHtml;
            string court = judCount > 1 ? judges[1].InnerHtml : "Судья неизвестен";

            string plaintiffList = "";

            for (int i = 0; i < plainCount; i++)
            {
                var pl = plaintiffs[i];
                string name;
                if (!pl.InnerHtml.Contains("<strong>"))
                    name = "Слишком много истцов! Найдите дело для подробной информации";
                else
                    name = pl.QuerySelector("strong").InnerHtml;


                plaintiffList += name;

                if (i != plainCount - 1) plaintiffList += "\n\n";
            }

            string respondentList = "";

            for (int i = 0; i < respCount; i++)
            {
                var res = respondents[i];
                string name;
                if (!res.InnerHtml.Contains("<strong>"))
                    name = "Слишком много ответчиков! Найдите дело для подробной информации";
                else
                    name = res.QuerySelector("strong").InnerHtml;

                respondentList += name;

                if (i != respCount - 1) respondentList += "\n\n";
            }

            string[] cardData = {
                caseNum,
                string.Join("\n\n", judge, court),
                plaintiffList,
                respondentList
            };

            return cardData;
        }
    }
}
