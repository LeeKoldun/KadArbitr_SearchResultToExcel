using Refit;
using System.Threading.Tasks;

namespace KadArbitr_SearchResultToExcel
{
    public interface IPostInfo
    {
        [Post("/Kad/SearchInstances")]
        Task<string> PostInformation([Body] PostRequest request, [Header("cookie")] string cookies);
    }
}