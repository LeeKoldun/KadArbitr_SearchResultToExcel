using Refit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace KadArbitr_SearchResultToExcel
{
    public interface IPostInfo
    {
        [Post("/Kad/SearchInstances")]
        Task<string> PostInformation([Body] PostRequest request, [Header("cookie")] string cookies);
    }
}