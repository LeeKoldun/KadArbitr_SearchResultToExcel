using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace KadArbitr_SearchResultToExcel
{
    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class PostRequest
    {
        public int Page { get; set; }

        public int Count { get; set; }

        public List<object>? Courts { get; set; }

        public object? DateFrom { get; set; }

        public object? DateTo { get; set; }

        public List<Side>? Sides { get; set; }

        public List<object>? Judges { get; set; }

        public List<object>? CaseNumbers { get; set; }

        public bool WithVKSInstances { get; set; }

        public class Side
        {
            public string? Name { get; set; }

            public int Type { get; set; }

            public bool ExactMatch { get; set; }
        }

    }
}
