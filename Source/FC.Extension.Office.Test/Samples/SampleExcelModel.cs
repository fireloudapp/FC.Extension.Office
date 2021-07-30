using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FC.Extension.Office.Test.Samples
{
    public class SampleExcelModel
    {
        public DateTime OrderData { get; set; }
        public string Region { get; set; }
        public string Rep { get; set; }
        public string Item { get; set; }
        public int Units { get; set; }

        public int UnitCost { get; set; }
        public int Total { get; set; }
    }
}
