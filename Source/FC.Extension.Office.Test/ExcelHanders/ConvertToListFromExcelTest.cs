using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit.Abstractions;
using Xunit;
using FC.Core.Extension.Office.ExcelHandlers;
using FC.Extension.Office.Test.Samples;
using System.IO;
using Shouldly;

namespace FC.Extension.Office.Test.ExcelHanders
{
    public class ConvertToListFromExcelTest
    {
        private readonly ITestOutputHelper _output;
        public ConvertToListFromExcelTest(ITestOutputHelper output)
        {
            this._output = output;
        }

        #region Convert to Excel
        [Fact]
        public void ToListToExcel_Test()
        {
            List<SampleExcelModel> dataList = new List<SampleExcelModel>();

            for (int i = 0; i < 10; i++)
            {
                dataList.Add(new SampleExcelModel()
                {
                    Item = "Teest " + i,
                    OrderData = DateTime.Now.AddDays(i),
                    Region = "IN",
                    Rep = "Got",
                    UnitCost = 50 * i,
                    Units = 5,
                    Total = (50 * i) * 5
                });
            }
            dataList.ToExcel("Sample.xlsx");
            bool result = File.Exists("Sample.xlsx");
            result.ShouldBeTrue();
            dataList.ShouldNotBeNull();

            _output.WriteLine($"List to Excel Length Lenght : {dataList.Count}");
        }
        #endregion
    }
}
