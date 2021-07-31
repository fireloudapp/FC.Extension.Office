using Xunit;
using Shouldly;
using Xunit.Abstractions;
using System.Collections.Generic;
using System.Data;
using FC.Core.Extension.DataHandlers;
using FC.Extension.Office.Test.Samples;
using FC.Extension.Office.ExcelHandlers;


namespace FC.Extension.Office.Test.ExcelHanders
{
    public class ToDataTableExtensionTest
    {
        private readonly ITestOutputHelper _output;
        public ToDataTableExtensionTest(ITestOutputHelper output)
        {
            this._output = output;
        }
        [Fact]
        public void ToDataTable_Test()
        {
            string excelPath = @"Samples/SampleData.xlsx";
            IList<SampleExcelModel> dataList = ConvertToListFromExcel.ConvertToList<SampleExcelModel>(excelPath, "SalesOrders");
            
            dataList.ShouldNotBeNull();
            DataTable dt = dataList.ToDataTable();

            if (dataList != null)
            {
                _output.WriteLine($"Rows {dt.Rows.Count} Column { dt.Columns.Count }");
            }
        }
    }
}
