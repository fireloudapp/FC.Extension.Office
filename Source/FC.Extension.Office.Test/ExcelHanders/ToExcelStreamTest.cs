using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using FC.Core.Extension.Office.ExcelHandlers;
using FC.Extension.Office.Test.Samples;
using Shouldly;
using Xunit;
using Xunit.Abstractions;

namespace FC.Extension.Office.Test.ExcelHanders
{
    public class ToExcelStreamTest
    {
        private readonly ITestOutputHelper _output;
        public ToExcelStreamTest(ITestOutputHelper output)
        {
            this._output = output;
        }

        [Fact]
        public void ToExcelStream_Test()
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
                }) ;
            }
            MemoryStream ms
                 = new MemoryStream();

            ms = dataList.ToExcelAsStream();
            ms.ShouldNotBeNull();
            _output.WriteLine($"MemoryStream Lenght : {ms.Length}");
        }
    }
}
