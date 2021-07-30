# FC.Extension.Office


[![Version](https://img.shields.io/nuget/v/FC.Extension.Office.svg)](https://www.nuget.org/packages/FC.Extension.Office/)
[![Downloads](https://img.shields.io/nuget/dt/FC.Extension.Office.svg)](https://www.nuget.org/packages/FC.Extension.Office/)


✅ **Project status: active**.

FC.Extension.Office is a library which adds reduces the coding effort for the development team in terms of handling Excel, Document, PPT all basic functionality. You may required literally a single line of code to handle the day to day operation.

This library supports 
+ Excel
+ Document
+ PPT

## Download

- [NuGet](https://www.nuget.org/packages/FC.Extension.Office/): Install-Package FC.Extension.Office

## Features

- Ready to use in the project as an extension
- Convert Excel data to IList
- Read Excel data as Stream
- Convert IList to Excel Data
- Targets .NET Core 3.1+
- Dependent on Spreedsheet Lite.

## Usage

### Quick start
#### Step 1: Setup the Package

```csharp
Install-Package FC.Extension.Office
```

#### Step 2:Initial Configuration

```csharp
using FC.Extension.Office;
	
````

#### Step 3:Data moved to Excel as stream

```csharp
..

List<SampleExcelModel> dataList = new List<SampleExcelModel>();
..
..//Add some data in th datalist
..
MemoryStream ms = new MemoryStream();
ms = dataList.ToExcelAsStream();
//Now all the excel data was in our memory, which is ready to save as Excel file.

````
#### Step 4:Excel to DataTable

```csharp
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

````


> ⚠️ All Set. Use the same method for other extension methods.
The full featured document available in the [Gitbook](https://app.gitbook.com/@sr-firecloud/s/fc-extension),
 

## Other Extension

- [AWS](https://www.nuget.org/packages/FC.Extension.AWS, "AWS Extension") 
	* [![Version](https://img.shields.io/nuget/v/FC.Core.Extension.svg)](https://www.nuget.org/packages/FC.Core.Extension/)
[![Downloads](https://img.shields.io/nuget/dt/FC.Core.Extension.svg)](https://www.nuget.org/packages/FC.Core.Extension/)

- [HTTP](https://www.nuget.org/packages/FC.Extension.HTTP/,"HTTP")
	* [![Version](https://img.shields.io/nuget/v/FC.Extension.HTTP.svg)](https://www.nuget.org/packages/FC.Extension.HTTP/)
[![Downloads](https://img.shields.io/nuget/dt/FC.Extension.HTTP.svg)](https://www.nuget.org/packages/FC.Extension.HTTP/)

- [RabbitMQ](https://www.nuget.org/packages/FC.Extension.RabbitMQ/,"RabbitMQ")
	* [![Version](https://img.shields.io/nuget/v/FC.Extension.RabbitMQ.svg)](https://www.nuget.org/packages/FC.Extension.RabbitMQ/)
[![Downloads](https://img.shields.io/nuget/dt/FC.Extension.RabbitMQ.svg)](https://www.nuget.org/packages/FC.Extension.RabbitMQ/)

- [Office](https://www.nuget.org/packages/FC.Extension.Office/,"Office")
	* [![Version](https://img.shields.io/nuget/v/FC.Extension.Office.svg)](https://www.nuget.org/packages/FC.Extension.Office/)
[![Downloads](https://img.shields.io/nuget/dt/FC.Extension.Office.svg)](https://www.nuget.org/packages/FC.Extension.Office/)

- [SQL](https://www.nuget.org/packages/FC.Extension.SQL/,"SQL")
	* [![Version](https://img.shields.io/nuget/v/FC.Extension.SQL.svg)](https://www.nuget.org/packages/FC.Extension.SQL/)
[![Downloads](https://img.shields.io/nuget/dt/FC.Extension.SQL.svg)](https://www.nuget.org/packages/FFC.Extension.SQL/)

>Complete