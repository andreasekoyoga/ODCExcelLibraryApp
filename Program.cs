using OutSystems.ExternalLib.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Filter;
using System.Data;
using Microsoft.Data.Sqlite;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Drawing;

class ExcelTest {


    static Random rd = new Random();
    internal static string CreateString(int stringLength)
    {
    const string allowedChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz0123456789!@$?_-";
    char[] chars = new char[stringLength];

    for (int i = 0; i < stringLength; i++)
    {
        chars[i] = allowedChars[rd.Next(0, allowedChars.Length)];
    }

    return new string(chars);
    }

    public void CreateSheet() {

        Worksheet[] ws = new Worksheet[2];
        ws[0].Name = "Sheet 2";
        ws[0].ColorHex = "#d10000";
        ws[1].Name = "Sheet 3";

        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(ws);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/TestExcelSheet.xlsx", excelFile);
    }

    public void AddSheet(string newSheet) {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("excelLogo.png");
        byte[] excelFile = ex.Worksheet_Add(tmpExcelName, newSheet);
        File.WriteAllBytes($"{Environment.CurrentDirectory}/excelLogo2.png", excelFile);
    }

    public void InsertImage() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(new Worksheet[0]);
        
        byte[] img1 = System.IO.File.ReadAllBytes("ExcelTest/img1.jpeg");

        excelFile = ex.Image_Insert(excelBinary: excelFile, imageFile: img1, cellName: "B1");
        
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestExcelImage.xlsx", excelFile);
    }

    public void InsertImageByName() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/TmpExcelImageNameR7.xlsx");
        
        byte[] img3 = System.IO.File.ReadAllBytes("img3.jpeg");
        byte[] excelFile = ex.Image_Insert(excelBinary: tmpExcelName, imageFile: img3, cellName: "Image3");
        
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestExcelImageName.xlsx", excelFile);
    }

    public void ExcelBorderFormat() {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet 1");

            //worksheet.Cells["B2:G10"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            worksheet.Cells["B2:G10"].Style.Border.Top.Style = ExcelBorderStyle.Double;
            worksheet.Cells["B2:G10"].Style.Border.Bottom.Style = ExcelBorderStyle.Double;
            worksheet.Cells["B2:G10"].Style.Border.Left.Style = ExcelBorderStyle.Double;
            worksheet.Cells["B2:G10"].Style.Border.Right.Style = ExcelBorderStyle.Double;


            byte[] excelFile = package.GetAsByteArray();
            File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestExcelBorder.xlsx", excelFile);
        }
    }

    public void BorderFormat() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(new Worksheet[0]);

        RangeBorderFormat[] rangeBorderFormats = new RangeBorderFormat[3];

        BorderStyleFormat borderStyleFormatThin = new BorderStyleFormat();
        borderStyleFormatThin.BorderStyle = "Thin";
        borderStyleFormatThin.BorderColorHex = "#830047";
        borderStyleFormatThin.IsRound = true;
        borderStyleFormatThin.IsTop = true;
        borderStyleFormatThin.IsBottom = true;
        borderStyleFormatThin.IsLeft = true;
        borderStyleFormatThin.IsRight = true;

        BorderStyleFormat borderStyleFormatThick = new BorderStyleFormat();
        borderStyleFormatThick.BorderStyle = "Thick";
        borderStyleFormatThick.BorderColorHex = "#6600d9";
        borderStyleFormatThick.IsTop = true;

        BorderStyleFormat borderStyleFormatThinLeft = new BorderStyleFormat();
        borderStyleFormatThinLeft.BorderStyle = "Thin";
        borderStyleFormatThinLeft.BorderColorHex = "#38761d";
        borderStyleFormatThinLeft.IsLeft = true;


        rangeBorderFormats[0].borderStyleFormat = borderStyleFormatThin;
        rangeBorderFormats[0].CellName = "B2:G10";
      
        rangeBorderFormats[1].borderStyleFormat = borderStyleFormatThick;
        rangeBorderFormats[1].CellName = "I6:N6";

        rangeBorderFormats[2].borderStyleFormat = borderStyleFormatThinLeft;
        rangeBorderFormats[2].CellName = "B30";

        excelFile = ex.Range_BorderFormat(excelBinary: excelFile, rangeBorderFormats: rangeBorderFormats);
        
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Resuklt/TestExcelBorder.xlsx", excelFile);
    }

    public void Filter() {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("TestFilter.xlsx");
        using (var package = ex.Excel_Open(tmpExcelName))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            ExcelRange excelRange = worksheet.Cells["A:F"];
            excelRange.AutoFilter = false;

            // ExcelValueFilterColumn colCompany = worksheet.AutoFilter.Columns.AddValueFilterColumn(3);
            // colCompany.Filters.Add("Jon Sullivan");
            // worksheet.AutoFilter.ApplyFilter();

            tmpExcelName = package.GetAsByteArray();
        }
        File.WriteAllBytes($"{Environment.CurrentDirectory}/TestFilter.xlsx", tmpExcelName);
    }

    public void LineChartAsync() {
        string connectionStr = "Data Source=EPPlusSample.sqlite;";
        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(new Worksheet[0]);
        using (var package = ex.Excel_Open(excelFile))
        {
            ExcelWorksheet ws1 = package.Workbook.Worksheets[0];

            ExcelRangeBase range;
            using (var sqlConn = new SqliteConnection(connectionStr))
            {
                sqlConn.Open();
                using (var sqlCmd = new SqliteCommand("select orderdate as OrderDate, SUM(ordervalue) as OrderValue, SUM(tax) As Tax,SUM(freight) As Freight from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId Where Currency='USD' group by OrderDate ORDER BY OrderDate desc limit 15", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = ws1.Cells["A1"].LoadFromDataReader(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 0, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd";
                    }
                    //Set the numberformat
                }
            }

            ExcelWorksheet ws = package.Workbook.Worksheets.Add("LineCharts");

            //Add a line chart
            var chart = ws.Drawings.AddLineChart("LineChartWithDroplines", eLineChartType.Line);
            var serie = chart.Series.Add(ws1.Cells[2, 2, 16, 2], ws1.Cells[2, 1, 16, 1]);
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With Droplines";
            chart.AddDropLines();
            chart.DropLine.Border.Width = 2;
            //Set style 12
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle12);

            //Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithErrorBars", eLineChartType.Line);
            serie = chart.Series.Add(ws1.Cells[2, 2, 16, 2], ws1.Cells[2, 1, 16, 1]);
            serie.Header = "Order Value";
            chart.SetPosition(21, 0, 6, 0);
            chart.SetSize(1200, 400);   //Make this chart wider to make room for the datatable.
            chart.Title.Text = "Line Chart With Error Bars";
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.PlotArea.CreateDataTable();

            //Set style 2
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle2);

            //Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithUpDownBars", eLineChartType.Line);
            var serie1 = chart.Series.Add(ws1.Cells[2, 2, 16, 2], ws1.Cells[2, 1, 16, 1]);
            serie1.Header = "Order Value";
            var serie2 = chart.Series.Add(ws1.Cells[2, 3, 16, 3], ws1.Cells[2, 1, 16, 1]);
            serie2.Header = "Tax";
            var serie3 = chart.Series.Add(ws1.Cells[2, 4, 16, 4], ws1.Cells[2, 1, 16, 1]);
            serie3.Header = "Freight";
            chart.SetPosition(42, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With Up/Down Bars";
            chart.AddUpDownBars(true, true);

            //Set style 10, Note: As this is a line chart with multiple series, we use the enum for multiple series. Charts with multiple series usually has a subset of of the chart styles in Excel.
            //Another option to set the style is to use the Excel Style number, in this case 236: chart.StyleManager.SetChartStyle(236)
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.LineChartStyle9);
            range.AutoFitColumns(0);


            //Add a line chart with high/low Bars
            chart = ws.Drawings.AddLineChart("LineChartWithHighLowLines", eLineChartType.Line);
            serie1 = chart.Series.Add(ws1.Cells[2, 2, 26, 2], ws1.Cells[2, 1, 26, 1]);
            serie1.Header = "Order Value";
            serie2 = chart.Series.Add(ws1.Cells[2, 3, 26, 3], ws1.Cells[2, 1, 26, 1]);
            serie2.Header = "Tax";
            serie3 = chart.Series.Add(ws1.Cells[2, 4, 26, 4], ws1.Cells[2, 1, 26, 1]);
            serie3.Header = "Freight";
            chart.SetPosition(63, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With High/Low Lines";
            chart.AddHighLowLines();


            //Add a line chart
            chart = ws.Drawings.AddLineChart("LineChart", eLineChartType.Line);
            serie = chart.Series.Add(ws1.Cells[2, 2, 16, 2], ws1.Cells[2, 1, 16, 1]);
            serie.Header = "Order Value";
            chart.SetPosition(90, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart";
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle2);


            //Set the style using the Excel ChartStyle number. The chart style must exist in the ExcelChartStyleManager.StyleLibrary[]. 
            //Styles can be added and removed from this library. By default it is loaded with the styles for EPPlus supported chart types.
            chart.StyleManager.SetChartStyle(237);
            range.AutoFitColumns(0);            

            excelFile = package.GetAsByteArray();
        }

        File.WriteAllBytes($"{Environment.CurrentDirectory}/TestChart.xlsx", excelFile);    
    }

    public void InsertRow() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/TableDemo.xlsx");
        tmpExcelName = ex.Row_Insert(tmpExcelName, 1, 1, 30);
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestTableDemo.xlsx", tmpExcelName);
    }
    public void InsertColumn() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/TableDemo2.xlsx");
        tmpExcelName = ex.Column_Insert(tmpExcelName, 3, 3, 100, true);
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestTableDemo2.xlsx", tmpExcelName);
    }

    public void FormatNumber() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/RangeFormatNumber.xlsx");

        RangeFormat rangeFormat = new RangeFormat();
        rangeFormat.CellName = "F3:F47";
        
        CellFormat cellFormat = new CellFormat();
        cellFormat.CellType = "Number";
        cellFormat.CellTypeFormat = "#,##0";

        rangeFormat.CellFormat = cellFormat;


        RangeFormat rangeFormatHeader = new RangeFormat();
        rangeFormatHeader.CellName = "A2:F2";
        
        FontStyleFormat fontStyleFormat = new FontStyleFormat();
        fontStyleFormat.IsBold = true;
        fontStyleFormat.HorizontalAlignment = "Center";
        fontStyleFormat.FontColorHex = "#ffffff";

        CellFormat cellFormatHeader = new CellFormat();
        cellFormatHeader.FontStyleFormat = fontStyleFormat;
        cellFormatHeader.BackgroundColorHex = "#444444";

        rangeFormatHeader.CellFormat = cellFormatHeader;

        RangeFormat[] rangeFormats = new RangeFormat[2];
        rangeFormats[0] = rangeFormat;
        rangeFormats[1] = rangeFormatHeader;
        tmpExcelName = ex.Range_Format(tmpExcelName, rangeFormats);

        BorderStyleFormat borderStyleFormat = new BorderStyleFormat();
        borderStyleFormat.BorderStyle = "Thin";
        borderStyleFormat.IsBottom = true;
        borderStyleFormat.IsLeft = true;
        borderStyleFormat.IsRight = true;
        borderStyleFormat.IsTop = true;
        borderStyleFormat.IsRound = true;

        RangeBorderFormat rangeBorderFormat = new RangeBorderFormat();
        rangeBorderFormat.borderStyleFormat = borderStyleFormat;
        rangeBorderFormat.CellName = "A2:F47";

        RangeBorderFormat[] rangeBorderFormats = new RangeBorderFormat[1];
        rangeBorderFormats[0] = rangeBorderFormat;

        tmpExcelName = ex.Range_BorderFormat(tmpExcelName, rangeBorderFormats);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestRangeFormatNumber.xlsx", tmpExcelName);
    }

    public void NumberFormat() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/TableConvert.xlsx");
        RangeFormat rangeFormat = new RangeFormat();
        rangeFormat.CellName = "F2:F46";

        CellFormat cellFormat = new CellFormat();
        cellFormat.CellType = "Number";
        cellFormat.CellTypeFormat = "#,##0";

        rangeFormat.CellFormat = cellFormat;

        RangeFormat[] rangeFormats = new RangeFormat[1];
        rangeFormats[0] = rangeFormat;
        tmpExcelName = ex.Range_Format(tmpExcelName, rangeFormats);

        using (var package = ex.Excel_Open(tmpExcelName))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            
            ExcelRange excelRange = worksheet.Cells["F4"];
            ExcelRange excelRange2 = worksheet.Cells["F5"];
        }


        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestTableConvert.xlsx", tmpExcelName);
    }

    public void LoadData() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(new Worksheet[0]);
        
        var myJsonString = File.ReadAllText("JSON/LeadData.json");

        DataWriteJSON dataWriteJSON = new DataWriteJSON();
        dataWriteJSON.JSONString = myJsonString;
        dataWriteJSON.CellName = "A1";

        DataWriteJSON[] dataWriteJSONs = new DataWriteJSON[1];
        dataWriteJSONs[0] = dataWriteJSON;

        excelFile = ex.Data_WriteJSON(excelFile, dataWriteJSONs);
        
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestLoadData.xlsx", excelFile);

    }

    public void LoadDataMultiSheet() {
        ExcelLibrary ex = new ExcelLibrary();

        Worksheet[] ws = new Worksheet[2];
        ws[0].Name = "Sheet 1";
        ws[1].Name = "Sheet 2";

        byte[] excelFile = ex.Workbook_Create(ws);
        
        var myJsonString1 = File.ReadAllText("JSON/LeadData1.json");

        DataWriteJSON dataWriteJSON1 = new DataWriteJSON();
        dataWriteJSON1.JSONString = myJsonString1;
        dataWriteJSON1.CellName = "A1";
        dataWriteJSON1.SheetName = "Sheet 1";

        var myJsonString2 = File.ReadAllText("JSON/LeadData2.json");

        DataWriteJSON dataWriteJSON2 = new DataWriteJSON();
        dataWriteJSON2.JSONString = myJsonString2;
        dataWriteJSON2.CellName = "A1";
        dataWriteJSON2.SheetName = "Sheet 2";
        dataWriteJSON2.IsShowHeader = true;

        DataWriteJSON[] dataWriteJSONs = new DataWriteJSON[2];
        dataWriteJSONs[0] = dataWriteJSON1;
        dataWriteJSONs[1] = dataWriteJSON2;

        excelFile = ex.Data_WriteJSON(excelFile, dataWriteJSONs);
        
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestLoadData.xlsx", excelFile);

    }

    public void Formula() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/FormulaDemo.xlsx");
        
        CellFormat cellFormat = new CellFormat();
        cellFormat.CellTypeFormat = "#,##0";
        cellFormat.CellType = "Formula";

        CellWrite cellWrite = new CellWrite();
        cellWrite.CellName = "F2";
        cellWrite.CellValue = "=D2*E2";
        cellWrite.CellFormat = cellFormat;

        CellWrite[] cellWrites = new CellWrite[1];
        cellWrites[0] = cellWrite;  

        tmpExcelName = ex.Cell_Write(tmpExcelName, cellWrites);

        CellCopy cellCopy = new CellCopy();
        cellCopy.SourceCellName = "F2";
        cellCopy.DestinationCellName = "F3:F46";

        tmpExcelName = ex.Cell_Copy(tmpExcelName, cellCopy);

        // using (var package = ex.Excel_Open(tmpExcelName))
        // {
        //     ExcelWorksheet worksheet = package.Workbook.Worksheets[0];            
        //     ExcelRange excelRange = worksheet.Cells["F2"];


        //     excelRange.Copy(worksheet.Cells["F3:F30"]);

        //     tmpExcelName = package.GetAsByteArray();
        // }

        
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestFormulaDemo.xlsx", tmpExcelName);
    }

    public void formulaValue2() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/DemoExcelV.xlsx");
        
        CellFormat cellFormat = new CellFormat();
        cellFormat.CellTypeFormat = "#,##0";
        cellFormat.CellType = "Formula";

        CellWrite cellWrite = new CellWrite();
        cellWrite.CellName = "C2";
        cellWrite.CellValue = "=IF(A2=\"\",0,VLOOKUP(A2,'Products'!$A$2:$D$11,4,FALSE))";
        cellWrite.CellFormat = cellFormat;

        CellWrite[] cellWrites = new CellWrite[1];
        cellWrites[0] = cellWrite;  

        //CellCopy cellCopy = new CellCopy();
        tmpExcelName = ex.Cell_Write(tmpExcelName, cellWrites);

        // cellCopy.SourceCellName = "A2:D2";
        // cellCopy.DestinationCellName = "A3:A10";

        // tmpExcelName = ex.Cell_Copy(tmpExcelName, cellCopy);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestFormulaDemo2.xlsx", tmpExcelName);
    }


    public void FindCell() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/FormulaDemo.xlsx");

        CellFindResult[] cellFindResults = ex.Cell_FindByValue(tmpExcelName, "Jawbone", true);

        foreach(CellFindResult cellFindResult in cellFindResults) {
            Console.WriteLine(cellFindResult.CellName + " with Value: " + cellFindResult.CellValue);
        }
    }

    public void AllWorksheets(byte[] tmpExcelName) {
        ExcelLibrary ex = new ExcelLibrary();
        Worksheet[] worksheets = ex.Workbook_GetWorksheet(tmpExcelName);
        foreach(Worksheet worksheet in worksheets) {
            Console.WriteLine("Index: " + worksheet.Index + " with Name: " + worksheet.Name);
        }
        Console.WriteLine("============================");
    }

    public void WorksheetTest() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/NewExcelDemo.xlsx");
        AllWorksheets(tmpExcelName);
        tmpExcelName = ex.Worksheet_Hide_Show(tmpExcelName, 3);
        AllWorksheets(tmpExcelName);
        tmpExcelName = ex.Worksheet_Rename(tmpExcelName, "New Name S3", 2);
        AllWorksheets(tmpExcelName);
        tmpExcelName = ex.Worksheet_Delete(tmpExcelName, 4);
        AllWorksheets(tmpExcelName);
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestNewExcelDemo.xlsx", tmpExcelName);
    }

    public void ExcelProp() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/NewExcelDemo.xlsx");

        KeyValue[] keyValues = new KeyValue[2];
        keyValues[0].Key = "AppVersion";
        keyValues[0].Value = "1.2.3";
        keyValues[1].Key = "Framework";
        keyValues[1].Value = "EPPlus 6.2.2";

        WorkbookProperties workbookProperties = new WorkbookProperties();
        workbookProperties.Author = "Andreas";
        workbookProperties.Company = "OutSystems";
        workbookProperties.Title = "Solution Architect";
        workbookProperties.Comments = "Test Document Properties";
        workbookProperties.KeyValues = keyValues;

        tmpExcelName = ex.Workbook_SetProperties(tmpExcelName, workbookProperties);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestNewExcelDemo.xlsx", tmpExcelName);
    }

    public void CellRead() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("TestExcel.xlsx");

        Console.WriteLine(ex.Cell_Read(tmpExcelName, 0, 0, "", "Sheet 5"));
    }

    public void ListDataValidation() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(new Worksheet[0]);

        DataValidationConfig dataValidationConfig = new DataValidationConfig();
        dataValidationConfig.IsShowInputMessage = true;
        dataValidationConfig.InputMessage = "Your Input!";
        dataValidationConfig.ErrorStyle = "Stop";
        dataValidationConfig.ErrorMessage = "Choose!";


        string[] items = new string[5];
        items[0] = "Text A";
        items[1] = "Text B";
        items[2] = "Text C";
        items[3] = "Text D";
        items[4] = "Text E";

        CellDataValidation cellDataValidation = new CellDataValidation();
        // CellRange cellRange = new CellRange();
        // cellRange.StartCellRow = 1;
        // cellRange.StartCellColumn = 1;
        // cellRange.EndCellRow = 1;
        // cellRange.EndCellColumn = 1;
        // cellDataValidation.CellRange = cellRange;

        cellDataValidation.CellName = "A:A";

        DataValidationListItem dataValidationListItem = new DataValidationListItem();
        dataValidationListItem.ItemList = items;
        dataValidationListItem.dataValidationConfig = dataValidationConfig;


        excelFile = ex.Data_Validation_List(excelFile, cellDataValidation, dataValidationListItem);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestListValidationDemo.xlsx", excelFile); 
    }

    public void ListDataValidation2() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/NewExcelDataValidationDemo.xlsx");

        DataValidationConfig dataValidationConfig = new DataValidationConfig();
        dataValidationConfig.IsShowInputMessage = true;
        dataValidationConfig.InputTitle = "Input Title";
        dataValidationConfig.InputMessage = "Your Input!";
        dataValidationConfig.ErrorStyle = "Stop";
        dataValidationConfig.ErrorTitle = "Error Title!";
        dataValidationConfig.ErrorMessage = "Choose!";

        CellDataValidation cellDataValidation = new CellDataValidation();
        cellDataValidation.CellName = "C:C";
        cellDataValidation.SheetName = "Sheet 2";

        DataValidationListItem dataValidationListItem = new DataValidationListItem();
        dataValidationListItem.ItemFormula = "'Sheet 1'!$A$2:$A$6";
        dataValidationListItem.dataValidationConfig = dataValidationConfig;

        tmpExcelName = ex.Data_Validation_List(tmpExcelName, cellDataValidation, dataValidationListItem);

        CellWrite cellWrite = new CellWrite();
        cellWrite.CellName = "C4";
        cellWrite.CellValue = "A6";
        cellWrite.SheetName = "Sheet 2";

        CellWrite[] cellWrites = new CellWrite[1];
        cellWrites[0] = cellWrite; 
        
        tmpExcelName = ex.Cell_Write(tmpExcelName, cellWrites);


        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestListValidationDemo2.xlsx", tmpExcelName); 
    }    

    public void ListDataValidation3() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/TestDataValidation.xlsx");

        DataValidationConfig dataValidationConfig = new DataValidationConfig();
        dataValidationConfig.IsShowInputMessage = true;
        dataValidationConfig.InputTitle = "Input Title";
        dataValidationConfig.InputMessage = "Your Input!";
        dataValidationConfig.ErrorStyle = "Stop";
        dataValidationConfig.ErrorTitle = "Error Title!";
        dataValidationConfig.ErrorMessage = "Choose!";

        CellRange cellRange = new CellRange();
        cellRange.StartCellRow = 5;
        cellRange.StartCellColumn = 3;
        cellRange.EndCellRow = 105;
        cellRange.EndCellColumn = 3;

        CellDataValidation cellDataValidation = new CellDataValidation();
        cellDataValidation.CellRange = cellRange;
        cellDataValidation.SheetName = "Booking";

        DataValidationListItem dataValidationListItem = new DataValidationListItem();
        dataValidationListItem.ItemFormula = "Stations!$A$1:$A$189";
        dataValidationListItem.dataValidationConfig = dataValidationConfig;

        tmpExcelName = ex.Data_Validation_List(tmpExcelName, cellDataValidation, dataValidationListItem);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestDataValidationResult.xlsx", tmpExcelName); 
    }    

    public byte[] CreateExcel() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] excelFile = ex.Workbook_Create(new Worksheet[0]);
        return excelFile;        
    }

    public byte[] IntDataValidation(byte[] excelFile) {
        ExcelLibrary ex = new ExcelLibrary();

        DataValidationConfig dataValidationConfigBWT = new DataValidationConfig();
        dataValidationConfigBWT.IsShowInputMessage = true;
        dataValidationConfigBWT.InputTitle = "Input Title";
        dataValidationConfigBWT.InputMessage = "Your Input!";
        dataValidationConfigBWT.ErrorStyle = "Stop";
        dataValidationConfigBWT.ErrorTitle = "Error Title!";
        dataValidationConfigBWT.ErrorMessage = "Choose!";
        dataValidationConfigBWT.ValidationOperator = "between";

        CellDataValidation cellDataValidationBWT = new CellDataValidation();
        cellDataValidationBWT.CellName = "C2";

        DataValidation dataValidationBWT = new DataValidation();
        dataValidationBWT.dataValidationConfig = dataValidationConfigBWT;
        dataValidationBWT.FormulaValue1 = "10";
        dataValidationBWT.FormulaValue2 = "20";


        DataValidationConfig dataValidationConfigGRT = new DataValidationConfig();
        dataValidationConfigGRT.IsShowInputMessage = true;
        dataValidationConfigGRT.InputTitle = "Input Title";
        dataValidationConfigGRT.InputMessage = "Your Input!";
        dataValidationConfigGRT.ErrorStyle = "Stop";
        dataValidationConfigGRT.ErrorTitle = "Error Title!";
        dataValidationConfigGRT.ErrorMessage = "Choose!";
        dataValidationConfigGRT.ValidationOperator = "GreaterThan";

        CellDataValidation cellDataValidationGRT = new CellDataValidation();
        cellDataValidationGRT.CellName = "C3";

        DataValidation dataValidationGRT = new DataValidation();
        dataValidationGRT.dataValidationConfig = dataValidationConfigGRT;
        dataValidationGRT.FormulaValue1 = "50";

        excelFile = ex.Data_Validation_Integer(excelFile, cellDataValidationBWT, dataValidationBWT);
        excelFile = ex.Data_Validation_Integer(excelFile, cellDataValidationGRT, dataValidationGRT);

        return excelFile;
    }    

    public byte[] DecDataValidation(byte[] excelFile) {
        ExcelLibrary ex = new ExcelLibrary();

        DataValidationConfig dataValidationConfigBWT = new DataValidationConfig();
        dataValidationConfigBWT.IsShowInputMessage = true;
        dataValidationConfigBWT.InputTitle = "Input Title";
        dataValidationConfigBWT.InputMessage = "Your Input!";
        dataValidationConfigBWT.ErrorStyle = "Stop";
        dataValidationConfigBWT.ErrorTitle = "Error Title!";
        dataValidationConfigBWT.ErrorMessage = "Choose!";
        dataValidationConfigBWT.ValidationOperator = "between";

        CellDataValidation cellDataValidationBWT = new CellDataValidation();
        cellDataValidationBWT.CellName = "D2";

        DataValidation dataValidationBWT = new DataValidation();
        dataValidationBWT.dataValidationConfig = dataValidationConfigBWT;
        dataValidationBWT.FormulaValue1 = "1.3";
        dataValidationBWT.FormulaValue2 = "1.8";


        DataValidationConfig dataValidationConfigGRT = new DataValidationConfig();
        dataValidationConfigGRT.IsShowInputMessage = true;
        dataValidationConfigGRT.InputTitle = "Input Title";
        dataValidationConfigGRT.InputMessage = "Your Input!";
        dataValidationConfigGRT.ErrorStyle = "Stop";
        dataValidationConfigGRT.ErrorTitle = "Error Title!";
        dataValidationConfigGRT.ErrorMessage = "Choose!";
        dataValidationConfigGRT.ValidationOperator = "lessthanequal";

        CellDataValidation cellDataValidationGRT = new CellDataValidation();
        cellDataValidationGRT.CellName = "D3";

        DataValidation dataValidationGRT = new DataValidation();
        dataValidationGRT.dataValidationConfig = dataValidationConfigGRT;
        dataValidationGRT.FormulaValue1 = "2.3";

        excelFile = ex.Data_Validation_Decimal(excelFile, cellDataValidationBWT, dataValidationBWT);
        excelFile = ex.Data_Validation_Decimal(excelFile, cellDataValidationGRT, dataValidationGRT);

        return excelFile;
    }    

    public byte[] DTDataValidation(byte[] excelFile) {
        ExcelLibrary ex = new ExcelLibrary();

        DateTime D1 = DateTime.Now;
        Console.WriteLine(D1.ToString());
        DateTime D2 = D1.AddDays(2);
        Console.WriteLine(D2.ToString());
        DateTime D3 = D1.AddDays(3);
        Console.WriteLine(D3.ToString());

        DataValidationConfig dataValidationConfigBWT = new DataValidationConfig();
        dataValidationConfigBWT.IsShowInputMessage = true;
        dataValidationConfigBWT.InputTitle = "Input Title";
        dataValidationConfigBWT.InputMessage = "Your Input!";
        dataValidationConfigBWT.ErrorStyle = "Stop";
        dataValidationConfigBWT.ErrorTitle = "Error Title!";
        dataValidationConfigBWT.ErrorMessage = "Choose!";
        dataValidationConfigBWT.ValidationOperator = "between";

        CellDataValidation cellDataValidationBWT = new CellDataValidation();
        cellDataValidationBWT.CellName = "E2";

        DataValidation dataValidationBWT = new DataValidation();
        dataValidationBWT.dataValidationConfig = dataValidationConfigBWT;
        dataValidationBWT.FormulaValue1 = D1.ToString();
        dataValidationBWT.FormulaValue2 = D2.ToString();


        DataValidationConfig dataValidationConfigGRT = new DataValidationConfig();
        dataValidationConfigGRT.IsShowInputMessage = true;
        dataValidationConfigGRT.InputTitle = "Input Title";
        dataValidationConfigGRT.InputMessage = "Your Input!";
        dataValidationConfigGRT.ErrorStyle = "Stop";
        dataValidationConfigGRT.ErrorTitle = "Error Title!";
        dataValidationConfigGRT.ErrorMessage = "Choose!";
        dataValidationConfigGRT.ValidationOperator = "GreaterThan";

        CellDataValidation cellDataValidationGRT = new CellDataValidation();
        cellDataValidationGRT.CellName = "E3";

        DataValidation dataValidationGRT = new DataValidation();
        dataValidationGRT.dataValidationConfig = dataValidationConfigGRT;
        dataValidationGRT.FormulaValue1 = D3.ToString();

        excelFile = ex.Data_Validation_DateTime(excelFile, cellDataValidationBWT, dataValidationBWT);
        excelFile = ex.Data_Validation_DateTime(excelFile, cellDataValidationGRT, dataValidationGRT);

        return excelFile;
    }    

    public void TestDataValidation() {
        byte[] excelFile = CreateExcel();

        excelFile = IntDataValidation(excelFile);
        excelFile = DecDataValidation(excelFile);
        excelFile = DTDataValidation(excelFile);

        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestListValidationDemo3.xlsx", excelFile); 
    }    

    public void CellWrites() {
        byte[] excelFile = CreateExcel();
        ExcelLibrary ex = new ExcelLibrary();

        CellFormat cellFormat = new CellFormat();
        cellFormat.IsAutoFitColumn = true;
        cellFormat.BackgroundColorHex = "#fff3ae";

        CellWrite[] cellWrites= new CellWrite[20];

        Cell cell = new Cell();
        cell.CellColumn = 1;
        cell.CellRow = 1;

        for(int i = 0; i < 10; i++) {
            CellWrite cellWrite= new CellWrite();
            cellWrite.CellFormat = cellFormat;
            cellWrite.CellValue = CreateString(rd.Next(10,30));
            cellWrite.Cell = cell;

            cellWrites[i] = cellWrite;

            cell.CellRow = cell.CellRow + 1;

        } 

        cell.CellColumn = 2;
        cell.CellRow = 1;

        for(int i = 10; i < 20; i++) {
            CellWrite cellWrite= new CellWrite();
            cellWrite.CellValue = CreateString(rd.Next(10,30));
            cellWrite.Cell = cell;

            cellWrites[i] = cellWrite;

            cell.CellRow = cell.CellRow + 1;

        }

        excelFile = ex.Cell_Write(excelFile, cellWrites);
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestCellWriteDemo.xlsx", excelFile); 

    }

    public void ReadVasialis() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/MW-302 WG WYKAZU 22309.xlsx");

        Console.WriteLine(ex.Cell_Read(tmpExcelName, 4,1));

    }


    public void CellWritesRich() {
        byte[] excelFile = CreateExcel();
        ExcelLibrary ex = new ExcelLibrary();

        CellWriteRichText[] cellWriteRichTexts = new CellWriteRichText[1];
        RichTextFormatText[] richTextFormatTexts = new RichTextFormatText[3];

        richTextFormatTexts[0].CellValue = "This is";
        richTextFormatTexts[0].FontColorHex = "#4455dd";
        richTextFormatTexts[0].FontName = "Courier New";

        richTextFormatTexts[1].CellValue = " Richtext";
        richTextFormatTexts[1].FontColorHex = "#990000";
        richTextFormatTexts[1].IsBold = true;
        richTextFormatTexts[1].FontSize = 16;

        richTextFormatTexts[2].CellValue = " Format Text!";
        //richTextFormatTexts[2].FontColorHex = "#527826";
        richTextFormatTexts[2].IsItalic = true;

        cellWriteRichTexts[0].CellName = "B2";
        cellWriteRichTexts[0].RichTextFormatTexts = richTextFormatTexts;
        cellWriteRichTexts[0].IsAutoFitColumn = true;

        excelFile = ex.Cell_Write_RichText(excelFile, cellWriteRichTexts);
        File.WriteAllBytes($"{Environment.CurrentDirectory}/Result/TestCellWriteRichText.xlsx", excelFile);     
    }

    public void CellRange_ReadTest() {
        ExcelLibrary ex = new ExcelLibrary();
        byte[] tmpExcelName = System.IO.File.ReadAllBytes("ExcelTest/ReadExcelDemo.xlsx");
        RangeCellRead[] cellRange_Read = new RangeCellRead[2];
        CellRange cellRange1 = new CellRange();

        cellRange1.StartCellRow = 5;
        cellRange1.StartCellColumn = 3;
        cellRange1.EndCellRow = 6;
        cellRange1.EndCellColumn = 7;

        cellRange_Read[0].CellRange = cellRange1;

        CellRange cellRange2 = new CellRange();

        cellRange2.StartCellRow = 7;
        cellRange2.StartCellColumn = 3;
        cellRange2.EndCellRow = 7;
        cellRange2.EndCellColumn = 7;

        cellRange_Read[1].CellRange = cellRange2;


        RangeCellValue[] cellValues = ex.Range_CellRead(tmpExcelName, cellRange_Read);

        foreach (RangeCellValue cellValue in cellValues) {
            Console.WriteLine("Row: " + cellValue.CellRow + " Col: " + cellValue.CellColumn +" Cell: " + cellValue.CellName + " Value: " + cellValue.Value);
        }


    }

}


class Program
{
    static void Main(string[] args)
    {
        ExcelTest excelTest = new ExcelTest();
        //excelTest.ListDataValidation2();
        //excelTest.TestDataValidation();
        //excelTest.CellWrites();
        //excelTest.ListDataValidation3();
        //excelTest.ReadVasialis();
        //excelTest.CellWritesRich();
        //excelTest.CellRange_ReadTest();
        excelTest.formulaValue2();
    }
}
