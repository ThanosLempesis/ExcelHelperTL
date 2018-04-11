using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Serilog;

namespace IOExcelHelpers
{
    public static class ExcelHelper
    {


        private static Dictionary<string, List<string>> _excelTemplateColumnCellReferences;
        
        private static int m_OriginatorId;
        private static string m_ConnStr;
        private static List<string> _errorsOnInvalidCoulumns;        



        public static void ReadTemplateIntoTable(DataSet dataSet, string fileName, List<string> supportedSheetNames, bool flag)
        {
            var logger = new LoggerConfiguration()
                                .WriteTo.Console()
                                .CreateLogger();


            try {

                using (var spredsheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    var workBookRoot = spredsheetDocument.WorkbookPart.Workbook;
                    var sheets = workBookRoot.Sheets;

                    _excelTemplateColumnCellReferences = new Dictionary<string, List<string>>();

                    foreach (var sh in sheets)
                    {
                        var dataTable = new DataTable();

                        var sheet = (Sheet)sh;
                        var relationshipId = sheet.Id.Value;
                        string sheetName = sheet.Name;

                        if (CheckSheetState((Sheet)sh, supportedSheetNames))
                            continue;

                        dataTable.TableName = sheetName;
                        var columnCellReferences = new List<string>();

                        var worksheetPart = spredsheetDocument.WorkbookPart.GetPartById(relationshipId) as WorksheetPart;
                        if (worksheetPart != null)
                        {
                            var workSheet = worksheetPart.Worksheet;
                            var sheetData = workSheet.GetFirstChild<SheetData>();
                            var rows = sheetData.Descendants<Row>();

                            foreach (var openXmlElement in rows.ElementAt(0))
                            {
                                var cell = (Cell)openXmlElement;
                                var cellRef = GetCellRef(cell.CellReference, 1);
                                var columnName = GetCellValue(spredsheetDocument, cell, 0u);
                                dataTable.Columns.Add(NewColumnName(cellRef, (columnName.Replace(" ", "")).ToUpperInvariant(), flag));
                                columnCellReferences.Add(cellRef);
                            }
                        }
                        dataSet.Tables.Add(dataTable);

                        _excelTemplateColumnCellReferences.Add(sheetName, columnCellReferences);
                    }
                }
            }
            catch (System.IO.IOException ex)
            {
                logger.Information("The system encountered the following error{Message}", ex.Message);

            }
           
        }

        public static void ReadDictionaryIntoTable(DataSet dataSet, string fileName, List<string> supportedSheetNames)
        {
            using (var spredsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                var workBookRoot = spredsheetDocument.WorkbookPart.Workbook;
                var sheets = workBookRoot.Sheets;

                var cellFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;

                foreach (var sh in sheets)
                {
                    var dataTable = new DataTable();

                    var sheet = (Sheet)sh;
                    var relationshipId = sheet.Id.Value;
                    string sheetName = sheet.Name;

                    if (CheckSheetState((Sheet)sh, supportedSheetNames))
                        continue;

                    dataTable.TableName = sheetName;
                    dataTable.Columns.Add(Constants.DictionaryColumnName);
                    dataTable.Columns.Add(Constants.DictionaryDataType);

                    var worksheetPart = spredsheetDocument.WorkbookPart.GetPartById(relationshipId) as WorksheetPart;
                    if (worksheetPart != null)
                    {
                        var workSheet = worksheetPart.Worksheet;
                        var sheetData = workSheet.GetFirstChild<SheetData>();
                        var rows = sheetData.Descendants<Row>(); 
                        foreach (var row in rows.Skip(1))
                        {
  
                            var tableRow = dataTable.NewRow();
                            tableRow[0] = GetCellValue(spredsheetDocument, (Cell)row.ElementAt(0), GetNumberFormatIdFromCellStyleIndex(cellFormats, (Cell)row.ElementAt(0))).Replace(" ", "");
                            tableRow[1] = GetCellValue(spredsheetDocument, (Cell)row.ElementAt(1), GetNumberFormatIdFromCellStyleIndex(cellFormats, (Cell)row.ElementAt(1))); ;
                            
                            dataTable.Rows.Add(tableRow);
                        }
                    }
                    dataSet.Tables.Add(dataTable);
                }
            }
        }


        public static List<string> ReadTemplateBasedExcelAsDataTables(DataSet dataSet, string fileName, List<string> supportedSheetNames)
        {
            var isExcelFileStructureValid = true;
            var errorsOnInvalidCoulumns = new List<string>();

            using (var spredsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                var workBookRoot = spredsheetDocument.WorkbookPart.Workbook;
                var sheets = workBookRoot.Sheets;

                var numberingFormats = new List<NumberingFormat>();
                var cellFormats = new CellFormats();

                if ( workBookRoot.WorkbookPart.WorkbookStylesPart != null 
                  && workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet != null
                  && workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null )
                {
                    numberingFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>().ToList();
                    cellFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
                }

                foreach (var sh in sheets)
                {
                    var sheet = (Sheet)sh;
                    var relationshipId = sheet.Id.Value;
                    string sheetName = sheet.Name;

                    if (CheckSheetState((Sheet)sh, supportedSheetNames))
                        continue;

                    var worksheetPart = spredsheetDocument.WorkbookPart.GetPartById(relationshipId) as WorksheetPart;
                    if (worksheetPart != null)
                    {
                        var workSheet = worksheetPart.Worksheet;
                        var sheetData = workSheet.GetFirstChild<SheetData>();
                        var rows = sheetData.Descendants<Row>();             
                        var dt = dataSet.Tables[sheetName];
                        if (dt == null)
                        {
                            errorsOnInvalidCoulumns.Add(string.Format("The sheet {0} is not supported in this template", sheetName));
                            continue;
                        }

                        var columnNames = new List<string>();


                        //TL: I check if the 1st row is valid or now and only then I will enter the loop of the rest rows.
                        if (rows.ElementAt(0).Descendants() == null)
                        {
                            continue;
                        }

                        foreach (var cell in rows.ElementAt(0).Descendants<Cell>())
                        {

                            if (cell.CellReference != null)
                            {
                                columnNames.Add(NewColumnName(GetCellRef(cell.CellReference, 1), (GetCellValue(spredsheetDocument, cell, 0u)).Replace(" ", ""), true));
                            }
                            else
                            { 
                                columnNames.Add(cell.CellValue.Text);

                            }

                            if (!dt.Columns.Contains(columnNames.Last()))
                            {
                                isExcelFileStructureValid = false;
                                errorsOnInvalidCoulumns.Add(sheetName + Constants.Slash + columnNames.Last());
                            }
                        }


                        if (isExcelFileStructureValid)
                        {
                            uint rowIndex = 2;
                            foreach (var row in rows.Skip(1))                                 
                            {

                                if (!row.HasChildren)
                                {
                                    break;
                                }

                                var tmp = row.Descendants<Cell>().ToList();
                                //Add rows to DataTable.
                                if ((tmp.Count >= 0 && tmp[0].CellValue == null))
                                {
                                    break;
                                }

                                dt.Rows.Add();

                                int columnIndex = -1;
                                foreach (var cell in row.Descendants<Cell>())
                                {
                                    if (cell.CellValue == null)
                                        continue;

                                    if (cell.CellReference != null)
                                    {
                                        var cellRef = GetCellRef(cell.CellReference, rowIndex);
                                        columnIndex = columnNames.FindIndex(x => x.StartsWith(cellRef));
                                    }
                                    else
                                    {
                                        columnIndex++;
                                    }

                                    uint numberFormatId = 0u;

                                    if (cellFormats != null && cellFormats.HasChildren)
                                    {
                                        numberFormatId = GetNumberFormatIdFromCellStyleIndex(cellFormats, cell);
                                    }

                                    var formatCode = "";
                                    if (numberingFormats != null && numberingFormats.Count > 0)
                                    {
                                        var matchingNumFormat = numberingFormats.Where(i => i.NumberFormatId.Value == numberFormatId).ToList();
                                        if (matchingNumFormat.Any())
                                            formatCode = matchingNumFormat.First().FormatCode;
                                    }
                                    dt.Rows[dt.Rows.Count - 1][columnIndex] = GetCellValue(spredsheetDocument, cell, numberFormatId, formatCode);
                                }
                                var appIdCell = dt.Rows[dt.Rows.Count - 1][0];
                                if (appIdCell is DBNull)
                                {
                                    dt.Rows[dt.Rows.Count - 1][0] = "";
                                }
                                if (sheetName == Constants.Contacts || sheetName == Constants.Properties || sheetName == Constants.OtherSecurities)
                                {
                                    var idCell = dt.Rows[dt.Rows.Count - 1][1];
                                    if (idCell is DBNull)
                                    {
                                        dt.Rows[dt.Rows.Count - 1][1] = "";
                                    }
                                }

                                rowIndex++;
                            }
                           
                        }
                    }
                }
            }

            return errorsOnInvalidCoulumns;
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        public static List<string> ReadFastLenderLtdTemplateBasedExcelAsDataTables(int originator, string connStr, DataSet dataSet, string fileName, string templateFileFullName,
                                                                            List<string> supportedSheetNames, Dictionary<string, Pair> additionalData)
        {
            m_OriginatorId = originator;
            m_ConnStr = connStr;

            const int extraRowNumber = 2;
            const uint rowFastLenderLtdStartReference = 10;
            const int addressItemsTotal = 3;

            uint discrepancy = 1;          // 1 or 2

            var district = "";
            var town = "";
            var county = "";
            var postcode = "";

            _errorsOnInvalidCoulumns = new List<string>();

            using (var spredsheetDocument = SpreadsheetDocument.Open(templateFileFullName, false))
            {
                var workBookRoot = spredsheetDocument.WorkbookPart.Workbook;
                var sheets = workBookRoot.Sheets;

                var numberingFormats = new List<NumberingFormat>();
                var cellFormats = new CellFormats();

                if (workBookRoot.WorkbookPart.WorkbookStylesPart != null
                  && workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet != null
                  && workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null)
                {
                    numberingFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>().ToList();
                    cellFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
                }

                using (var sprSheetDocumentFastLenderLtd = SpreadsheetDocument.Open(fileName, false))
                {
                    try
                    {
                        var workBookFastLenderLtdRoot = sprSheetDocumentFastLenderLtd.WorkbookPart.Workbook;
                        var sheetsFastLenderLtd = workBookFastLenderLtdRoot.Sheets;

                        var numberingFormatsFastLenderLtd = new List<NumberingFormat>();
                        var cellFormatsFastLenderLtd = new CellFormats();

                        if (workBookFastLenderLtdRoot.WorkbookPart.WorkbookStylesPart != null
                          && workBookFastLenderLtdRoot.WorkbookPart.WorkbookStylesPart.Stylesheet != null
                          && workBookFastLenderLtdRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null)
                        {
                            numberingFormatsFastLenderLtd = workBookFastLenderLtdRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>().ToList();
                            cellFormatsFastLenderLtd = workBookFastLenderLtdRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;
                        }

                        var sheetFastLenderLtd = (Sheet)sheetsFastLenderLtd.ToArray()[0];
                        var relationshipIdFastLenderLtd = sheetFastLenderLtd.Id.Value;
                        //string sheetNameFastLenderLtd = sheetFastLenderLtd.Name;

                        var worksheetPartFastLenderLtd = sprSheetDocumentFastLenderLtd.WorkbookPart.GetPartById(relationshipIdFastLenderLtd) as WorksheetPart;

                        if (worksheetPartFastLenderLtd == null)
                        {
                            throw new InvalidMCContentException();
                        }

                        var workSheetFastLenderLtd = worksheetPartFastLenderLtd.Worksheet;
                        var sheetDataFastLenderLtd = workSheetFastLenderLtd.GetFirstChild<SheetData>();
                        var rowsFastLenderLtd = sheetDataFastLenderLtd.Descendants<Row>().ToArray();   // IEnumerable<Row>                     // DocumentFormat.OpenXml.Spreadsheet.Row

                        var rowsFastLenderLtdCount = rowsFastLenderLtd.Length;

                        var count = 1u;
                        uint rowFastLenderLtdStartNumber = 8;
                        foreach (var row in rowsFastLenderLtd)
                        {
                            if (row.RowIndex == rowFastLenderLtdStartReference)
                            {
                                rowFastLenderLtdStartNumber = count;
                                discrepancy = rowFastLenderLtdStartReference - count + 1;
                                break;
                            }
                            count++;
                        }

                        foreach (var sh in sheets)
                        {
                            var sheet = (Sheet)sh;
                            var relationshipId = sheet.Id.Value;
                            string sheetName = sheet.Name;

                            if (CheckSheetState((Sheet)sh, supportedSheetNames))
                                continue;

                            var worksheetPart = spredsheetDocument.WorkbookPart.GetPartById(relationshipId) as WorksheetPart;
                            if (worksheetPart != null)
                            {
                                var workSheet = worksheetPart.Worksheet;
                                var sheetData = workSheet.GetFirstChild<SheetData>();
                                var rows = sheetData.Descendants<Row>().ToArray();                    // IEnumerable<Row>                     // DocumentFormat.OpenXml.Spreadsheet.Row   

                                var dt = dataSet.Tables[sheetName];
                                if (dt == null)
                                {
                                    _errorsOnInvalidCoulumns.Add(string.Format(Constants.FileRejectedMsg + "The sheet {0} is not supported in this template", sheetName));
                                    continue;
                                }

                                #region loopsForRows

                                Row value2NdRow = rows[extraRowNumber - 1];
                                Row value3RdRow = rows[extraRowNumber];
                                Row value4ThRow = rows[extraRowNumber + 1];

                                var columnNames = new List<string>();

                                //Loop through the Worksheet rows.

                                var isLastRow = false;
                                var isExtraPropertyRow = false;

                                //foreach (var row in rows)                                   // DocumentFormat.OpenXml.Spreadsheet.Row
                                for (uint j = rowFastLenderLtdStartNumber - 1; j <= rowsFastLenderLtdCount - 1; j++)
                                {
                                    if (!rowsFastLenderLtd[j].HasChildren)
                                    {
                                        break;
                                    }

                                    //Use the first row to validate column in the DataTable.
                                    if (j == rowFastLenderLtdStartNumber - 1)
                                    {
                                        #region cellLoopFor1stRow
                                        foreach (var cell in rows[0].Descendants<Cell>())
                                        {
                                            var cellRef = GetCellRef(cell.CellReference, 1);
                                            var columnName = (GetCellValue(spredsheetDocument, cell, 0u)).Replace(" ", "");

                                            columnNames.Add(NewColumnName(cellRef, columnName, true));
                                        }
                                        #endregion cellLoopFor1stRow
                                    }

                                    //Add rows to DataTable.
                                    dt.Rows.Add();

                                    int columnIndex = -1;

                                    if (isExtraPropertyRow && sheetName != Constants.Properties)
                                    {
                                        continue;
                                    }

                                    #region cellLoopForAfter1stRows
                                    foreach (var cell in value2NdRow.Descendants<Cell>())
                                    {
                                        if (cell.CellValue == null)
                                            continue;

                                        if (cell.CellReference != null)
                                        {
                                            var cellRef = GetCellRef(cell.CellReference, extraRowNumber);
                                            columnIndex = columnNames.FindIndex(x => x.StartsWith(cellRef));
                                        }
                                        else
                                        {
                                            columnIndex++;
                                        }

                                        uint numberFormatId = 0u;

                                        if (cellFormats != null && cellFormats.HasChildren)
                                        {
                                            numberFormatId = GetNumberFormatIdFromCellStyleIndex(cellFormats, cell);
                                        }

                                        var formatCode = "";
                                        if (numberingFormats != null && numberingFormats.Count > 0)
                                        {
                                            var matchingNumFormat = numberingFormats.Where(i => i.NumberFormatId.Value == numberFormatId).ToList();
                                            if (matchingNumFormat.Any())
                                                formatCode = matchingNumFormat.First().FormatCode;
                                        }

                                        object realCellValue;

                                        var tmpCellValue = (GetCellValue(spredsheetDocument, cell, numberFormatId, formatCode)).Trim().ToLower();
                                        var newCell = value3RdRow.Descendants<Cell>().ToArray()[columnIndex];
                                        var newCellTempValue = GetCellValue(spredsheetDocument, newCell, numberFormatId, formatCode).Trim();

                                        // Here the 'else if' wins over 'switch' because it is less lines of code after all as well as due to the 'break' of the external loop in each/some cases
                                        //
                                        if (tmpCellValue == "default")
                                        {
                                            realCellValue = newCellTempValue;
                                        }
                                        else if (tmpCellValue == "auto")
                                        {
                                            realCellValue = (j - (rowFastLenderLtdStartNumber - 1 ) + 1).ToString();
                                        }
                                        else // means: all "ref" and "maps"
                                        {
                                            var cellFastLenderLtd = rowsFastLenderLtd[j].Descendants<Cell>().FirstOrDefault(x => x.CellReference == (newCellTempValue + (j + discrepancy).ToString()));

                                            if (cellFastLenderLtd == null)
                                                realCellValue = "";
                                            else
                                            {
                                                var numberFormatIdFastLenderLtd = 0u;

                                                if (cellFormatsFastLenderLtd != null && cellFormatsFastLenderLtd.HasChildren)
                                                {
                                                    numberFormatIdFastLenderLtd = GetNumberFormatIdFromCellStyleIndex(cellFormatsFastLenderLtd, cellFastLenderLtd);
                                                }

                                                var formatCodeFastLenderLtd = "";
                                                if (numberingFormatsFastLenderLtd != null && numberingFormatsFastLenderLtd.Count > 0)
                                                {
                                                    var matchingNumFormat = numberingFormatsFastLenderLtd.Where(i => i.NumberFormatId.Value == numberFormatIdFastLenderLtd).ToList();
                                                    if (matchingNumFormat.Any())
                                                        formatCodeFastLenderLtd = matchingNumFormat.First().FormatCode;
                                                }

                                                if (tmpCellValue == "ref" || tmpCellValue == "map4" || tmpCellValue == "map5")
                                                {
                                                    // If there is no value in FastLenderLtd 'B-next' cell, we assume this is the end of account rows data
                                                    if (cell.CellReference == "A2" && (cellFastLenderLtd.CellValue == null))
                                                    {
                                                        if (sheetName == Constants.Accounts)
                                                            _errorsOnInvalidCoulumns.Add(Constants.FileRejectedMsg + string.Format("The row {0} does not have any value in the 'Primary Loan ID' column", cellFastLenderLtd.CellReference));

                                                        realCellValue = "????";
                                                    }
                                                    else
                                                        realCellValue = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                }
                                                else
                                                {
                                                    realCellValue = "";
                                                    switch (tmpCellValue)
                                                    {
                                                        case "map1":
                                                        {
                                                            var newCell2 = value4ThRow.Descendants<Cell>().ToArray()[columnIndex];
                                                            var newCellTempValue2 = GetCellValue(spredsheetDocument, newCell2, numberFormatId, formatCode);
                                                            var cell2FastLenderLtd = rowsFastLenderLtd[j].Descendants<Cell>().FirstOrDefault(x => x.CellReference == (newCellTempValue2 + (j + discrepancy).ToString()));
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd).Replace(",", "");
                                                            var realValue2 = GetCellValue(sprSheetDocumentFastLenderLtd, cell2FastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd).Replace(",", ""); 

                                                            realCellValue = (Convert.ToDouble(realValue1 == "" ? "0" : realValue1) - Convert.ToDouble(realValue2 == "" ? "0" : realValue2)).ToString(CultureInfo.InvariantCulture);

                                                            break;
                                                        }
                                                        case "map2":
                                                        {
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                            realCellValue = MapFastLenderLtdChargeType(realValue1);
                                                            break;
                                                        }
                                                        case "map3":
                                                        {
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);

                                                            if (realValue1 == "Y")
                                                            {
                                                                isExtraPropertyRow = true;
                                                                realCellValue = realValue1;
                                                                //j++;
                                                            }
                                                            else
                                                            {
                                                                isExtraPropertyRow = false;
                                                                realCellValue = "N";
                                                            }

                                                            break;
                                                        }
                                                        //case "map4":
                                                        //    break;
                                                        //case "map5":
                                                        //    break;
                                                        case "map6":
                                                        {
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                            var celList = new List<string>();
                                                            var cells = celList.ToArray();   // line.Split(',');

                                                            if (cells.Length < addressItemsTotal || cells.Length > (addressItemsTotal + 1))
                                                            {
                                                                break;
                                                            }

                                                            if (cells.Length == 0)
                                                            {
                                                                //errorsOnInvalidCoulumns.Add(string.Format("The address does not have any items in it {0}", realValue1));
                                                                break;
                                                            }

                                                            realCellValue = cells[0].Trim();        // street
                                                            district = "";                          // District
                                                            town = cells[1];                        // Town


                                                            if (cells.Length == addressItemsTotal)
                                                            {
                                                                county = "";                // County
                                                                postcode = cells[2];        // Postcode
                                                            }
                                                            else if (cells.Length == (addressItemsTotal + 1))
                                                            {
                                                                county = cells[2];          // County
                                                                postcode = cells[3];        // cells[3]     // Postcode
                                                            }
                                                            break;
                                                        }
                                                        case "map61":
                                                        {
                                                            realCellValue = district.Trim();
                                                            break;
                                                        }
                                                        case "map62":
                                                        {
                                                            realCellValue = town.Trim();
                                                            break;
                                                        }
                                                        case "map63":
                                                        {
                                                            realCellValue = county.Trim();
                                                            break;
                                                        }
                                                        case "map64":
                                                        {
                                                            realCellValue = postcode.Trim();
                                                            break;
                                                        }
                                                        case "map7":
                                                        {
                                                            var newCell2 = value4ThRow.Descendants<Cell>().ToArray()[columnIndex];
                                                            var newCellTempValue2 = GetCellValue(spredsheetDocument, newCell2, numberFormatId, formatCode);
                                                            var cell2FastLenderLtd = rowsFastLenderLtd[j].Descendants<Cell>().FirstOrDefault(x => x.CellReference == (newCellTempValue2 + (j + discrepancy).ToString()));
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd).Replace(",", "");
                                                            var realValue2 = GetCellValue(sprSheetDocumentFastLenderLtd, cell2FastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd).Replace(",", "");     // assuming here the cell2Toork has the same formats 

                                                            Row value5ThRow = rows[extraRowNumber + 2];
                                                            var newCell3 = value5ThRow.Descendants<Cell>( ).FirstOrDefault(x => x.CellReference == "X5");
                                                            var newCellTempValue3 = GetCellValue(spredsheetDocument, newCell3, numberFormatId, formatCode);
                                                            var cell3FastLenderLtd = rowsFastLenderLtd[j].Descendants<Cell>().FirstOrDefault(x => x.CellReference == (newCellTempValue3 + (j + discrepancy).ToString()));
                                                            var realValue3 = GetCellValue(sprSheetDocumentFastLenderLtd, cell3FastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd).Replace(",", "");     // assuming here the cell3Toork has the same formats 

                                                            var doubleValue1 = Convert.ToDouble(realValue1 == "" ? "0" : realValue1);
                                                            var doubleValue2 = Convert.ToDouble(realValue2 == "" ? "0" : realValue2);
                                                            var doubleValue3 = Convert.ToDouble(realValue3 == "" ? "0" : realValue3);

                                                            if (doubleValue1 > 0.0)
                                                            {
                                                                realCellValue = doubleValue1;
                                                            }
                                                            else if (doubleValue2 > 0.0)
                                                            {
                                                                realCellValue = doubleValue2 / doubleValue3;
                                                            }
                                                            else
                                                            {
                                                                realCellValue = "0";
                                                            }

                                                            break;
                                                        }
                                                        case "map8":
                                                        case "map9":
                                                        {
                                                            realCellValue = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);

                                                            if (realCellValue == null || ((string)realCellValue) == "")
                                                            {
                                                                //errorsOnInvalidCoulumns.Add(string.Format("The margin does not have any value in {0} row", dt.Rows.Count));
                                                            }
                                                            else
                                                            {
                                                                realCellValue = ((double.Parse((string)realCellValue)) * 100.0).ToString(CultureInfo.InvariantCulture);
                                                            }
                                                            break;
                                                        }
                                                        case "map10":
                                                        {
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                            realCellValue = MapFastLenderLtdPropertyTypes((string)(dt.Rows[dt.Rows.Count - 1][0]), realValue1);
                                                            break;
                                                        }
                                                        case "map11":
                                                        {
                                                            if ( !(dt.Rows.Count > 1 && (string) dt.Rows[dt.Rows.Count - 1][0] == (string) dt.Rows[dt.Rows.Count - 2][0]) )
                                                            {
                                                                var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                                realCellValue = MapFastLenderLtdLoanPurpose((string)(dt.Rows[dt.Rows.Count - 1][0]), realValue1);
                                                            }
                                                            break;
                                                        }
                                                        case "map12":
                                                        {
                                                            if (!(dt.Rows.Count > 1 && (string)dt.Rows[dt.Rows.Count - 1][0] == (string)dt.Rows[dt.Rows.Count - 2][0]))
                                                            {
                                                                var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                                realCellValue = MapFastLenderLtdRepaymentMethod((string)(dt.Rows[dt.Rows.Count - 1][0]), realValue1, additionalData);
                                                            }
                                                            break;
                                                        }
                                                        case "map13":
                                                        {
                                                            var realValue1 = GetCellValue(sprSheetDocumentFastLenderLtd, cellFastLenderLtd, numberFormatIdFastLenderLtd, formatCodeFastLenderLtd);
                                                            realCellValue = MapFastLenderLtdTenure((string)(dt.Rows[dt.Rows.Count - 1][0]), realValue1);
                                                            break;
                                                        }
                                                        default:
                                                            // MK TODO What do we do here?
                                                            ;
                                                            break;
                                                    }
                                                }
                                            }
                                        }

                                        dt.Rows[dt.Rows.Count - 1][columnIndex] = realCellValue;
                                    }
                                    #endregion cellLoopForAfter1stRows

                                    if (isLastRow)
                                        break;

                                    var appIdCell = dt.Rows[dt.Rows.Count - 1][0];
                                    if (appIdCell is DBNull || ((string)appIdCell) == "")
                                    {
                                        dt.Rows.RemoveAt(dt.Rows.Count - 1);
                                        continue;
                                    }
                                    if (sheetName == Constants.Contacts || sheetName == Constants.Properties || sheetName == Constants.OtherSecurities)
                                    {
                                        var idCell = dt.Rows[dt.Rows.Count - 1][1];
                                        if (idCell is DBNull)
                                        {
                                            dt.Rows[dt.Rows.Count - 1][1] = "";
                                        }
                                    }
                                    if ((dt.Rows.Count > 1) && (sheetName == Constants.Accounts || sheetName == Constants.Contacts))
                                    {
                                        if ((string)appIdCell == (string)dt.Rows[dt.Rows.Count - 2][0])
                                        {
                                            dt.Rows.RemoveAt(dt.Rows.Count - 1);
                                        }
                                    }
                               }
                                #endregion loopsForRows
                            }
                        }
                    }
#pragma warning disable CS0168 // The variable 'ex' is declared but never used
                    catch (Exception ex)
#pragma warning restore CS0168 // The variable 'ex' is declared but never used
                    {
                        var logger = new LoggerConfiguration()
                                                        .WriteTo.Console()
                                                        .CreateLogger();
                        logger.Information("The system encountered the following error{Message}", ex.Message);
                    }
                }
            }

            return _errorsOnInvalidCoulumns;
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        private static string MapFastLenderLtdChargeType(string propType)
        {
            string result;

            if (propType.Length == 0)
                return FastLenderLtdConstants.SecondCharge;

            switch (propType.Substring(0, 1).ToUpper())
            {
                case "Y":
                    {
                        result = FastLenderLtdConstants.FirstCharge;
                        break;
                    }
                case "N":
                    {
                        result = FastLenderLtdConstants.SecondCharge;
                        break;
                    }
                default:
                    {
                        result = FastLenderLtdConstants.NotYetCharged;
                        break;
                    }
            }

            return result;
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        private static int MapFastLenderLtdPropertyTypes(string appId, string propType)
        {
            var fieldName = "PropertyType";
            int result = 63;

            DataTable table;
            WebServiceStorProcs.GetOriginatorEnumMappings(m_ConnStr, m_OriginatorId, fieldName, out table);

            EnumerableRowCollection results = from myRow in table.AsEnumerable()
                where (string.Equals((myRow.Field<string>(Constants.OriginatorValue)).Trim().ToUpper(), propType.Trim().ToUpper(), StringComparison.InvariantCultureIgnoreCase))
                select myRow;

            var firstOrDefault = results.Cast<DataRow>().FirstOrDefault();
            if (firstOrDefault != null)
            {
                result = Convert.ToInt32(firstOrDefault[Constants.EnumName]);
            }
            else
            {
                // validation error
                _errorsOnInvalidCoulumns.Add(string.Format(Constants.FileRejectedMsg + "The '{0}' value in column '{1}' is invalid for Application = {2}", propType, fieldName, appId));
                result = 63;
            }

            return result;
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        private static string MapFastLenderLtdLoanPurpose(string appId, string enumValue)
        {
            const string fieldName = "LoanPurpose";
            const string result = "Purchase";

            return MapEnumStringValue(appId, enumValue, fieldName, result);
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        private static string MapFastLenderLtdRepaymentMethod(string appId, string enumValue, Dictionary<string, Pair> additionalData)
        {
            const string fieldName = "RepaymentMethod";
            const string result = "InterestOnly";

            additionalData.Add(appId, new Pair() { Name = fieldName, Value = enumValue });

            return MapEnumStringValue(appId, enumValue, fieldName, result);
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        private static string MapFastLenderLtdTenure(string appId, string enumValue)
        {
            const string fieldName = "Tenure";
            const string result = "Unknown";

            return MapEnumStringValue(appId, enumValue, fieldName, result);
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        private static string MapEnumStringValue(string appId, string enumValue, string fieldName, string result)
        {
            DataTable table;
            WebServiceStorProcs.GetOriginatorEnumMappings(m_ConnStr, m_OriginatorId, fieldName, out table);

            EnumerableRowCollection results = from myRow in table.AsEnumerable()
                                              where (string.Equals((myRow.Field<string>(Constants.OriginatorValue)).Trim().ToUpper(), enumValue.Trim().ToUpper(), StringComparison.CurrentCultureIgnoreCase))
                                              select myRow;

            var firstOrDefault = results.Cast<DataRow>().FirstOrDefault();
            if (firstOrDefault != null)
            {
                result = (string)firstOrDefault[Constants.EnumName];
            }
            else
            {
                // validation error
                _errorsOnInvalidCoulumns.Add(string.Format(Constants.FileRejectedMsg + "The '{0}' value in column '{1}' is invalid for Application = {2}", enumValue, fieldName, appId));
                result = enumValue;
            }

            return result;
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        public static void ExportToExcelFile(DataSet dataSet, string fileName, bool isCsvFile)
        {
            using (var workbook = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                //var workbookPart = workbook.AddWorkbookPart();
                workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };

                CreateStyleSheet(workbook.WorkbookPart.Workbook);

                uint sheetId = 1;

                foreach (DataTable table in dataSet.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    sheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                    var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<Sheet>().Any())
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    var sheet = new Sheet { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    var headerRow = new Row ()
                    {
                        RowIndex = 1u
                    };

                    string lastRef;
                    var columns = new List<string>();

                    List<string> columnRefs;
                    _excelTemplateColumnCellReferences.TryGetValue(table.TableName, out columnRefs);

                    if (columnRefs == null)
                    {
                        throw new NullReferenceException("The _excelTemplateColumnCellReferences dictionary does not contain the columnRefs list!..");
                    }

                    const string poolCodeName = "PoolCode";
                    int columnIndx = 0;
                    foreach (DataColumn column in table.Columns)
                    {
                        if (column.ColumnName == poolCodeName)
                            continue;

                        columns.Add(column.ColumnName);

                        lastRef = columnRefs[columnIndx] + 1;

                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue((isCsvFile)
                                                        ? column.ColumnName
                                                        : column.ColumnName.Substring(column.ColumnName.IndexOf("_", StringComparison.Ordinal) + 1)),
                            CellReference = lastRef
                        };
                        headerRow.AppendChild(cell);

                        columnIndx++;
                    }

                    sheetData.AppendChild(headerRow);

                    var rowIndex = 1;
                    foreach (DataRow dsrow in table.Rows)
                    {
                        var newRow = new Row()
                        {
                            RowIndex = (uint)(rowIndex + 1)
                        };
                        
                        columnIndx = 0;
                        foreach (var column in columns)
                        {
                            if (column == poolCodeName)
                                continue;

                            lastRef = columnRefs[columnIndx] + (rowIndex + 1);
                            var cell = new Cell
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(dsrow[column].ToString()),
                                CellReference = lastRef
                            };

                            newRow.AppendChild(cell);
                            columnIndx++;
                        }

                        sheetData.AppendChild(newRow);
                        rowIndex++;
                    }
                }
            }
        }

        public static uint GetNumberFormatIdFromCellStyleIndex(CellFormats cellFormats, Cell cell)
        {
            uint numberFormatIdValue = 0u;

            if (cell.StyleIndex == null) 
                return numberFormatIdValue;


            if (cellFormats.ElementAt((int)cell.StyleIndex.Value) == null)
                return numberFormatIdValue;


            numberFormatIdValue = ((CellFormat)cellFormats.ElementAt((int)cell.StyleIndex.Value)).NumberFormatId.Value;

            return numberFormatIdValue;
        }

        private static string GetCellRef(string cellRef, uint rowIndex)
        {
            // This works like this:
            // "A1"     and     rowIndex = 1    ==>     cellRef = "A"
            // "AB22"   and     rowIndex = 22   ==>     cellRef = "AB"

            var rIndx = cellRef.LastIndexOf(rowIndex.ToString(), StringComparison.Ordinal);

            if (rIndx < 0)
                return cellRef;

            return cellRef.Substring(0, rIndx);
        }

        private static string NewColumnName(string cellRef, string columnName, bool flag)
        {
            return 
                flag ?
                    (cellRef + Constants.Underscore + columnName) : 
                    columnName;
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell, uint numberFormatId, string formatCode = "")
        {
            if (cell.CellValue == null)
                return "";

            var stringTablePart = document.WorkbookPart.SharedStringTablePart ;         // DocumentFormat.OpenXml.Packaging.WorkbookPart.SharedStringTablePart
            var value = cell.CellValue.InnerXml;

            if (cell.DataType != null)
            {
                switch (cell.DataType.Value)
                {
                    case CellValues.SharedString:
                    {
                        // For shared strings, look up the value in the shared strings table.
                        return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;   // this is normally a string or a string lookup values which are actually re-used in the Excel sheet
                    }
                    case CellValues.Boolean:        // these cases below have not been tested yet though! 13/07/2016
                        switch (value)
                        {
                            case "0":
                                return Constants.False;
                            //case "1":
                            default:
                                return Constants.True;
                        }
                    case CellValues.Number:
                    case CellValues.String:
                    case CellValues.InlineString:
                    {
                        return value;
                    }
                    case CellValues.Error:
                    {
                        return "";      // ???
                    }
                    case CellValues.Date:
                    {
                        return DateTime.FromOADate(int.Parse(value)).ToShortDateString();       // ??
                    }
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
            else if (value == Constants.Blank)
            {
                return value; // this is the empty string
            }
            else
            {
                if ((numberFormatId >= 14 && numberFormatId <= 22)
                    || (formatCode.Contains("d") && formatCode.Contains("m") && formatCode.Contains("yy"))) 
                {
                    return DateTime.FromOADate(double.Parse(value)).ToShortDateString(); // this is the DateTime type
                }
                else
                {
                    return value; // this is the actual string value of any numeric type like decimal type, etc.
                }
            }
        }

        private static void CreateStyleSheet(Workbook workbook)
        {
            // Stylesheet
            var workbookStylesPart = workbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = new Stylesheet();

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats != null)
            {
                // <Appending NumberingFormats>
                var numberingFormats = new NumberingFormats()
                {
                    Count = workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Count
                };
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats)
                {
                    var numberingFormat = (NumberingFormat) openXmlElement;
                    var newNumberingFormat = (NumberingFormat)numberingFormat.CloneNode(true);
                    numberingFormats.Append(newNumberingFormat);
                }
                workbookStylesPart.Stylesheet.Append(numberingFormats);
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Colors != null)
            {
                var colors = new MruColors();
                foreach (var color in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Colors.MruColors)
                {
                    var newColor = (Color)color.CloneNode(true);
                    colors.Append(newColor);
                }
                workbookStylesPart.Stylesheet.Append(colors);
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts != null)
            {
                var fonts = new Fonts();                                // <Appending Fonts>
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Fonts)
                {
                    var font = (Font) openXmlElement;
                    var newFont = (Font)font.CloneNode(true);
                    fonts.Append(newFont);
                }
                workbookStylesPart.Stylesheet.Append(fonts);
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills != null)
            {
                var fills = new Fills();                                 // <Appending Fills>
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Fills)
                {
                    var fill = (Fill) openXmlElement;
                    var newFill = (Fill)fill.CloneNode(true);
                    fills.Append(newFill);
                }
                workbookStylesPart.Stylesheet.Append(fills);
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders != null)
            {
                var borders = new Borders();                             // <Appending Border>
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.Borders)
                {
                    var border = (Border) openXmlElement;
                    var newBorder = (Border)border.CloneNode(true);
                    borders.Append(newBorder);
                }
                workbookStylesPart.Stylesheet.Append(borders);
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellStyleFormats != null)
            {
                var cellStyleFormats = new CellStyleFormats ();        // <Appending CellFormats>
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellStyleFormats)
                {
                    var cellStyleFormat = (CellFormat) openXmlElement;
                    var newCellFormat = (CellFormat)cellStyleFormat.CloneNode(true);
                    cellStyleFormats.Append(newCellFormat);
                }
                workbookStylesPart.Stylesheet.Append(cellStyleFormats);
                //workbookStylesPart.Stylesheet.CellStyleFormats = cellStyleFormats;
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats != null)
            {
                var cellFormats = new CellFormats();                    // <Appending CellFormats>
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats)
                {
                    var cellFormat = (CellFormat) openXmlElement;
                    var newCellFormat = (CellFormat)cellFormat.CloneNode(true);
                    cellFormats.Append(newCellFormat);
                }
                workbookStylesPart.Stylesheet.Append(cellFormats);
            }

            if (workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellStyles != null)
            {
                var cellStyles = new CellStyles();                      // <Appending CellStyles>
                foreach (var openXmlElement in workbook.WorkbookPart.WorkbookStylesPart.Stylesheet.CellStyles)
                {
                    var cellStyle = (CellStyle) openXmlElement;
                    var newCellStyle = (CellStyle)cellStyle.CloneNode(true);
                    cellStyles.Append(newCellStyle);
                }
                workbookStylesPart.Stylesheet.Append(cellStyles);
            }
        }
        //-----------------------------------------------------------------------------------------------------------------------------
        public static void ExportErrorsToExcelFile(List<Tuple<string, string, string, string>> errorDetails, string fileName, bool isCsvFile)
        {
            using (var workbook = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };

                CreateStyleSheet(workbook.WorkbookPart.Workbook);

                uint sheetId = 1;

                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                if (sheets.Elements<Sheet>().Any())
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                var sheet = new Sheet { Id = relationshipId, SheetId = sheetId, Name = "Validation Errors" };
                sheets.Append(sheet);

                var headerRow = new Row()
                {
                    RowIndex = 1u
                };

                headerRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue("AccountId"), CellReference = "A1" });
                headerRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue("XPath"), CellReference = "B1" });
                headerRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue("Message"), CellReference = "C1" });
                headerRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue("Severity"), CellReference = "D1" });

                sheetData.AppendChild(headerRow);

                var rowIndex = 1;
                foreach (var errorDetail in errorDetails)
                {
                    var newRow = new Row()
                    {
                        RowIndex = (uint)(rowIndex + 1)
                    };

                    newRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue(errorDetail.Item1), CellReference = "A" + (rowIndex + 1)});
                    newRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue(errorDetail.Item2), CellReference = "B" + (rowIndex + 1)});
                    newRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue(errorDetail.Item3), CellReference = "C" + (rowIndex + 1)});
                    newRow.AppendChild(new Cell { DataType = CellValues.String, CellValue = new CellValue(errorDetail.Item4), CellReference = "D" + (rowIndex + 1)});

                    sheetData.AppendChild(newRow);
                    rowIndex++;
                }
            }
        }

        static bool CheckSheetState(Sheet sheet, List<string> supportedSheetNames)
        {
            bool ret = false;

            if (sheet.State != null && sheet.State == Constants.Hidden)
                ret = true;

            if (!supportedSheetNames.Contains(sheet.Name.ToString().ToUpper()))
                ret = true;

            return ret;

        }
    }
    
}