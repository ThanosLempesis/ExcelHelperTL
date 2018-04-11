using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IOExcelHelpers;
using Constants = IOExcelHelpers.Constants;

namespace ExcelHelperUnitTests
{
    /// <summary>
    ///This is a test class for ExcelHelperTest and is intended
    ///to contain all ExcelHelperTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ExcelHelperTest
    {
        private const string DictionaryColumnName = "ColumnName";
        private const string DictionaryDataType = "DataType";

        internal const string Accounts = "ACCOUNTS";
        internal const string Properties = "PROPERTIES";
        internal const string Contacts = "CONTACTS";
        internal const string OtherSecurities = "OTHER SECURITIES";

        internal static readonly List<string> SupportedSheetNames = new List<string> { Accounts, Properties, Contacts, OtherSecurities };


        public static void ReadTemplate(DataSet dataSet, string testDataFileName)
        {
            var folderPath = ReflectionHelper.GetExecutingAssemblyFolder(Assembly.GetExecutingAssembly());

            var testDataFullFileName = Path.Combine(folderPath, testDataFileName);

            var shortcutTarget = TextFileIoHelper.GetShortcutTarget(testDataFullFileName);

            if (!File.Exists(shortcutTarget))
            {
                Assert.Inconclusive("The test data file {0} does not exist", shortcutTarget);
            }

            ExcelHelper.ReadTemplateIntoTable(dataSet, shortcutTarget, SupportedSheetNames, true);
           
        }

        public static void ReadDictionary(DataSet dataSet, string testDataFileName)
        {
            const string accounts = "ACCOUNTS";
            const string properties = "PROPERTIES";
            const string contacts = "CONTACTS";
            const string otherSecurities = "OTHER SECURITIES";

            var folderPath = ReflectionHelper.GetExecutingAssemblyFolder(Assembly.GetExecutingAssembly());

            var testDataFullFileName = Path.Combine(folderPath, testDataFileName);

            var shortcutTarget = TextFileIoHelper.GetShortcutTarget(testDataFullFileName);

            if (!File.Exists(shortcutTarget))
            {
                Assert.Inconclusive("The test data file {0} does not exist", shortcutTarget);
            }

            var supportedSheetNames = new List<string> { accounts, properties, contacts, otherSecurities };

            ExcelHelper.ReadDictionaryIntoTable(dataSet, shortcutTarget, supportedSheetNames);

        }

        private static List<string> ReadTemplateBasedExcelAsDataTablesTest(DataSet dataSet, string fileName)
        {
            var errorsOnInvalidCoulumns = new List<string>();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                var workBookRoot = spreadsheetDocument.WorkbookPart.Workbook;
                var sheets = workBookRoot.Sheets;

                var numberingFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats.Elements<NumberingFormat>().ToList();
                var cellFormats = workBookRoot.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats;

                foreach (var sh in sheets)
                {
                    var sheet = (Sheet)sh;
                    var relationshipId = sheet.Id.Value;
                    string sheetName = sheet.Name;

                    if ((sheet).State != null && (sheet).State == Constants.Hidden)
                        continue;

                    var dataTable = dataSet.Tables[sheetName];
                    var worksheetPart = spreadsheetDocument.WorkbookPart.GetPartById(relationshipId) as WorksheetPart;
                    if (worksheetPart != null)
                    {
                        var workSheet = worksheetPart.Worksheet;
                        var sheetData = workSheet.GetFirstChild<SheetData>();
                        var rows = sheetData.Descendants<Row>();                    // IEnumerable<Row>                     // DocumentFormat.OpenXml.Spreadsheet.Row   

                        // Loop through the Worksheet rows.
                        foreach (var row in rows)                                   // DocumentFormat.OpenXml.Spreadsheet.Row
                        {
                            dataTable.Rows.Add();

                            var cellCounter = -1;
                            foreach (var cell in row.Descendants<Cell>())
                            {
                                // Below is also for testing purpose only
                                cellCounter++;
                                if (cellCounter >= 1)
                                    break;

                                // Below is also changed for testing purpose only to work for the test case of testing the blank cell
                                if (cell.CellValue == null)
                                {
                                    dataTable.Rows[dataTable.Rows.Count - 1][0] = ExcelHelper.GetCellValue(spreadsheetDocument, cell, 0u);
                                    dataTable.Rows[dataTable.Rows.Count - 1][1] = 0u;
                                    dataTable.Rows[dataTable.Rows.Count - 1][2] = 0u;
                                    dataTable.Rows[dataTable.Rows.Count - 1][3] = "";

                                    continue;
                                }

                                var numberFormatId = ExcelHelper.GetNumberFormatIdFromCellStyleIndex(cellFormats, cell);

                                var matchingNumFormat = numberingFormats.Where(i => i.NumberFormatId.Value == numberFormatId).ToList();

                                var formatCode = "";
                                if (matchingNumFormat.Any())
                                {
                                    formatCode = matchingNumFormat.First().FormatCode;
                                }

                                dataTable.Rows[dataTable.Rows.Count - 1][0] = ExcelHelper.GetCellValue(spreadsheetDocument, cell, numberFormatId, formatCode);
                                dataTable.Rows[dataTable.Rows.Count - 1][1] = (cell.StyleIndex != null) ? cell.StyleIndex.Value : 0u;
                                dataTable.Rows[dataTable.Rows.Count - 1][2] = numberFormatId;
                                dataTable.Rows[dataTable.Rows.Count - 1][3] = formatCode;
                            }
                        }
                    }
                }
            }

            return errorsOnInvalidCoulumns;
        }

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext { get; set; }

        /// <summary>
        ///A test for GetCellValue
        ///</summary>
        [TestMethod()]
        public void GetCellValueTest()
        {
            var methodName = DiagnosticsHelper.GetCallingMethodName(1);

            const string testDataFileName = "ExcelHelperTestData.xlsx";
            const string sheetName1 = "DateTime Types Test";
            const string sheetName2 = "Other Data Types Test";

            var dataSet = new DataSet();

            var folderPath = ReflectionHelper.GetExecutingAssemblyFolder(Assembly.GetExecutingAssembly());

            var testDataFullFileName = Path.Combine(folderPath, testDataFileName);

            if (!File.Exists(testDataFullFileName))
            {
                Assert.Inconclusive("The test data file {0} does not exist", testDataFullFileName);
            }

            // I am checking if the file is open from someone else
            try
            {
                SpreadsheetDocument workbook = SpreadsheetDocument.Open(testDataFullFileName, false);
                workbook.Close();
            }
            catch
            {
                Assert.Inconclusive("The test data file {0} is open from a differnet resource", testDataFullFileName);
            }
           

            var dataTable1 = new DataTable(sheetName1);
            var dataTable2 = new DataTable(sheetName2);

            dataTable1.Columns.Add("Value");
            dataTable1.Columns.Add("CellStyleIndex");
            dataTable1.Columns.Add("numberFormatId");
            dataTable1.Columns.Add("formatCode");

            dataTable2.Columns.Add("Value");
            dataTable2.Columns.Add("CellStyleIndex");
            dataTable2.Columns.Add("numberFormatId");
            dataTable2.Columns.Add("formatCode");
            
            dataSet.Tables.Add(dataTable1);
            dataSet.Tables.Add(dataTable2);

            var errorList = ReadTemplateBasedExcelAsDataTablesTest(dataSet, testDataFullFileName);

            Assert.IsTrue(errorList.Count == 0, "There is unexpected errors in " + methodName + " call!..");

            Assert.AreEqual((dataTable1.Rows[0]).ItemArray[0].ToString(), "09/04/1966", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[1]).ItemArray[0].ToString(), "10/07/1989", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[2]).ItemArray[0].ToString(), "09/07/1972", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[3]).ItemArray[0].ToString(), "21/08/1965", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[4]).ItemArray[0].ToString(), "08/11/1964", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[5]).ItemArray[0].ToString(), "11/12/1955", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[6]).ItemArray[0].ToString(), "24/11/1962", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[7]).ItemArray[0].ToString(), "25/11/1962", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[8]).ItemArray[0].ToString(), "08/07/2016", "In " + methodName + " call");
            Assert.AreEqual((dataTable1.Rows[9]).ItemArray[0].ToString(), "08/07/2016", "In " + methodName + " call");

            Assert.AreEqual((dataTable2.Rows[0]).ItemArray[0].ToString(), "100065143", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[1]).ItemArray[0].ToString(), "N", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[2]).ItemArray[0].ToString(), "Y", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[3]).ItemArray[0].ToString(), "230580", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[4]).ItemArray[0].ToString(), "230581", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[5]).ItemArray[0].ToString(), "Homer Simpson", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[6]).ItemArray[0].ToString(), "126827.5", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[7]).ItemArray[0].ToString(), "", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[8]).ItemArray[0].ToString(), " ", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[9]).ItemArray[0].ToString(), "09890048", "In " + methodName + " call");
        }      

        /// <summary>
        ///A test for ReadTemplateIntoTable
        ///</summary>
        [TestMethod()]
        public void ReadTemplateIntoTableTest()
        {
            var methodName = DiagnosticsHelper.GetCallingMethodName(1);

            var dataSet = new DataSet();

            //const string testDataFileName = "CleverLenderLtd Originations Data Feed Template.xlsx.lnk";         // MK: mind the '*.lnk' here!!!...
            const string testDataFileName = "CleverLenderLtd PBTL Data Feed Template.xlsx.lnk";                   // MK: mind the '*.lnk' here!!!...

            ReadTemplate(dataSet, testDataFileName);

            var dataTable = dataSet.Tables["ACCOUNTS"];

            Assert.AreEqual((dataTable.Columns[0].ToString()).ToUpper(), "A_ApplicationID".ToUpper(), "In " + methodName + " call");
        }

        /// <summary>
        ///A test for ReadDictionaryIntoTable
        ///</summary>
        [TestMethod()]
        public void ReadDictionaryIntoTableTest()
        {
            var methodName = DiagnosticsHelper.GetCallingMethodName(1);

            const string testDataFileName = "CleverLenderLtd Originations Data Feed Dictionary.xlsx.lnk";         // MK: mind the '*.lnk' here!!!...
    
            var dataSet = new DataSet();

            ReadDictionary(dataSet, testDataFileName);

            var dataTable = dataSet.Tables["ACCOUNTS"];

            Assert.AreEqual((dataTable.Columns[0].ToString()).ToUpper(), DictionaryColumnName.ToUpper(), "In " + methodName + " call");
            Assert.AreEqual((dataTable.Columns[1].ToString()).ToUpper(), DictionaryDataType.ToUpper(), "In " + methodName + " call");

            Assert.AreEqual(((dataTable.Rows[0]).ItemArray[0].ToString()).Trim().ToUpper(), "ApplicationID".Trim().ToUpper(), "In " + methodName + " call");
            Assert.AreEqual(((dataTable.Rows[0]).ItemArray[1].ToString()).Trim().ToUpper(), "INT".Trim().ToUpper(), "In " + methodName + " call");
            Assert.AreEqual(((dataTable.Rows[1]).ItemArray[0].ToString()).Trim().ToUpper(), "Further Advance Flag".Trim().Replace(" ", "").ToUpper(), "In " + methodName + " call");
            Assert.AreEqual(((dataTable.Rows[1]).ItemArray[1].ToString()).Trim().ToUpper(), "Y/N".Trim().ToUpper(), "In " + methodName + " call");
        }

        /// <summary>
        ///A test for ReadTemplateBasedExcelAsDataTables
        ///</summary>
        [TestMethod()]
        public void ReadTemplateBasedExcelAsDataTablesTest()
        {
            var methodName = DiagnosticsHelper.GetCallingMethodName(1);

            //const string testDataFileName = "CleverLenderLtd Originations Data Feed Template.xlsx.lnk";         // MK: mind the '*.lnk' here!!!...
            const string testTemplateFileName = "CleverLenderLtd PBTL Data Feed Template.xlsx.lnk";               // MK: mind the '*.lnk' here!!!...

            //const string testDataFileName = "ExcelHelperTestData.xlsx";
            const string testDataFileName = "CleverLenderLtd_PBTL_08-07-2016.xlsx";

            var dataSet = new DataSet();
            //var dictionaryTables = new DataSet();

            var folderPath = ReflectionHelper.GetExecutingAssemblyFolder(Assembly.GetExecutingAssembly());

            var testDataFullFileName = Path.Combine(folderPath, testDataFileName);

            if (!File.Exists(testDataFullFileName))
            {
                Assert.Inconclusive("The test data file {0} does not exist in " + methodName + " call", testDataFullFileName);
            }

            ReadTemplate(dataSet, testTemplateFileName);
            //ReadDictionary(dictionaryTables, testDictionaryFileName);

            var supportedSheetNames = new List<string> { "ACCOUNTS", "PROPERTIES", "CONTACTS" };
            var errorList = ExcelHelper.ReadTemplateBasedExcelAsDataTables(dataSet, testDataFullFileName, supportedSheetNames);

            Assert.IsTrue(errorList.Count == 0, "There is unexpected errors in " + methodName + " call!..");

            var dataTable = dataSet.Tables["ACCOUNTS"];

            Assert.AreEqual((dataTable.Rows[0]).ItemArray[0].ToString(), "100065143", "In " + methodName + " call");
            Assert.AreEqual((dataTable.Rows[1]).ItemArray[0].ToString(), "100071200", "In " + methodName + " call");

            var dataTable2 = dataSet.Tables["CONTACTS"];

            Assert.AreEqual((dataTable2.Rows[0]).ItemArray[7].ToString(), "09/04/1966", "In " + methodName + " call");
            Assert.AreEqual((dataTable2.Rows[1]).ItemArray[7].ToString(), "10/07/1989", "In " + methodName + " call");
        }
    }
}
