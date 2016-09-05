using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;

namespace ExcelLoader.Helpers
{
    class ExcelHelper {

        public static long MAX_POSSIBLE_ROWS = 1000000;

        #region Properties

        List<DataColumn> DataColumns { get; set; } = new List<DataColumn>();

        public ObservableCollection<RuleDefinition> Rules { get; set; } = new ObservableCollection<RuleDefinition>();

        private Workbooks workbooks { get; set; }
        private Workbook wb { get; set; }

        public Worksheet RuleWorksheet { get; private set; }

        public Worksheet DataWorksheet { get; private set; }
        public Application xlApp { get; private set; }
        public int NumLoadedRows { get; private set; }
        public bool WorkbookOpen { get; set; } = false;

        #endregion Properties


        public bool Load(string fileName)
        {
            try
            {
                OpenWorkbook(fileName);
                Console.WriteLine("Opened and initialized file : " + fileName);

                LoadRules();
                LoadDataColumns();
                // get the data now so the actual values can be calculated
                foreach (RuleDefinition rule in Rules)
                {
                    List<DataColumn> dataCols = new List<DataColumn>();
                    foreach (string colName in rule.ColumnsToUse)
                    {
                        dataCols.Add(DataColumns.Where(dc=> colName == dc.ColumnName).First());
                    }
                    rule.DataColumns = dataCols;
                }

                Console.WriteLine("Loaded excel file : " + fileName);
                return true;
            } catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return false;
            } finally
            {
                CloseWorkbook();
            }
        }

        // retrieve each rule from the rule worksheet
        private List<RuleDefinition> LoadRules()
        {
            Rules.Clear();
            if (RuleWorksheet == null)
            {
                throw new Exception("No Rule worksheet was found");
            }
             Range usedRange = RuleWorksheet.UsedRange;

            foreach (Range row in usedRange.Rows)
            {
                if (!string.IsNullOrEmpty(row.Cells[1]?.Value))
                LoadRule(row);
            }
            Console.WriteLine("Completed reading rules");
            return new List<RuleDefinition>();
        }

        //
        private void LoadRule(Range row)
        {
            RuleDefinition ruleDef = new RuleDefinition();
            IEnumerator cellEnum = row.Cells.GetEnumerator();
            // loop through each column and add to RuleDefinition list of column names
            List<string> colDefs = new List<string>();
            while(cellEnum.MoveNext())
            {
                var curCell = cellEnum.Current;
                string colValue = (curCell as Range).Value?.ToString();
                if (!string.IsNullOrEmpty(colValue))
                {
                    colDefs.Add(colValue.Trim());
                } else
                    break; // hit a null value
            }
            ruleDef.ColumnsToUse = colDefs;
            long actualRows = Convert.ToInt64(Math.Pow(NumLoadedRows, colDefs.Count));
            Rules.Add(ruleDef);
            Console.WriteLine("Added Rule: " + ruleDef.RuleText);
        }

        private void LoadDataColumns()
        {
            // create data columns by using the headers
            List<string> colHeaders = GetDataColumnHeaders();
            foreach( string colHeader in colHeaders)
            {
                DataColumns.Add(new DataColumn(colHeader));
            }

            // load all datacolumns 
            Range usedRange = DataWorksheet.UsedRange;
            bool first = true;
            foreach (Range row in usedRange.Rows)
            {
                if (first)
                {
                    first = false;
                    continue;
                }
                LoadDataColumn(row);
            }
            Console.WriteLine("Loaded Data Columns. Found " + DataColumns.Count.ToString() + " columns");
        }

        private void LoadDataColumn(Range row)
        {
            IEnumerator cellEnum = row.Cells.GetEnumerator();
            int i = 0;          
            while (cellEnum.MoveNext())
            {
                DataColumn curCol = DataColumns[i++];
                var curCell = cellEnum.Current;
                string colValue = (curCell as Range).Value;
                if (!string.IsNullOrEmpty(colValue))
                {
                    curCol.Add(colValue);
                }
            }
        }

        private List<string> GetDataColumnHeaders()
        {
            List<string> colHeaders = new List<String>();
            Range headerRow = DataWorksheet.Rows[1];
            IEnumerator cellEnum = headerRow.Cells.GetEnumerator();
            Console.Write("Getting headers: ");
            bool done = false;
            while (cellEnum.MoveNext() && !done)
            {
                var curCell = cellEnum.Current;
                string colValue = (curCell as Range).Value;
                if (string.IsNullOrEmpty(colValue))
                {
                    done = true;
                }
                else {
                    colHeaders.Add(colValue.Trim());
                    Console.Write(" | " + colValue + " | ");
                }
            }
            Console.WriteLine();
            return colHeaders;
        }

        public void GenerateOutputSheets(string fileName)
        {
            OpenWorkbook(fileName);
            RemoveOutputWorksheets();
            try
            {
                // for each rule, create a local list of ColumnDefinitions for that rule.
                // Then iterate through all permutations
                foreach (RuleDefinition rule in Rules)
                {
                    ExcelLoaderViewModel.GetInstance().CurrentStatus = "Loading: " + rule.RuleText;
                    var results = CartesianProduct(rule.DataColumns.Select(dc=> dc.ColData));
                    ExcelLoaderViewModel.GetInstance().CurrentStatus = "Writing: " + rule.RuleText;
                    WriteSheet(results);
                }
                //}
                //catch(Exception e)
                //{
                //    Console.WriteLine(e.StackTrace);
            }
            finally
            {
                CloseWorkbook();
            }
        }

        private IEnumerable<IEnumerable<string>> GetColumnsForRule(RuleDefinition rule)
        {
            var results = rule.ColumnsToUse
                .Select(colName => DataColumns.Find(dc => colName == dc.ColumnName))
                .Select(dc => dc.ColData);
            return results as IEnumerable<IEnumerable<string>>;
        }

        public IEnumerable<IEnumerable<object>> CartesianProduct(IEnumerable<IEnumerable<object>> inputs)
        {
            IEnumerable<object[]> prevItem = new object[][] { new object[0] };
            foreach (var input in inputs)
            {
                var currentInput = input;
                prevItem = prevItem.SelectMany(prevProductItem =>
                    from item in currentInput
                    select prevProductItem.Concat(new object[] { item }).ToArray());
            }
            return prevItem;
        }

        private void WriteSheet(IEnumerable<IEnumerable<object>> results)
        {
            if (results == null)
            {
                Console.WriteLine("No Strings Generated");
                return;
            }
            int i = 1;
            Worksheet resultSheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            foreach (var result in results)
            {
                if (i > MAX_POSSIBLE_ROWS)
                {
                    break;
                }
                else {
                    resultSheet.Cells[i++, 1] = string.Join(" ", result);
                }
            }
        }

        private void RemoveOutputWorksheets()
        {
            int numSheets = wb.Worksheets.Count;
            if (numSheets > 2)
            {
                for (int i = numSheets; i > 2; i--)
                {
                    wb.Worksheets[i].Delete();
                }
                wb.Save();
            }
        }
        private void OpenWorkbook(string fileName)
        {
            xlApp = new Application();
            xlApp.DisplayAlerts = false;
            workbooks = xlApp.Workbooks;
            wb = workbooks.Open(fileName);
            RuleWorksheet = wb.Sheets[1];
            DataWorksheet = wb.Sheets[2];
            NumLoadedRows = DataWorksheet.UsedRange.Rows.Count - 1;
            WorkbookOpen = true;
        }

        public void CloseWorkbook()
        {
            if (WorkbookOpen)
            {
                try
                {

                    foreach (var sheet in wb.Sheets)
                    {
                        Marshal.ReleaseComObject(sheet);
                    }
                    wb.Save();
                    wb.Close(0);
                    xlApp.Quit();

                    Marshal.ReleaseComObject(wb);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(xlApp);
                }
                catch
                {
                    // ignore
                }
            }
        }
    }
}
