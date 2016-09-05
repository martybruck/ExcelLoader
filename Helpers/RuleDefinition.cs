using System;
using System.Collections.Generic;

namespace ExcelLoader.Helpers
{
    
    public class RuleDefinition
    {
        public long ActualOutputCount { get; set; }
        public long PossibleOutputCount {
            get
            {
                // muliply # of colums together
                if (ColumnsToUse.Count == 0)
                {
                    return 0;
                }
                long numResults = 1;
                foreach (DataColumn col in DataColumns)
                {
                    numResults = numResults * col.NumEntries;
                }
                ActualOutputCount = Convert.ToInt64(Math.Min(numResults, ExcelHelper.MAX_POSSIBLE_ROWS));
                return numResults;
            }
        }
        internal List<DataColumn> DataColumns { get; set; }

        private string ruleText;
        public string RuleText {
            get
            {
                return ruleText;
            }
            set
            {
                ruleText = value;
            }
        }
        public List<string> columnsToUse = new List<string>();
        public List<string> ColumnsToUse {
            get
            {
                return columnsToUse;
            } 
            set
            {
                columnsToUse = value;
                if (value != null)
                {
                    RuleText = string.Join(",", value);
                }
            }
        }
    }
}