using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelComparision {
    public static class Extensions {
        public static void ToCSV(this DataTable dt,string csvFile) {
            StringBuilder sb = new StringBuilder();

            var columnNames = dt.Columns.Cast<DataColumn>().Select(column => "\"" + column.ColumnName.Replace("\"", "\"\"") + "\"").ToArray();
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows) {
                var fields = row.ItemArray.Select(field => "\"" + field.ToString().Replace("\"", "\"\"") + "\"").ToArray();
                sb.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(csvFile, sb.ToString(), Encoding.Default);
        }
    }
}
