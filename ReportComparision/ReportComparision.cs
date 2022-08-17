using ExcelComparision;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;



namespace TestProject2
{

    [TestClass]
    public class ReportComparision
    {
        [TestMethod]
        public void categorisationReportSRHQ()
        {
            DataTable dataTableReport = datatable("C:\\Test1\\testing\\Department Report_SR.xlsx", "Sheet");
            DataTable dataTableReportCompare = datatable("C:\\Test1\\testing\\Categorisation Report_HQ.xlsx", "Sheet");
            DataTable table = new DataTable();

            int countOfSR = dataTableReport.Rows.Count;
            int countOfHQ = dataTableReportCompare.Rows.Count;
            bool isDeptCodeExist = true;
            foreach (DataColumn column in dataTableReport.Columns)
            {
                table.Columns.Add(column.ColumnName);
            }

            foreach (DataRow dr in dataTableReport.Rows)
            {
                var depVal = dr.ItemArray[0];
                if ((countOfSR - 1) == dataTableReport.Rows.IndexOf(dr))
                {
                    break;
                }
                DataRow row = table.NewRow();
                foreach (DataRow dr2 in dataTableReportCompare.Rows)
                {
                    var depValHQ = dr2.ItemArray[0];

                    if (Convert.ToInt32(depValHQ) == Convert.ToInt32(depVal))
                    {
                        table = matchedRow(dr, dr2, row, table);
                        isDeptCodeExist = true;
                        break;
                    }
                    else
                    {
                        isDeptCodeExist = false;
                    }
                }
                Console.WriteLine("Department code :: " + depVal + " is exist" + isDeptCodeExist);
                if (!isDeptCodeExist)
                {
                    for (int i = 0; i < dr.ItemArray.Length; i++)
                    {
                        var columnName = table.Columns[i];
                        if (i == 0)
                        {
                            row[columnName] = depVal;
                        }
                        else
                        {
                            row[columnName] = "Null in HQ";
                        }

                    }
                    table.Rows.Add(row.ItemArray);
                }
            }
            table.ToCSV("C:\\Test1\\testing\\Categorisation Report_HQ_Differece.csv");
            Console.WriteLine("Execution Completed..");
        }

        public DataTable matchedRow(DataRow d1, DataRow d2, DataRow newRow, DataTable dataTable)
        {

            for (int i = 0; i < d1.ItemArray.Length; i++)
            {
                var columnName = dataTable.Columns[i];
                if ((d1.ItemArray[i] != System.DBNull.Value) && (d2.ItemArray[i] != System.DBNull.Value))
                {
                    var origional = Convert.ToDecimal(d1.ItemArray[i]);
                    origional = Math.Round(origional, 1);
                    var compare = Convert.ToDecimal(d2.ItemArray[i]);
                    compare = Math.Round(compare, 1);
                    if (i == 0)
                    {
                        newRow[columnName] = compare;
                    }
                    else
                    {
                        if (origional != compare)
                        {

                            newRow[columnName] = compare;
                        }
                    }
                }
                else
                {
                    newRow[columnName] = (d2.ItemArray[i] != System.DBNull.Value) ? d2.ItemArray[i] : "Null in HQ";
                }

            }
            dataTable.Rows.Add(newRow.ItemArray);
            return dataTable;
        }
        public DataTable datatable(string path, string sheetName)
        {
            DataTable dt = new DataTable();
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";
                    comm.Connection = conn;
                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }



            }
        }
    }
}