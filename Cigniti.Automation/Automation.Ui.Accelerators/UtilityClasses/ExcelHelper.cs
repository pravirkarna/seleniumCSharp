using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automation.Ui.Accelerators.UtilityClasses
{
    public static class ExcelHelper
    {

        private static string OleDFourConnectionStr = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007  
        private static string OledTwelveConnectionStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HRD=NO';";//for above excel 2007  
        public static DataTable ReadExcel(string fileName, string fileExt, string sheetName)
        {
            string connectionString = FetchConnectionsString(fileExt, fileName);

            return GetData(connectionString, sheetName);
        }

        public static DataTable GetData(string connectionString, string sheetName)
        {
            using (OleDbConnection con = new OleDbConnection(connectionString))
            {
                var dtexcel = new DataTable();
                var oleAdpt = new OleDbDataAdapter($"select * from [{ sheetName } $]", con);
                oleAdpt.Fill(dtexcel);
                return dtexcel;
            }
        }

        /// <summary>
        /// Read Excel
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        public static DataTable ReadExcel(string fileName, string sheetName)
        {
            var fileExt = Path.GetExtension(fileName);

            if (fileExt != ".xls" && fileExt != ".xlsx")
                throw new Exception("Not a .xls or .xlsx file");
            return ReadExcel(fileName, fileExt, sheetName);
        }
       
        private static string FetchConnectionsString(string fileExt, string fileName)
        {
            return fileExt.CompareTo(".xls") == 0
                   ? string.Format(OleDFourConnectionStr, fileName)
                   : string.Format(OledTwelveConnectionStr, fileName);
        }
    }
}