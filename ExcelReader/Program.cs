using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using static OfficeOpenXml.ExcelErrorValue;

namespace ExcelReader
{
    internal static class Program
    {

        public static string _excelPath = ConfigurationManager.AppSettings["ExcelPath"];
        public static string _sqlconnection = ConfigurationManager.AppSettings["SqlConnection"];
        public static string _sheetName = "";
        public static string _tableName = "";

        public static Form1 form1;
        public static Form2 form2;
        public static Form3 form3;
        public static Form4 form4;

        public static ErrorForm error_form;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);


            Application.ThreadException += new ThreadExceptionEventHandler(ErrorForm.Application_ThreadException);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(ErrorForm.CurrentDomain_UnhandledException);
            


            form1 = new Form1();
            form2 = new Form2();
            form3 = new Form3();
            form4 = new Form4();
            error_form = new ErrorForm();


            

            Application.Run(form1);
        }


        public static void CheckSQLConnection() 
        {
            using (SqlConnection connection = new SqlConnection(Program._sqlconnection))
            {

                try
                {
                    connection.Open();
                }
                catch (Exception ex)
                {
                    throw ex;
                }

            }
        }

        public static (bool,string) run_insert(string query, SqlConnection connection)
        {
            try
            {
                
                SqlCommand command = new SqlCommand(query, connection);
                command.ExecuteNonQuery();
                return (true, "Success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
        public static (bool, string) run_bulk_insert(string query, SqlConnection connection, SqlTransaction transaction)
        {
            try
            {

                SqlCommand command = new SqlCommand(query, connection, transaction);
                command.ExecuteNonQuery();
                return (true, "Success");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

        public static List<string> Get_Excel_Sheets()
        {
            var sheetNames = new List<string>();
            using (var package = new ExcelPackage(new System.IO.FileInfo(Program._excelPath)))
            {
                
                var sheets = package.Workbook.Worksheets;
           
                foreach (var sheet in sheets)
                {
                    sheetNames.Add(sheet.Name);
                }

                
            }
            return sheetNames;

        }

        public static List<string> Get_Tables()
        {
            string connectionString = Program._sqlconnection;
            var tableNames = new List<string>();

            string query = "SELECT name FROM sys.tables WHERE type = 'U'";

           
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataReader reader = command.ExecuteReader())
                    {    
                        while (reader.Read())
                        {
                            string tableName = reader.GetString(0);
                            tableNames.Add(tableName);
                        }
                    }
                }
            }
            return tableNames;

        }


        public static List<string> Get_Excel_Column_Names()
        {
            var columnNames = new List<string>();
            using (var package = new ExcelPackage(new System.IO.FileInfo(Program._excelPath)))
            {

                var sheet = package.Workbook.Worksheets[Program._sheetName];

                ExcelRange firstRow = sheet.Cells["1:1"];

                foreach (var cell in firstRow)
                {
                    columnNames.Add(cell.Value.ToString());

                }
                
            }
            return columnNames;


        }

        public static List<string> Get_Table_Column_Names()
        {

            
            string query = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName";
            var columnNames = new List<string>();

            using (SqlConnection connection = new SqlConnection(Program._sqlconnection))
            {

                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {
        
                    command.Parameters.AddWithValue("@tableName", Program._tableName);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                    
                        while (reader.Read())
                        {
                            string columnName = reader.GetString(0);
                            columnNames.Add(columnName);
                        }
                    }
                    
                }
            }
            return columnNames;

        }

        public static string GetSQLColumnType(string columnName)
        {
            string query = "SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = @tableName AND COLUMN_NAME = @colName";
            string columnType = "";
            using (SqlConnection connection = new SqlConnection(Program._sqlconnection))
            {

                connection.Open();

                using (SqlCommand command = new SqlCommand(query, connection))
                {

                    command.Parameters.AddWithValue("@tableName", Program._tableName);
                    command.Parameters.AddWithValue("@colName", columnName);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {

                        while (reader.Read())
                        {
                            columnType = reader.GetString(0);
                            
                        }
                    }

                }
            }
            return columnType;
        }
        public static string GetColumnType(string columnName, string sqlColType)
        {
            object value = null;
            using (var package = new ExcelPackage(new System.IO.FileInfo(Program._excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[Program._sheetName];
                var idx = worksheet.Cells["1:1"].First(c => c.Value.ToString() == columnName).Start.Column;
                value = worksheet.Cells[2, idx].Value;

                if (value == null)
                {
                    return "Null";
                }
                if (sqlColType.IndexOf("Date", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    try 
                    { 
                        Convert.ToDateTime(worksheet.Cells[2, idx].Text);
                        return "System.Date";
                    } 
                    catch 
                    {
                        return "Null";
                    }
                }
            }
            return value.GetType().ToString();
            
        }
        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.Message, "Unhandled Thread Exception");
        }

        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception).Message, "Unhandled UI Exception");
        }
    }
}
