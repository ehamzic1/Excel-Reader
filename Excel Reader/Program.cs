using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;
using System.Data.SQLite;
using OfficeOpenXml;

namespace Excel_Reader
{
    internal class Program
    {        
        public static string dbPath { get; set; } = @"C:\Sqlite databases\Excel_database.db";
        public static string connectionString { get; set; } = @"Data Source = C:\Sqlite databases\Excel_database.db";
        static void Main(string[] args)
        {
            // check if the db exists
            // create a new one if it does not
            // if it does - delete it and create a new one
            SQLiteConnection.CreateFile(dbPath);
            CreateDbTable();

            // read from excel
            string excelPath = @"C:\Users\Ervin.Hamzic\Desktop\Employees.xlsx";
            using (ExcelPackage package = new ExcelPackage(excelPath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets["Employees"];
                Program p1 = new Program();
                var employees = p1.GetList<Employee>(sheet);
                // save to the db
                p1.StoreInDb(employees);
            }
            
            
            
            // read from the db 
            // present on the console
        }
        private void StoreInDb(List<Employee> employees)
        {
            using(SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using(SQLiteCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"INSERT INTO Employees(Id, Firstname, Lastname, StartDate, Rank) 
                                        VALUES(@id, @firstname, @lastname, @startdate, @rank)";
                    foreach (Employee employee in employees)
                    {
                        cmd.Parameters.AddWithValue("@id", employee.EmployeeId);
                        cmd.Parameters.AddWithValue("@firstname", employee.Firstname);
                        cmd.Parameters.AddWithValue("@lastname", employee.Lastname);
                        cmd.Parameters.AddWithValue("@startdate", employee.JoinDate.ToString("dd.MM.yyyy"));
                        cmd.Parameters.AddWithValue("@rank", employee.Rank);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        // create all tables 
        static void CreateDbTable()
        {
            SQLiteConnection connection = new SQLiteConnection($"Data Source={dbPath}; Version=3;");
            connection.Open();
            using(var tableCmd = connection.CreateCommand())
            {
                tableCmd.CommandText = @"CREATE TABLE Employees (Id INTEGER NOT NULL UNIQUE,
                            Firstname TEXT NOT NULL,
                            Lastname TEXT NOT NULL,
                            StartDate TEXT NOT NULL,
                            Rank TEXT NOT NULL,
                            PRIMARY KEY(Id AUTOINCREMENT))";

                tableCmd.ExecuteNonQuery();
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> list = new List<T>();
            
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                new {Index=n, ColumnName=sheet.Cells[1,n].Value.ToString()}
            );
            for(int row=3; row<sheet.Dimension.Rows; row++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach(var prop in typeof(T).GetProperties())
                {
                    int col = columnInfo.SingleOrDefault(c => c.ColumnName == prop.Name).Index;
                    var val = sheet.Cells[row, col].Value;
                    var propType = prop.PropertyType;
                    prop.SetValue(obj, Convert.ChangeType(val, propType));
                }
                list.Add(obj);
            }
            return list;

        }


        // 3 parts of the app
        //1. db creation
        //2. read from the excel file and return a list
        //3. populate the db
    }

    internal class Employee
    {
        public int EmployeeId { get; set; }
        public string Firstname { get; set; }
        public string Lastname { get; set; }
        public DateTime JoinDate { get; set; }
        public string Rank { get; set; }
    }
}
