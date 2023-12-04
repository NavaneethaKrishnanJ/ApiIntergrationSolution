using System.Collections.Generic;
using System.Data;
using System;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Configuration;

namespace TestExcel
{
    public  class ExcelData
    {
        private string myFolderPath = string.Empty;
        private string myFileName = string.Empty;
        private DateTime myDate = DateTime.Now;
        private string shortDate = string.Empty;
        private string SQLQuery = string.Empty;
        private string ConnectionString = string.Empty;
        IWorkbook workbook;
        public ExcelData()
        {
            SQLQuery = ConfigurationSettings.AppSettings["SQLQuery"];
            
            workbook = new XSSFWorkbook();
            myFolderPath = Directory.GetCurrentDirectory() + "/ATTData/";
            shortDate = myDate.ToString("yyyyMMdd");
            myFileName = ConfigurationSettings.AppSettings["FileName"].ToString() !=""?
                            ConfigurationSettings.AppSettings["FileName"].ToString() :
                            myFolderPath + shortDate + ".xlsx";

            // Connection string to connect to the SQL Server database
            ConnectionString = ConfigurationSettings.AppSettings["ConnectionString"].ToString() !=""? 
                                ConfigurationSettings.AppSettings["ConnectionString"].ToString():
                               "Data Source=DSS4203021\\SQLEXPRESS;Initial Catalog=CosecDB;Integrated Security=True;TrustServerCertificate=True;Connect Timeout=300";
        }

        public void GenerateExcel()
        {
            try
            {
                string query = SQLQuery + " where EventDateTime > '" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd")+" 09:30:00'"+ " and  EventDateTime < '" + DateTime.Now.ToString("yyyy-MM-dd") + " 09:30:00'";

                // Create a DataTable to store the data fetched from the database
                DataTable dataTable = new DataTable();

                // Connect to the database and fetch the data
                using (SqlConnection connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    SqlCommand command = new SqlCommand(query, connection);
                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    // adapter.SelectCommand.CommandType = CommandType.TableDirect;
                    adapter.Fill(dataTable);
                }

                // Create a new instance of a workbook
                workbook = new XSSFWorkbook();
                // Add a new worksheet to the workbook and specify the name
                ISheet worksheet = workbook.CreateSheet("Attendence");

                // Loop through the columns of the DataTable and add them to the Excel worksheet
                IRow headerRow = worksheet.CreateRow(0);

                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    headerRow.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
                }

                // Loop through the rows of the DataTable and add them to the Excel worksheet
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    IRow dataRow = worksheet.CreateRow(i + 1);
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        if (j < 5)
                        {
                            dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                        }
                        else
                        {
                            IDataFormat dataFormatCustom = workbook.CreateDataFormat();
                            var cell = dataRow.CreateCell(j);
                            var data = dataTable.Rows[i][j].ToString();
                            data = data == "" ? "NULL" : data;
                            string result;
                            if (!string.IsNullOrEmpty(data) && j == 5)
                            {
                                cell.SetCellValue(Convert.ToDateTime(data));
                            }
                            else if (data != "NULL" && (j == 6 || j == 7))
                            {
                                cell.SetCellValue(Convert.ToDateTime(data).Hour + ":" + Convert.ToDateTime(data).Minute + ":" + Convert.ToDateTime(data).Second);
                            }
                            else
                            {
                                cell.SetCellValue(data);
                            }



                        }
                    }

                }

                // Write the workbook to a file
                using (FileStream stream = new FileStream(myFileName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream, false);
                }
            }
            catch (Exception e) { Console.WriteLine(e.Message); }
        }

        
       
      
        
    }
}
