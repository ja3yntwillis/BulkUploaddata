using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
namespace BulkUploaddata
{
    internal class op
    {
        public static void upload(string q)
        {
            DatabaseConnection dbConnectObj = new DatabaseConnection();
            try
            {
                DatabaseConnection.Connection.Open();

                SqlCommand command = new SqlCommand(q, DatabaseConnection.Connection);
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                DatabaseConnection.Connection.Close();
            }
        }
        public static void formQuery(string schema, string tablename) {
            string filepath = ConfigurationManager.AppSettings["testdatasheet"];
            string sheetname = ConfigurationManager.AppSettings["insertsheetname"];
            string configsheetname = ConfigurationManager.AppSettings["runner"];
            string headers = Excel.GetExcelSheetHeaders(filepath, sheetname);
            string occupied= headers.Substring(1);
            string headerString = "insert into " + schema + "." + tablename + " ("+ occupied+") values({{data}})";
            DataTable wholedateset = Excel.ExcelDataToDataTable(filepath, sheetname, false);
            DataTable datatype = Excel.ExcelDataToDataTable(filepath, configsheetname, false);
            DataTable query=new DataTable();
            query.Columns.Add("Query");
            query.Columns.Add("Status");
            for(int i=0;i<wholedateset.Rows.Count;i++)
            {
                string q1 = "";
                for(int j=0;j<datatype.Columns.Count;j++)
                {
                    if (datatype.Rows[0][j].ToString().ToUpper().Trim() == "TRUE")
                    {
                        if(datatype.Rows[0][j].ToString().ToUpper().Trim() == null || datatype.Rows[0][j].ToString().ToUpper().Trim() == "NULL"|| datatype.Rows[0][j].ToString().ToUpper().Trim() == "")
                        {
                                q1 =q1+wholedateset.Rows[i][j] +",";
                        }
                        else
                        {
                            q1 = q1+"'" + wholedateset.Rows[i][j] + "',";
                        }
                                
                    }
                    else
                    {
                        q1 = q1 + wholedateset.Rows[i][j].ToString() + ",";
                    }
                }
                q1 = q1.Substring(0, q1.Length - 1);
               string headerStringrep = headerString.Replace("{{data}}", q1);
                query.Rows.Add(headerStringrep);
            }

            DatabaseConnection.Connection.Open();
            for (int i = 0; i < query.Rows.Count; i++)
            {
                try
                {
                    if (DatabaseConnection.Connection.State!=ConnectionState.Open)
                    {
                        DatabaseConnection.Connection.Open();
                    }
                    SqlCommand command = new SqlCommand(query.Rows[i]["Query"].ToString(), DatabaseConnection.Connection);
                    command.ExecuteNonQuery();
                    query.Rows[i]["Status"] = "Passed";
                }
                catch (Exception ex)
                {
                    query.Rows[i]["Status"] = "Failed";
                    Console.WriteLine(ex.Message);

                }
                finally
                {
                    DatabaseConnection.Connection.Close();
                }
            }

            Excel.WriteResultDataToASheet("insert", query, schema, tablename);

        }

       

    }

    
}
