using MarksExec.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Dapper;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using Z.Dapper.Plus;

namespace MarksExec.Services
{
    class BigSalesServices
    {
        //抓指定路徑的符合檔名的檔案
        /// <summary>
        ///抓指定路徑的符合檔名的檔案
        /// </summary>
        /// <param name="path">指定路徑</param>
        /// <param name="OutToPath">匯出路徑</param>
        public static void GetFile(string path, string OutToPath)
        {
            //找檔案
            DirectoryInfo di = new DirectoryInfo(path);
            bool Issuccess;
            int i = 0;
            string now = DateTime.Now.ToString("yyyyMMdd");

            foreach (var item in di.GetFiles())
            {
                i++;
                //檔案名稱
                string filename = item.Name;
                //檔案路徑
                string filepath = item.FullName;
               
                if (filename.Contains("BigSales"))
                {
                    Issuccess=StartInsert(filepath, filename);
                    if (Issuccess)
                    {
                        File.Move(filepath, Path.Combine(OutToPath, filename), true);
                    }
            
                }
            }
        }
        //開始新增
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            string sql = "select * from C6_BigSales";
            try
            {
                Common.WriteLog("開始讀取 " + finame);
                DataTable dt = GetOleDbDataTable(path, sql);

                var bigSales=  DataTableToList(dt);
                result= InsertToSql(bigSales);

                Common.WriteLog(finame + "讀取完成，一共 " + bigSales.Count + " 筆");
                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("BigSalesServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;

            }



        }


        //連線ACCESS參考項目
        public static OleDbConnection OleDbOpenConn(string Database)
        {
            string cnstr = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Database);
            OleDbConnection icn = new OleDbConnection();
            icn.ConnectionString = cnstr;
            if (icn.State == ConnectionState.Open) icn.Close();
            icn.Open();
            return icn;
        }


        //連線ACCESS，取得資料表:

        public static DataTable GetOleDbDataTable(string Database, string OleDbString)
        {
            DataTable myDataTable = new DataTable();
            OleDbConnection icn = OleDbOpenConn(Database);
            OleDbDataAdapter da = new OleDbDataAdapter(OleDbString, icn);
            DataSet ds = new DataSet();
            ds.Clear();
            da.Fill(ds);
            myDataTable = ds.Tables[0];
            if (icn.State == ConnectionState.Open) icn.Close();
            return myDataTable;
        }



        //進DB
        private static bool InsertToSql(List<BigSales> bigSales)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);
          
            using (conn)
            {
                conn.Open();
             
                    try
                    {
                        DapperPlusManager.Entity<BigSales>().Table("BigSales_TMP");
                        conn.BulkInsert(bigSales);
                        conn.Close();
                        Common.WriteLog("新增成功");
                        return result;
                    }
                    catch (Exception e)
                    {
                      
                        Common.WriteLog("新增失敗");
                        Common.WriteLog(e.ToString());
                        conn.Close();
                        result = false;
                        return result;
                    }              
            }

        }




        public static List<BigSales> DataTableToList(DataTable dt)
        {
            List<BigSales> bigSales = new List<BigSales>();
            foreach (DataRow dr in dt.Rows)
            {
                BigSales doc = new BigSales();
                doc.YYYYMM = dr.Field<string>("YYYYMM");
                doc.Business_Date = dr.Field<DateTime>("Business_Date").ToString("yyyy/MM/dd");
                doc.StoreName = dr.Field<string>("StoreName");
                doc.FullCode = dr.Field<string>("FullCode");
                doc.Item_Sales_Qty_Unit = dr.Field<Int32>("Item_Sales_Qty_Unit").ToString();
                doc.Item_Sales_Amount = dr.Field<Int32>("Item_Sales_Amount").ToString();
                doc.ItemCName = dr.Field<string>("ItemCName");
                doc.Input = dr.Field<string>("Input");
                bigSales.Add(doc);
            }
            return bigSales;
        }
    }
}
