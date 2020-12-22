using Dapper;
using MarksExec.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Z.Dapper.Plus;

namespace MarksExec.Services
{

    class DailyStockServices
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

            //刪除DB，成功在進行Insert
            Common.WriteLog("刪除DB");
            Issuccess = DeleteTable();

            if (!Issuccess)
            {
                return;
            }
            foreach (var item in di.GetFiles())
            {
                i++;
                //檔案名稱
                string filename = item.Name;
                //檔案路徑
                string filepath = item.FullName;
                //解壓路徑
                string outpath = "";
                string finame = Common.GetNameCode(filename);

                if (finame.Length == 12 && Common.GetLastCode(finame) == "A")
                {
                    if (Path.GetExtension(filename).Contains("zip"))
                    {
                        outpath = Path.Combine(path, now + "A");
                        Common.CreateLocation(outpath, "");
                        //解壓縮
                        Issuccess = Common.UnZipToFile(filepath, outpath);
                        if (Issuccess)
                        {
                            Issuccess = StartInsert(Path.Combine(outpath, finame), finame);
                        }
                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Issuccess = Common.MoveFile(filepath, OutToPath, filename);
                        }
                        if (Issuccess)
                        {
                            Common.DeleteFolder(outpath);
                        }
                        continue;
                    }

                    StartInsert(filepath, filename);
                    continue;
                }
            }
        }
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;


            DailyStock daily = new DailyStock();
            List<DailyStock> stocks = new List<DailyStock>();

            string line;
            try
            {
                //使用FileStream讀取檔案
                FileStream fileStream = File.OpenRead(path);
                StreamReader reader = new StreamReader(fileStream, Encoding.UTF8);
                string filename = Path.GetFileName(finame).Substring(0, 3);
                int i = 0;
                line = reader.ReadLine();
                Common.WriteLog("切割字串");


                //Console.WriteLine("切割字串");
                //start = DateTime.Now;

                while (line != null )
                {
                    line = new Regex("[\\s]+").Replace(line, " ");
                    if (line == " " || line == "")
                    {
                        line = reader.ReadLine();
                        continue;
                    }
                    daily = ChangString(filename, line);
                    stocks.Add(daily);
                    //Read the next line
                    line = reader.ReadLine();
                    i++;
                }

                //end = DateTime.Now;
                //Console.WriteLine(start.ToString("hh:mm:ss"));
                //Console.WriteLine(end.ToString("hh:mm:ss"));


                reader.Close();
               
                Common.WriteLog("新增DB");
                result = InsertToSql(stocks);
                



                Common.WriteLog("新增 " + finame + " 成功，一共 " + i + " 筆");

                return result;
     
            }
            catch (Exception ex)
            {

                Common.WriteLog(ex.ToString());
                result = false;
                return result;
            }
        }



        private static DailyStock ChangString(string filename, string line)
        {
            DailyStock dailSales = new DailyStock();
            try
            {
                //Common.WriteLog("切割字串 " + line);
                //去除多空白
                line = new Regex("[\\s]+").Replace(line, " ");
                //空白切割
                string[] words = line.Split(' ');
                string words22 = "";
                if (words[22].Length==8)
                {
                     words22 = words[22].Substring(4, 4);
                }
                


                dailSales = new DailyStock
                {
                    StoreCode = filename,
                    DepartmentCode = words[0].Substring(0, 2),
                    ItemCode = words[0].Substring(2, 6),
                    SubCode = words[0].Substring(8, 3),
                    TransactionDate = words[0].Substring(11, 10),
                    SalesStatus = words[0].Substring(21, 1),
                    BalanceQty = words[1],
                    Normal_OrderQty = words[2],
                    Promotion_OrderQty = words[3],
                    Normal_FreeGoods = words[4],
                    Promotion_FreeGoods = words[5],
                    Normal_ReceivedQty = words[6],
                    Promotion_ReceivedQty = words[7],
                    Normal_ReturnedQty = words[8],
                    Promotion_ReturnedQty = words[9],
                    SalesQty = words[10],
                    SalesAmount = words[11],
                    PurchaseAmount = words[12],
                    AdjustedQty = words[13],
                    BalanceForward = words[14],
                    Balance_Forward_OrderNotYetReceivedBasicgoods = words[15],
                    Balance_Forward_OrderNotYetReceivedFreegoods = words[16],
                    Rebate = words[17],
                    AvgCost = words[18],
                    AvgAmount = words[19],
                    BFAvgAmount = words[20],
                    PurchasePrice = words[21],
                    MainSuppliercode = words[22].Substring(0, 4),
                    DSSuppliercode = words22,
                    StopSubmonth = " ",
                    StopSubyear = " ",
                    StopSubreason = " "
                };
                return dailSales;
            }
            catch (Exception ex)
            {
               
                Common.WriteLog("切割字串 " + line + " 失敗");

                Common.WriteLog(ex.ToString());
                return dailSales;
            }



        }

        private static bool InsertToSql(List<DailyStock> sasc4)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);
           
            using (conn)
            {
                conn.Open();

                try
                {
                    DapperPlusManager.Entity<DailyStock>().Table("DailyStock_TMP");
                    conn.BulkInsert(sasc4);
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

        //進DB前先刪除
        public static bool DeleteTable()
        {

            bool result = true;
            string strsql = "delete from DailyStock_TMP ";

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);

            using (conn)
            {
                conn.Open();
               
                    try
                    {
                        SqlCommand cmd = new SqlCommand(strsql, conn);
                        cmd.ExecuteNonQuery();
                        conn.Close();
  
                        Common.WriteLog("刪除成功");
                        return result;

                    }
                    catch (Exception e)
                    {
                      
                        Common.WriteLog("刪除失敗");
                        Common.WriteLog(e.ToString());
                        conn.Close();
                        result = false;
                        return result;
                    }
                
            }

        }
    }
}
