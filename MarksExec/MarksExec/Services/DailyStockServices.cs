using Dapper;
using MarksExec.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MarksExec.Services
{
    class DailyStockServices
    {
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
                Common.WriteLog("開始新增");



                result= InsertToSql(stocks);


                Console.WriteLine("新增 " + finame + " 成功，一共 " + i + " 筆");

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
                Console.WriteLine("切割字串 " + line + " 失敗");
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
                //加上BeginTrans
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string strsql = "Insert into  DailyStock_TMP " +
                               "values(@StoreCode,@DepartmentCode,@ItemCode,@SubCode"
                             + ", @TransactionDate, @SalesStatus, @BalanceQty"
                             + ", @Normal_OrderQty, @Promotion_OrderQty, @Normal_FreeGoods"
                             + ", @Promotion_FreeGoods, @Normal_ReceivedQty, @Promotion_ReceivedQty"
                             + ", @Normal_ReturnedQty, @Promotion_ReturnedQty, @SalesQty"
                             + ", @SalesAmount, @PurchaseAmount, @AdjustedQty"
                             + ", @BalanceForward, @Balance_Forward_OrderNotYetReceivedBasicgoods"
                             + ", @Balance_Forward_OrderNotYetReceivedFreegoods, @Rebate"
                             + ", @AvgCost, @AvgAmount, @BFAvgAmount, @PurchasePrice"
                             + ", @MainSuppliercode, @DSSuppliercode, @StopSubmonth"
                             + ", @StopSubyear, @StopSubreason)";


                        conn.Execute(strsql, sasc4, transaction);


                        //正確就Commit
                        transaction.Commit();
                        conn.Close();
                        Common.WriteLog("新增成功");
                        
                        return result;
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
                        conn.Close();
                        Console.WriteLine("新增失敗");
                        Common.WriteLog("新增失敗");
                        Common.WriteLog(e.ToString());
                        result = false;
                        return result;

                    }
                }
            }
        }
    }
}
