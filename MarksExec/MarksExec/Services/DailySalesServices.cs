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
    class DailySalesServices
    {
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            DailySales daily = new DailySales();
            List<DailySales> sales = new List<DailySales>();

            string line;
            try
            {
                //使用FileStream讀取檔案
                FileStream fileStream = File.OpenRead(path);
                StreamReader reader = new StreamReader(fileStream,Encoding.UTF8);
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
                    sales.Add(daily);
                    //Read the next line
                    line = reader.ReadLine();
                    i++;
                }

                //end = DateTime.Now;
                //Console.WriteLine(start.ToString("hh:mm:ss"));
                //Console.WriteLine(end.ToString("hh:mm:ss"));


                reader.Close();
                Common.WriteLog("開始新增");

                 InsertToSql(sales);



                Console.WriteLine("新增 " + finame + " 成功，一共 " + i + " 筆");

               
                return result;
            }
            catch (Exception ex)
            {
                result = false;
                Common.WriteLog(ex.ToString());
                return result;
            }
        }



        private static DailySales ChangString(string filename, string line)
        {
            DailySales dailSales = new DailySales();
            try
            {
                //Common.WriteLog("切割字串 " + line);
                //去除多空白
                line = new Regex("[\\s]+").Replace(line, " ");
                //空白切割
                string[] words = line.Split(' ');


                dailSales = new DailySales
                {
                    StoreCode = filename,
                    DepartmentCode = words[0].Substring(0, 2),
                    ItemCode = words[0].Substring(2, 6),
                    SubCode = words[0].Substring(8, 3),
                    SalesDate = words[0].Substring(11, 10),
                    UnitCode = words[0].Substring(21, 2),
                    StatusPromotion = words[0].Substring(23, 1),
                    SPType = words[0].Substring(24, 1),
                    VatRate = words[1],
                    SalesQty = words[2],
                    SalesAmount = words[3],
                    SalesPrice = words[4],
                    PurchasePrice = words[5],
                    Rebate = words[6],
                    VatAmount = words[7],
                    OriginalSellingPrice = words[8],
                    CSTAmount = words[9]
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

        private static bool InsertToSql(List<DailySales> sasc4)
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
                        string strsql = "Insert into  DailySales_TMP " +
                               "values(@StoreCode,@DepartmentCode,@ItemCode,@SubCode,@SalesDate,@UnitCode,@StatusPromotion"
                               + ",@SPType,@VatRate,@SalesQty,@SalesAmount,@SalesPrice,@PurchasePrice,@Rebate,@VatAmount,@OriginalSellingPrice,@CSTAmount)";


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
