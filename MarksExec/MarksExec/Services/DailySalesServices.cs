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
    class DailySalesServices
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
                //解壓路徑
                string outpath = "";
                string finame = Common.GetNameCode(filename);

                if (finame.Length == 12 && Common.GetLastCode(finame) == "L")
                {
                    if (Path.GetExtension(filename).Contains("zip"))
                    {
                        outpath = Path.Combine(path, now + "L");
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
                            Issuccess=Common.MoveFile(filepath, OutToPath, filename);
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



                Common.WriteLog("新增 " + finame + " 成功，一共 " + i + " 筆");

               
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

                try
                {
                    DapperPlusManager.Entity<DailySales>().Table("DailySales_TMP");
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

    }
}
