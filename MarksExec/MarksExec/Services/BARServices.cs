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
    class BARServices
    {
        //抓指定路徑的符合檔名的檔案
        /// <summary>
        ///抓指定路徑的符合檔名的檔案
        /// </summary>
        /// <param name="path">指定路徑</param>
        /// <param name="OutToPath">匯出路徑</param>
        public static void GetFile(string path,string OutToPath)
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
                //解壓檔案名稱(gz檔才需要)
                string outfile = "";
                if (filename.Contains("BAR"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(path, now + "BAR");
                        outfile = string.Format("{0}{1}{2}", "BAR", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        if (Issuccess)
                        {
                            Issuccess = StartInsert(Path.Combine(outpath, outfile), filename);
                        }


                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Issuccess=Common.MoveFile(filepath, OutToPath, filename);
                        }
                        //刪除檔案
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




        //path 完整路徑  finame檔名
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<BAR> bars = new List<BAR>();
            BAR bar = new BAR();
            try
            {
                Common.WriteLog("開始切割 " + finame);
                FileStream fileStream = File.OpenRead(path);
                using (StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(0)))
                {
                    line = reader.ReadLine();
                    while (line != null)
                    {
                        line = new Regex("[\\s]+").Replace(line, " ");
                        if (line == " " || line == "")
                        {
                            line = reader.ReadLine();
                            continue;
                        }
                        bar = splitstrng(line);
                        bars.Add(bar);
                        line = reader.ReadLine();
                        i++;

                    }

                    reader.Close();
                }

                Common.WriteLog("新增DB");
                result = InsertToSql(bars);

                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");

                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("BARServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }

            
        }


        //切割字串
        public static BAR splitstrng(string input)
        {
            BAR bar = new BAR();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split('|');

                bar = new BAR
                {
                    Dept=sArray[0],
                    ItemCode = sArray[1],
                    SubCode = sArray[2],
                    UnitCode = sArray[3],
                    Internal_ExternalBarcode = sArray[4],
                    EAN13Barcode = sArray[5],
                    OriginalBarcodeType = sArray[6],
                    OriginalBarcodeNo = sArray[7],
                    SNISpecialTypeBarcode = sArray[8],
                    MajorBarcode = sArray[9],
                    StartStopDate = sArray[10],
                    EndStopDate = sArray[11],
                    StopItemReason = sArray[12],
                    StopReasonDescription = sArray[13],
                    ItemID = sArray[14],
                    ItemSubcodeID = sArray[15],
                    Length = sArray[16],
                    Width = sArray[17],
                    Height = sArray[18],
                    NotForSellFlag = sArray[19],
                    StoreCode = sArray[20],
                    NotForSellReason = sArray[21],
                    ItemCodeEnglishDescription = sArray[22],
                    ItemCodeChineseDescription = sArray[23],
                    Status = sArray[24]


                };

                return bar;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return bar;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<BAR> bars)
        {
            bool result = true;

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);
          

            using (conn)
            {
                conn.Open();            
                try
                {
                        DapperPlusManager.Entity<BAR>().Table("BARCODE_TMP");
                        conn.BulkInsert(bars);
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
