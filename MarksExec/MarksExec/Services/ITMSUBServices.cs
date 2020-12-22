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
    class ITMSUBServices
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
                //解壓檔案名稱(gz檔才需要)
                string outfile = "";
                if (filename.Contains("ITMSUB"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(path, now + "ITMSUB");
                        outfile = string.Format("{0}{1}{2}", "ITMSUB", i, ".txt");
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
                            Issuccess = Common.MoveFile(filepath, OutToPath, filename);
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

        //path 路徑  finame檔名
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<ITMSUB> iTMSUBs = new List<ITMSUB>();
            ITMSUB iTMSUB = new ITMSUB();
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

                        iTMSUB = splitstrng(line);
                        iTMSUBs.Add(iTMSUB);
                        line = reader.ReadLine();

                        i++;

                    }

                    reader.Close();
                }

             
                Common.WriteLog("新增DB");
                result = InsertToSql(iTMSUBs);
                

                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");
                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("ITMSUBServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
               
            }

           

        }


        //切割字串
        public static ITMSUB splitstrng(string input)
        {
            ITMSUB iTMSUB = new ITMSUB();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split('|');

                iTMSUB = new ITMSUB
                {
                    Dept=sArray[0],
                    ItemCode = sArray[1],
                    SubID = sArray[2],
                    SubCode = sArray[3],
                    ItemID = sArray[4],
                    SubCodeEDes = sArray[5],
                    SubCodeCDes = sArray[6],
                    StartStopDate = sArray[7],
                    StopItemReasonCode = sArray[8],
                    ProductCharacteristic = sArray[9],
                    EndStopDate = sArray[10],
                    ReturnableforNonStopItem = sArray[11],
                    ReturnableforTemporaryStop = sArray[12],
                    ReturnableforPermanentStop = sArray[13],
                    WHCheckingMark = sArray[14],
                    ItemLevel = sArray[15],
                    UsebyDate = sArray[16],
                    CNCode = sArray[17]

                };

                return iTMSUB;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return iTMSUB;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<ITMSUB> iTMSUBs)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);
           
            using (conn)
            {
                conn.Open();

                try
                {
                    DapperPlusManager.Entity<ITMSUB>().Table("ITMSUB_TMP");
                    conn.BulkInsert(iTMSUBs);
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
            string strsql = "delete from ITMSUB_TMP ";

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);

            using (conn)
            {
                conn.Open();
                //加上BeginTrans
               
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

