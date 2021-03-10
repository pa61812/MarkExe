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
    class SampleServices
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
            //使否成功
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

            //在指定路徑搜尋
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
                //檔名有對在坐下去
                if (filename.Contains("Sample"))
                {
                    //若副檔名為GZ
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        //解壓目錄
                        outpath = Path.Combine(path, now + "Sample");
                        //解壓致XXX.txt檔
                        outfile = string.Format("{0}{1}{2}", "Sample", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        //如果成功在讀取
                        if (Issuccess)
                        {
                            Issuccess = StartInsert(Path.Combine(outpath, outfile), filename);
                        }

                        //進DB成功後搬檔
                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Issuccess = Common.MoveFile(filepath, OutToPath, filename);
                        }
                        //刪除解壓縮的檔案
                        if (Issuccess)
                        {
                            Common.DeleteFolder(outpath);
                        }

                        continue;
                    }

                    //若副檔名為ZIP
                    if (Path.GetExtension(filename).Contains("zip"))
                    {
                        //解壓目錄
                        outpath = Path.Combine(path, now + "Sample");
                        //建立解壓目錄
                        Common.CreateLocation(outpath, "");
                        //解壓縮
                        Issuccess = Common.UnZipToFile(filepath, outpath);
                        //如果解壓成功
                        if (Issuccess)
                        {
                            //將解壓出來檔案進行讀取。(預設解壓後檔名=壓縮檔名)
                            Issuccess = StartInsert(Path.Combine(outpath, filename), filename);
                        }
                        //進DB成功後搬檔
                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Issuccess = Common.MoveFile(filepath, OutToPath, filename);
                        }
                        //刪除解壓縮的檔案
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



        //開始新增
        //path 完整路徑  finame檔名
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<Sample> samples = new List<Sample>();
            Sample sample = new Sample();
            try
            {
                Common.WriteLog("開始切割 " + finame);
                //開始讀取檔案
                FileStream fileStream = File.OpenRead(path);
                using (StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(0)))
                {
                    line = reader.ReadLine();
                    //!= null再繼續做下去
                    while (line != null)
                    {
                        line = new Regex("[\\s]+").Replace(line, " ");
                        if (line == " " || line == "")
                        {
                            line = reader.ReadLine();
                            continue;
                        }
                        //切割完後
                        sample = splitstrng(line);
                        //進LIST
                        samples.Add(sample);
                        //讀下一行
                        line = reader.ReadLine();
                        i++;

                    }

                    reader.Close();
                }




                Common.WriteLog("新增DB");
                //進DB
                result = InsertToSql(samples);



                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");

                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("SampleServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }


        }


        //切割字串
        public static Sample splitstrng(string input)
        {
            Sample sample = new Sample();
            try
            {
                 //去除多餘的空白    
                input = new Regex("[\\s]+").Replace(input, " ");
                //切割的類型  
                string[] sArray = input.Split('|');

                sample = new Sample
                {
                    SampleID=sArray[0],

                    SampleText=sArray[1]
                    //...
                    //...
                    //...
                    //無限下加


                };

                return sample;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return sample;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<Sample> samples)
        {
            bool result = true;

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);


            using (conn)
            {
                conn.Open();
                try
                {
                    var ec = Common.ConvertToDataTable<Sample>(samples);


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[欲新增的Table]",
                        BatchSize = 1000
                    };
                    bulkCopy.WriteToServer(ec);
                    bulkCopy.Close();
                    //DapperPlusManager.Entity<Sample>().Table("欲新增的Table");
                    //conn.BulkInsert(samples);
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
            string strsql = "truncate table 欲刪除的Table ";

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
