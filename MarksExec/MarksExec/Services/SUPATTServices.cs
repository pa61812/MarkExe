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
    class SUPATTServices
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
                if (filename.Contains("SUPATT"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(path, now + "SUPATT");
                        outfile = string.Format("{0}{1}{2}", "SUPATT", i, ".txt");
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
            List<SUPATT> Sups = new List<SUPATT>();
            SUPATT Sup = new SUPATT();
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

                        Sup = splitstrng(line);
                        Sups.Add(Sup);
                        line = reader.ReadLine();

                        i++;

                    }

                    reader.Close();
                }

               
                Common.WriteLog("新增DB");
                result = InsertToSql(Sups);
                

                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");
                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("SUPATTServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }

          
        }


        //切割字串
        public static SUPATT splitstrng(string input)
        {
            SUPATT Sup = new SUPATT();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split('|');

                Sup = new SUPATT
                {
                    StoreCode=sArray[0],
                    AttributeClassCode = sArray[1],
                    AttributeClassDescription = sArray[2],
                    AttributeCode = sArray[3],
                    AttributeCodeDescription = sArray[4],
                    AlphanumericValue = sArray[5],
                    NumberValue = sArray[6],
                    Date = sArray[7],
                    Time = sArray[8],
                    StartDate = sArray[9],
                    EndDate = sArray[10]
                };

                return Sup;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return Sup;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<SUPATT> Sup)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);
          
            using (conn)
            {
                conn.Open();

                try
                {
                    var dtSup = Common.ConvertToDataTable<SUPATT>(Sup);


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[SUPATT_TMP]",
                        BatchSize = 1000
                    };
                    bulkCopy.WriteToServer(dtSup);
                    bulkCopy.Close();
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
            string strsql = "TRUNCATE TABLE SUPATT_TMP ";

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
