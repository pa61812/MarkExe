using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace MarksExec
{
    public static  class Common
    {
        #region 常用
        public static void WriteLog(String logMsg)
        {
            //Log目錄
            String logPath = ConfigurationManager.AppSettings["LogLocation"];
            logPath = CheckLocation(logPath);
            //Log檔名
            String logFileName = DateTime.Now.ToString("yyyyMMdd") + "_Log.log";
            //Log檔
            String logFile = logPath + logFileName;

            CreateLocation(logPath, logFile);

            try
            {

                using (StreamWriter sw = new StreamWriter(logFile, true))
                {
                    sw.WriteLine("--執行時間 " + DateTime.Now.ToString("yyyy/MM/dd hh:mm:ss") + "--");
                    sw.WriteLine(logMsg);
                    sw.Close();
                }

            }
            catch (Exception ex)
            {
                ex.ToString();
                throw;
            }
        }
        /// <summary>
        /// 確認路徑是否最後為\
        /// </summary>
        /// <param name="str">要檢查的路徑</param>
        public static string CheckLocation(string str)
        {

            string _str = str.Substring(str.Length - 1, 1);
            if (_str != "\\")
            {
                str = str + "\\";
            }
            return str;

        }

        /// <summary>
        /// 確認路徑是否存在
        /// </summary>
        /// <param name="FilePath">要檢查的路徑</param>
        /// <param name="FilePath">要檢查的檔案(可為空)</param>
        public static void CreateLocation(String FilePath, string FileName)
        {
            //資料夾若不在
            if (!Directory.Exists(FilePath))
            {
                //建立資料夾
                Directory.CreateDirectory(FilePath);
            }
            //檔若不在
            if (!File.Exists(FileName) && FileName != "")
            {
                //建立檔案
                File.Create(FileName).Close();
            }
        }


        //Out路徑
        public static string OutLocation()
        {
            //寫檔目錄
            string filelocation = ConfigurationManager.AppSettings["FileLocation"];
            string now = DateTime.Now.ToString("yyyyMMdd");
            filelocation = Path.Combine(filelocation, now);
            Common.CreateLocation(filelocation, "");
            //分類
            string[] dirs = Directory.GetDirectories(filelocation);


            filelocation = Path.Combine(filelocation, (dirs.Length + 1).ToString());
            Common.CreateLocation(filelocation, "");

            //決定目錄
            filelocation = Common.CheckLocation(filelocation);
            return filelocation;
        }

        //抓最後一個字
        public static string GetLastCode(string _str)
        {
            _str = _str.Substring(_str.Length - 1, 1);
            return _str;
        }


        //抓檔名
        public static string GetNameCode(string _str)
        {
            _str = _str.Substring(0, _str.Length - 4);
            return _str;
        }

        #endregion

        #region 壓縮檔處理
        /// <summary>
        ///解壓GZ縮檔
        /// </summary>
        /// <param name="GzPath">要解壓的檔案</param>
        /// <param name="outpath">解壓的路徑</param>
        /// <param name="Outfile">解壓後完整路徑</param>
        public static bool UnGZToFile(string GzPath,string outpath, string Outfile)
        {
            

            int count = 0;
            int bufferSize = 4096;
            byte[] buffer = new byte[4096];
            bool result = true;
            //資料夾若在
            if (Directory.Exists(Outfile))
            {
                DeleteFolder(Outfile);
            }
            CreateLocation(outpath, Outfile);
            try
            {
                //讀取壓縮的 *.gz 的檔案
                using (FileStream GzFile = new FileStream(GzPath, FileMode.Open, FileAccess.Read, FileShare.Read))

                //建立一個解壓縮後的檔案路徑與名稱
                using (FileStream GzOutFile = new FileStream(Outfile, FileMode.Create, FileAccess.Write, FileShare.None))

                //解壓縮 *.gz 的檔案
                using (GZipStream gz = new GZipStream(GzFile, CompressionMode.Decompress, true))

                    //讀取 *.gz 中的壓縮檔內容，並且寫入到新建立的檔案當中
                    //直到內容結束為止
                    while (true)
                    {
                        count = gz.Read(buffer, 0, bufferSize);

                        if (count != 0)
                            GzOutFile.Write(buffer, 0, count);

                        if (count != bufferSize)
                            break;
                    }

                //if (File.Exists(GzPath))
                //{
                //    File.Delete(GzPath);
                //}
                return result;
            }
            catch (Exception ex)
            {
                //如果檔案已產生，刪除
                if (File.Exists(Outfile))
                {
                    File.Delete(Outfile);
                }
                WriteLog(ex.ToString());
                result = false;
                return result;
                
            }


        }
        /// <summary>
        ///解壓ZIP縮檔
        /// </summary>
        /// <param name="filepath">要解壓的檔案</param>
        /// <param name="outpath">解壓後的路徑</param>
        public static bool UnZipToFile(string filepath,string outpath)
        {

            bool result = true;
            try
            {
                //資料夾若在
                if (Directory.Exists(outpath))
                {
                    DeleteFolder(outpath);
                }
                //解壓縮
                ZipFile.ExtractToDirectory(filepath, outpath);
                //File.Delete(filepath);
                return result; 
            }
            catch (Exception ex)
            {
                //if (File.Exists(outpath))
                //{
                //    File.Delete(outpath);
                //}
                WriteLog(ex.ToString());
                result = false;
                return result;
               
            }
          
          
        }

      
        #endregion

        #region 移檔刪檔
        //將已匯入檔案寫至FileLocation
        /// <summary>
        /// 將檔案搬至OUT
        /// </summary>
        /// <param name="filepath">要搬檔檔案</param>
        /// <param name="outpath">搬檔後路徑</param>
        /// <param name="filename">檔名</param>
        public static bool MoveFile(string filepath, string outpath, string filename)
        {
            bool result = true;
            string path = Path.Combine(outpath, filename);

            CreateLocation(outpath, path);

            try
            {

                File.Move(filepath, path,true);

                //刪除已匯入檔案
                if (File.Exists(filepath))
                {
                    File.Delete(filepath);
                }
                WriteLog("移至FileLocation成功");
                return result;
            }
            catch (Exception ex)
            {
                result = false;
                WriteLog("移至FileLocation失敗");
                WriteLog(ex.ToString());
                return result;
            }
        }

        //刪除資料夾
        public static void DeleteFolder(string dir)
        {
            WriteLog("檔案已存在，DeleteFolder");
            try
            {
                foreach (string di in Directory.GetFileSystemEntries(dir))
                {
                    if (File.Exists(di))
                    {
                        FileInfo fi = new FileInfo(di);
                        if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                        {
                            fi.Attributes = FileAttributes.Normal;
                        }
                        File.Delete(di);//直接删除其中的文件   
                    }
                    else
                        DeleteFolder(di);//删除子資料夹夾  
                }
                Directory.Delete(dir);//删除已空文件夾  
            }
            catch (Exception ex)
            {
                WriteLog("DeleteFolder失敗");
                WriteLog(ex.ToString());
                throw;
            }


        }
        //public static void WriteFile(string filepath, string outpath, string filename)
        //{

        //    string path = Path.Combine(outpath, filename);

        //    CreateLocation(outpath, path);

        //    try
        //    {

        //        byte[] ansiBytes = File.ReadAllBytes(filepath);
        //        var utf8String = Encoding.GetEncoding(0).GetString(ansiBytes);
        //        File.WriteAllText(path, utf8String, Encoding.UTF8);

        //        //刪除已匯入檔案
        //        if (File.Exists(filepath))
        //        {
        //            File.Delete(filepath);
        //        }
        //        WriteLog("移至FileLocation成功");
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteLog(ex.ToString());

        //    }
        //}
        #endregion




    }
}
