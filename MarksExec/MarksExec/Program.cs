using MarksExec.Services;
using System;
using System.Configuration;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace MarksExec
{
    class Program
    {     
        static void Main(string[] args)
        {

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


            //吃檔目錄
            string filelocation = ConfigurationManager.AppSettings["FileInPut"];
            //Out路徑
            string OutToPath = OutLocation();
            string now = DateTime.Now.ToString("yyyyMMdd");
            int i = 1;
            //是否成功
            bool Issuccess = true;

            #region Master
            string filemaster = Path.Combine(filelocation, "Master");

            DirectoryInfo di = new DirectoryInfo(filemaster);
            
            foreach (var item in di.GetFiles())
            {
                //檔案名稱
                string filename = item.Name;
                //檔案路徑
                string filepath = item.FullName;
                //解壓路徑
                string outpath = "";
                //解壓檔案名稱(gz檔才需要)
                string outfile = "";

                if (filename.Contains("ITMGEL"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(filelocation, now + "ITMGEL");
                        outfile = string.Format("{0}{1}{2}", "ITMGEL", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, outfile);
                            Issuccess = ITMGELServices.StartInsert(outpath, filename);
                        }
                       

                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, outfile);
                        }
                        

                        continue;
                    }
                    ITMGELServices.StartInsert(filepath, filename);
                    continue;
                }

                if (filename.Contains("STM"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(filelocation, now + "STM");
                        outfile = string.Format("{0}{1}{2}", "STM", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, outfile);
                            Issuccess = STMTMPServices.StartInsert(outpath, filename);
                        }
                       

                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, outfile);
                        }

                        continue;
                    }
                    STMTMPServices.StartInsert(filepath, filename);
                    continue;
                }

                if (filename.Contains("SUPATT"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(filelocation, now + "SUPATT");
                        outfile = string.Format("{0}{1}{2}", "SUPATT", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, outfile);
                            Issuccess = SUPATTServices.StartInsert(outpath, filename);
                        }
                     

                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, outfile);
                        }

                        continue;
                    }
                    SUPATTServices.StartInsert(filepath, filename);
                    continue;
                }

                if (filename.Contains("ITMSUB"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(filelocation, now + "ITMSUB");
                        outfile = string.Format("{0}{1}{2}", "ITMSUB", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, outfile);
                            Issuccess = ITMSUBServices.StartInsert(outpath, filename);
                        }
                       

                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, outfile);
                        }

                        continue;
                    }

                    ITMSUBServices.StartInsert(filepath, filename);
                    continue;
                }

                if (filename.Contains("ITMSUPGEXCRTNC"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(filelocation, now + "ITMSUPGEXCRTNC");
                        outfile = string.Format("{0}{1}{2}", "ITMSUPGEXCRTNC", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, outfile);
                            Issuccess = ITMSUPGEXCRTNCServices.StartInsert(outpath, filename);
                        }
                       
                 
                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, outfile);
                        }

                        continue;
                    }
                    ITMSUPGEXCRTNCServices.StartInsert(filepath, filename);
                    continue;
                }

                if (filename.Contains("BAR"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(filelocation, now + "BAR");
                        outfile = string.Format("{0}{1}{2}", "BAR", i, ".txt");
                        //解壓縮
                        Issuccess = Common.UnGZToFile(filepath, outpath, Path.Combine(outpath, outfile));                       
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, outfile);
                            Issuccess = BARServices.StartInsert(outpath, filename);
                        }
                      

                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, outfile);
                        }

                        continue;
                    }
                    BARServices.StartInsert(filepath, filename);
                    continue;
                }

                #region BigSales

                if (filename.Contains("BigSales"))
                {
                    BigSalesServices.StartInsert(filepath, filename);
                    File.Move(filepath,Path.Combine(OutToPath, filename));
                }

                #endregion



                i++;
            }
            #endregion

            #region Sales
            string filesales = Path.Combine(filelocation, "Sales");
            di = new DirectoryInfo(filesales);

            //sales = L

            foreach (var item in di.GetFiles())
            {
                //檔案名稱
                string filename = item.Name;
                //檔案路徑
                string filepath = item.FullName;
                //解壓路徑
                string outpath = "";

                string finame = GetNameCode(filename);
                if (finame.Length == 12 && GetLastCode(finame) == "L")
                {
                    if (Path.GetExtension(filename).Contains("zip"))
                    {
                        outpath = Path.Combine(filelocation, now + "L");
                        Common.CreateLocation(outpath, "");
                        //解壓縮
                        Issuccess = Common.UnZipToFile(filepath, outpath);
                        if (Issuccess)
                        {
                            outpath = Path.Combine(outpath, finame);
                            Issuccess = DailySalesServices.StartInsert(outpath, finame);
                        }
                      

                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, filename);
                        }
                        continue;
                    }

                    DailySalesServices.StartInsert(filepath, filename);
                    continue;
                }

               
            }
            #endregion

            #region Stock

            string filestock = Path.Combine(filelocation, "Stock");
            di = new DirectoryInfo(filestock);
            foreach (var item in di.GetFiles())
            {
                //檔案名稱
                string filename = item.Name;
                //檔案路徑
                string filepath = item.FullName;
                //解壓路徑
                string outpath = "";
                string finame = GetNameCode(filename);
                if (finame.Length == 12 && GetLastCode(finame) == "A")
                {
                    if (Path.GetExtension(filename).Contains("zip"))
                    {
                        outpath = Path.Combine(filelocation, now + "A");
                        Common.CreateLocation(outpath, "");
                        //解壓縮
                        Issuccess= Common.UnZipToFile(filepath, outpath);

                        if (Issuccess)
                        {
                           outpath = Path.Combine(outpath, finame);
                          Issuccess = DailyStockServices.StartInsert(outpath, finame);
                        }
                        if (Issuccess)
                        {
                            //移至FileLocation
                            Common.WriteLog("移至FileLocation");
                            Common.WriteFile(outpath, OutToPath, filename);
                        }
                        continue;
                    }

                    DailyStockServices.StartInsert(filepath, filename);
                    continue;
                }
            }
            #endregion


        }
        //抓最後一個字
        private static string GetLastCode(string _str)
        {
            _str = _str.Substring(_str.Length - 1, 1);
            return _str;
        }


        //抓檔名
        private static string GetNameCode(string _str)
        {
            _str = _str.Substring(0, _str.Length - 4);
            return _str;
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

    }
}
