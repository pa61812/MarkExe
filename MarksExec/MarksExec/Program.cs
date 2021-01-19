using MarksExec.Services;
using System;
using System.Collections.Generic;
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

            //時間測試
            //DateTime start = DateTime.Now;



            //吃檔目錄
            string filelocation = ConfigurationManager.AppSettings["FileInPut"];
            string filemaster = Path.Combine(filelocation, "Master");
            string filestock = Path.Combine(filelocation, "Stock");
            string filesales = Path.Combine(filelocation, "Sales");          

            //AppConfig的KEY
            var  appset = GetAppSetting();

            //若參數不存於KEY，結束程式
            string input = args[0];
            string appvalue = "";
            if (appset.Contains(input.ToUpper()))
            {
                appvalue = ConfigurationManager.AppSettings[input.ToUpper()];                
            }
            else
            {
                Console.WriteLine("參數不正確，請重新輸入");
                return;
            }

            //Out路徑
            string OutToPath = Common.OutLocation();

            switch (appvalue)
            {
                case "BARCODE":
                    BARServices.GetFile(filemaster, OutToPath);
                break;
                case "BIGSALES":
                    BigSalesServices.GetFile(filemaster, OutToPath);
                    break;
                case "DAILYSALES":
                    DailySalesServices.GetFile(filesales, OutToPath);
                    break;
                case "DAILYSTOCK":
                    DailyStockServices.GetFile(filestock, OutToPath);
                    break;
                case "ITMGEL":
                    ITMGELServices.GetFile(filemaster, OutToPath);
                    break;
                case "ITMSUB":
                    ITMSUBServices.GetFile(filemaster, OutToPath);
                    break;
                case "ITMSUPGEXCRTNC":
                    ITMSUPGEXCRTNCServices.GetFile(filemaster, OutToPath);
                    break;
                case "STM":
                    STMTMPServices.GetFile(filemaster, OutToPath);
                    break;
                case "SUPATT":
                    SUPATTServices.GetFile(filemaster, OutToPath);
                    break;
                case "DSL":
                    DSLServices.GetFile(@"D:\BankPro\Carrefour\20210112\DSL_ACC_20201110", OutToPath);
                    break;
                case "EMP_DATA":
                    EmpDataServices.GetFile(filemaster, OutToPath);
                    break;

            }


            //時間測試
            //DateTime end = DateTime.Now;
            //Console.WriteLine(appvalue);
            //Console.WriteLine(start.ToString("hh:mm:ss"));
            //Console.WriteLine(end.ToString("hh:mm:ss"));


        }

        //AppConfig所有的KEY
        public static List<string> GetAppSetting()
        {
            List<string> result = new List<string>();

            foreach (string s in ConfigurationManager.AppSettings)
            {
                if (s== "FileInPut" || s== "FileLocation" || s== "LogLocation")
                {
                    continue;
                }


                result.Add(s.ToUpper());
            }


            return result;

        }


       
     

    }
}
