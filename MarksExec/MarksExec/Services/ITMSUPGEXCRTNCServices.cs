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
    class ITMSUPGEXCRTNCServices
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
                if (filename.Contains("ITMSUPGEXCRTNC"))
                {
                    if (Path.GetExtension(filename).Contains("gz"))
                    {
                        outpath = Path.Combine(path, now + "ITMSUPGEXCRTNC");
                        outfile = string.Format("{0}{1}{2}", "ITMSUPGEXCRTNC", i, ".txt");
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
            List<ITMSUPGEXCRTNC> iTMSUPs = new List<ITMSUPGEXCRTNC>();
            ITMSUPGEXCRTNC iTMSUP = new ITMSUPGEXCRTNC();

            try
            {
                Common.WriteLog("開始切割 " + finame);
                FileStream fileStream = File.OpenRead(path);
                using (StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(0)))
                {
                    line = reader.ReadLine();
                    //第一行為標頭，從第二行開始
                    if (i == 0)
                    {
                        line = reader.ReadLine();
                    }
                    while (line != null)
                    {
                        line = new Regex("[\\s]+").Replace(line, " ");
                        if (line == " " || line == "")
                        {
                            line = reader.ReadLine();
                            continue;
                        }

                        iTMSUP = splitstrng(line);
                        iTMSUPs.Add(iTMSUP);
                        line = reader.ReadLine();

                        i++;

                    }

                    reader.Close();
                }
               
                 Common.WriteLog("新增DB");
                 result = InsertToSql(iTMSUPs);
                

               

                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");
                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("ITMSUPGEXCRTNCServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }

           
        }


        //切割字串
        public static ITMSUPGEXCRTNC splitstrng(string input)
        {
            ITMSUPGEXCRTNC iTMSUP = new ITMSUPGEXCRTNC();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split('|');

                iTMSUP = new ITMSUPGEXCRTNC
                {
                    DC=sArray[0],
                    dpno=sArray[1],
                    fullcode=sArray[2],
                    SupplierCode=sArray[3],
                    stopreasoncode=sArray[4],
                    Cname = sArray[5],
                    mainsupp = sArray[6],
                    realsupp = sArray[7],
                    suppCname = sArray[8],
                    typeofsales = sArray[9],
                    Grade = sArray[10],
                    spec = sArray[11],
                    seasoncode = sArray[12],
                    Permanentstop = sArray[13],
                    Supstop = sArray[14],
                    Supstopdate = sArray[15],
                    Goodexchangeflag = sArray[16],
                    OnlyBreakage = sArray[17],
                    Venderneedcheckproduct = sArray[18],
                    Goodexchangeflagforstopitem = sArray[19],
                    OnlyBreakage1 = sArray[20],
                    Venderneedcheckproduct1 = sArray[21],
                    ReturnFlagNonStopItem = sArray[22],
                    NoReturn = sArray[23],
                    AcceptBreakage = sArray[24],
                    OnlyGoodProduct = sArray[25],
                    OnlyBreakage_incl = sArray[26],
                    OnlyBreakage_ProductDefective = sArray[27],
                    OnlyIntactStockProduct = sArray[28],
                    MDSpecialReutrn = sArray[29],
                    SupplierSpecialDemands = sArray[30],
                    VenderNeedCheckProduct2 = sArray[31],
                    CompleteOriginalAccessaryDocumentA = sArray[32],
                    OriginalCoverBox = sArray[33],
                    OriginalProductPackage = sArray[34],
                    Return_Flag_TemporaryStopItem = sArray[35],
                    NoReturn1 = sArray[36],
                    Accept_BreakageandGood_Product1 = sArray[37],
                    OnlyGoodProduct1 = sArray[38],
                    OnlyBreakage_Near_Over_ExpireDat1 = sArray[39],
                    OnlyBreakage_ProductDefective1 = sArray[40],
                    OnlyIntactStockProduct1 = sArray[41],
                    MDSpecialReutrn1 = sArray[42],
                    SupplierSpecialDemands_RefertoLogis1 = sArray[43],
                    VenderNeedCheckProduct3 = sArray[44],
                    CompleteOriginalAccessaryDocument = sArray[45],
                    OriginalCoverBox1 = sArray[46],
                    OriginalProductPackage1 = sArray[47],
                    ReturnFlag_PermanentStopItem = sArray[48],
                    NoReturn2 = sArray[49],
                    AcceptBreakageandGoodProduct2 = sArray[50],
                    OnlyGoodProduct2 = sArray[51],
                    OnlyBreakage_Near_Over_ExpireDat2 = sArray[52],
                    OnlyBreakage_ProductDefective2 = sArray[53],
                    OnlyIntactStockProduct2 = sArray[54],
                    MDSpecialReutrn2 = sArray[55],
                    SupplierSpecialDemandsRefertoLogis2 = sArray[56],
                    VenderNeedCheckProduct4 = sArray[57],
                    CompleteOriginalAccessaryDocumentA2 = sArray[58],
                    OriginalCoverBox2 = sArray[59],
                    OriginalProductPackage2 = sArray[60]


                };

                return iTMSUP;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return iTMSUP;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<ITMSUPGEXCRTNC> iTMSUPs)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);
         
            using (conn)
            {
                conn.Open();

                try
                {
                    DapperPlusManager.Entity<ITMSUPGEXCRTNC>().Table("ITMSUPGEXCRTNC_TMP");
                    conn.BulkInsert(iTMSUPs);
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
            string strsql = "delete from ITMSUPGEXCRTNC_TMP ";

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
