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
    class ITMSUBServices
    {
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
                //加上BeginTrans
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string strsql = "Insert into  ITMSUB_TMP " +
                              "values(@Dept,@ItemCode,@SubID,@SubCode"
                             +",@ItemID,@SubCodeEDes,@SubCodeCDes,@StartStopDate"
                             +",@StopItemReasonCode,@ProductCharacteristic,@EndStopDate"
                             +",@ReturnableforNonStopItem,@ReturnableforTemporaryStop,@ReturnableforPermanentStop"
                             +",@WHCheckingMark,@ItemLevel,@UsebyDate,@CNCode)";




                        conn.Execute(strsql, iTMSUBs, transaction);


                        //正確就Commit
                        transaction.Commit();
                        conn.Close();

                        Common.WriteLog("新增成功");
                        return result;
                    }
                    catch (Exception e)
                    {
                        transaction.Rollback();
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
}
