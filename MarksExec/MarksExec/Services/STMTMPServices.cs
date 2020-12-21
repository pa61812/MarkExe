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
    class STMTMPServices
    {
        //path 路徑  finame檔名

        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<STM_TMP> stms = new List<STM_TMP>();
            STM_TMP stm = new STM_TMP();

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

                        stm = splitstrng(line);
                        stms.Add(stm);
                        line = reader.ReadLine();

                        i++;

                    }

                    reader.Close();
                }

                Common.WriteLog("新增DB");
                result = InsertToSql(stms);

                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");

                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("STMTMPServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }

          
        }


        //切割字串
        public static STM_TMP splitstrng(string input)
        {
            STM_TMP stm = new STM_TMP();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split('|');

                stm = new STM_TMP
                {
                    Dept=sArray[0],
                    StructureCode = sArray[1],
                    Cname = sArray[2],
                    Ename = sArray[3],
                    StructureQty = sArray[4],
                    UsedQty = sArray[5],
                    LogisticBoxRequire = sArray[6],
                    MultiTaxRate = sArray[7],
                    LicenseYN = sArray[8],
                    LicenseType = sArray[9],
                    FoodOrNonFood = sArray[10]

                };

                return stm;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return stm;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<STM_TMP> stm)
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
                        string strsql = "Insert into  STM_TMP " +
                               "values(@Dept,@StructureCode,@Cname,@Ename,@StructureQty,@UsedQty"
                                   + ",@LogisticBoxRequire,@MultiTaxRate,@LicenseYN,@LicenseType,@FoodOrNonFood)";



                        conn.Execute(strsql, stm, transaction);


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
