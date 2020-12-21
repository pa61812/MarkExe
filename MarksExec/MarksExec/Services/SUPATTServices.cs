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
                //加上BeginTrans
                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string strsql = "Insert into  SUPATT_TMP " +
                               "values(@StoreCode,@AttributeClassCode,@AttributeClassDescription"
                                    + ",@AttributeCode,@AttributeCodeDescription,@AlphanumericValue"
                                    + ",@NumberValue,@Date,@Time,@StartDate,@EndDate)";




                        conn.Execute(strsql, Sup, transaction);


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
