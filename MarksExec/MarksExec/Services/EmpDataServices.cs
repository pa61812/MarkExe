using Dapper;
using MarksExec.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;


namespace MarksExec.Services
{
    class EmpDataServices
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
                bool checkfilename=false;

                if (filename.Contains("EMP_DATA"))
                {
                    checkfilename = CheckFilename(Path.GetFileNameWithoutExtension(filepath));
                }
                if (checkfilename)
                {
                    Issuccess = StartInsert(filepath, filename);
                    if (Issuccess)
                    {
                        //移至FileLocation
                        Common.WriteLog("移至FileLocation");
                        Issuccess = Common.MoveFile(filepath, OutToPath, filename);
                    }
                    //刪除檔案
                    //if (Issuccess)
                    //{
                    //    Common.DeleteFolder(outpath);
                    //}
                    continue;
                }
            }
        }

        //檢查日期是否為今天
        //filename檔名
        public static bool CheckFilename(string filename)
        {
            bool result = true;

            string datenow = DateTime.Now.ToString("yyyyMMdd");


            if (datenow== filename.Substring(filename.Length - 8,8 ))
            {
                result = true;
            }
            else
            {
                result = false;
            }


            return result;
        }


        //path 完整路徑  finame檔名
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<EMP_DATA> empdatas = new List<EMP_DATA>();
            EMP_DATA empdata = new EMP_DATA();
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
                        empdata = splitstrng(line);
                        empdatas.Add(empdata);
                        line = reader.ReadLine();
                        i++;

                    }

                    reader.Close();
                }




                Common.WriteLog("新增DB");
                result = InsertToSql(empdatas);



                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");

                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("EmpDataServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }


        }


        //切割字串
        public static EMP_DATA splitstrng(string input)
        {
            EMP_DATA empdata = new EMP_DATA();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split(',');

                empdata = new EMP_DATA
                {
                    Emp_No= sArray[0],
                    Login_Name = sArray[1],
                    Emp_Name = sArray[2],
                    Mail_External = sArray[3],
                    Mail_Internal = sArray[4],
                    Entry_Date = sArray[5],
                    Dept_Code = sArray[6],
                    Org_code = sArray[7],
                    Org_NAME = sArray[8],
                    Job_Desc_Eng = sArray[9],
                    Job_Desc_Cht = sArray[10],
                    Direct_Manager = sArray[11],
                    Region = sArray[12],
                    REGION_NAME = sArray[13],
                    TW_sTORE_ID = sArray[14],
                    TW_sTORE_Name = sArray[15],
                    bu = sArray[16],
                    bu_NAME = sArray[17],
                    Job_code = sArray[18],
                    bank_no = sArray[19],
                    on_job = sArray[20],
                    NIGHT_SHIFT = sArray[21],
                    SPEC_TYPE = sArray[22],
                    JOB_TITLE = sArray[23],
                    FT_PT = sArray[24],
                    JOB_LEVEL = sArray[25],
                    STORE_ID = sArray[26],
                    DEPT_ID = sArray[27],
                    MOD_TIME = sArray[28],
                    MOD_USER = sArray[29],
                    MOD_PGM = sArray[30],
                    MOD_WS = sArray[31],
                    STORE_NAME_E = sArray[32],
                    STORE_NAME_C = sArray[33],
                    ID_NO = sArray[34],
                    OUT_DATE = sArray[35],
                    SEX = sArray[36],
                    Birthday = sArray[37]
                };

                return empdata;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return empdata;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<EMP_DATA> empdata)
        {
            bool result = true;

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);


            using (conn)
            {
                conn.Open();
                try
                {
                    var dtempdata = Common.ConvertToDataTable<EMP_DATA>(empdata);


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[EMP_DATA_TMP]",
                        BatchSize = 1000
                    };
                    bulkCopy.WriteToServer(dtempdata);
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
            string strsql = "truncate table EMP_DATA_TMP ";

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
