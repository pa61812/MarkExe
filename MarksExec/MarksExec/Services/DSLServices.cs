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
    class DSLServices
    {
        public static string[] itemcode = { "943074","020053","310062","050002","200006"
                                ,"533232","102054","020906","502924","312039"};
        public static string[] dept = { "20","14","15","10","12"
                             ,"31","10","12","66","76"};
        public static List<SearchStoreCode> searchstore= SearchStore();

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

          

            foreach (var item in di.GetFiles())
            {
                i++;
                //檔案名稱
                string filename = item.Name;
                //檔案路徑
                string filepath = item.FullName;
                //解壓路徑
                string outpath = "";
                string finame = Common.GetNameCode(filename);

                if (filename.Contains("DSL_ACC"))
                {
                    Issuccess = StartInsert(filepath, filename);
                    if (true)
                    {
                        Issuccess = Common.MoveFile(filepath, OutToPath, filename);
                    }        
                    continue;
                }
            }
        }
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            DSLDailySales daily = new DSLDailySales();
            List<DSLDailySales> sales = new List<DSLDailySales>();

            string line;
            try
            {
                //使用FileStream讀取檔案
                FileStream fileStream = File.OpenRead(path);
                StreamReader reader = new StreamReader(fileStream, Encoding.UTF8);
                string filename = Path.GetFileName(finame).Substring(0, 3);
                int i = 0;
                line = reader.ReadLine();
                Common.WriteLog("切割字串");

              
                //Console.WriteLine("切割字串");
                //start = DateTime.Now;

                while (line != null)
                {
                    if (i == 0)
                    {
                        line = reader.ReadLine();
                    }
                    line = new Regex("[\\s]+").Replace(line, " ");
                    if (line == " " || line == "")
                    {
                        line = reader.ReadLine();
                        continue;
                    }
                    daily = ChangString(line);
                    if (daily.DepartmentCode != null)
                    {
                        sales.Add(daily);
                        if (i % 100000 == 0)
                        {
                            result = InsertToSql(sales);
                            sales = new List<DSLDailySales>();
                            if (!result)
                            {
                                return false;
                            }
                        }
                    }
                   
                    //Read the next line
                    line = reader.ReadLine();
                    i++;
                }

                //end = DateTime.Now;
                //Console.WriteLine(start.ToString("hh:mm:ss"));
                //Console.WriteLine(end.ToString("hh:mm:ss"));


                reader.Close();


                Common.WriteLog("新增DB");
                result = InsertToSql(sales);




                Common.WriteLog("新增 " + finame + " 成功，一共 " + i + " 筆");
                Console.WriteLine("新增 " + finame + " 成功，一共 " + i + " 筆");

                return result;
            }
            catch (Exception ex)
            {
                result = false;
                Common.WriteLog(ex.ToString());
                return result;
            }
        }



        private static DSLDailySales ChangString( string line)
        {
            DSLDailySales dailSales = new DSLDailySales();
           // var searchstore = SearchStore();
            string storecode = "";
            try
            {
                //Common.WriteLog("切割字串 " + line);
                //去除多空白
                line = new Regex("[\\s]+").Replace(line, " ");
                //空白切割
                string[] words = line.Split('|'); 
                //string[] words = Split2(line + "|", "|").ToArray();

                storecode =searchstore.Where(x => x.Store == words[0]).Select(x=>x.StoreCode).FirstOrDefault();
                //if (Array.IndexOf(itemcode, words[2]) < 0 || Array.IndexOf(dept, words[1]) < 0)
                //{
                //    return dailSales;
                //}

                dailSales = new DSLDailySales
                {
                    StoreCode = storecode,
                    DepartmentCode = words[1],
                    ItemCode = words[2],
                    SubCode = words[3],
                    SalesDate = words[4],
                    UnitCode = words[5],
                    StatusPromotion = words[6],
                    SPType = words[7],
                    VatRate = words[8],
                    SalesQty = words[9],
                    SalesAmount = words[10],
                    SalesPrice = words[11],
                    PurchasePrice = words[12],
                    Rebate = words[13],
                    VatAmount = words[14],
                    Input= words[15],
                    OriginalSellingPrice = "0",
                    CSTAmount = "0"                    
                };
                return dailSales;
            }
            catch (Exception ex)
            {

                Common.WriteLog("切割字串 " + line + " 失敗");

                Common.WriteLog(ex.ToString());
                return dailSales;
            }



        }

        public static IEnumerable<string> Split2(IEnumerable<char> source, string splitStr)
        {
            var enumerator = source.GetEnumerator();
            string line = string.Empty;
            while (enumerator.MoveNext())
            {
                line += (char)enumerator.Current;
                if (line.EndsWith(splitStr))
                {
                    yield return line.Substring(0, line.Length - splitStr.Length);
                    //Console.WriteLine(line);
                    line = string.Empty;
                }
            }
        }

        private static bool InsertToSql(List<DSLDailySales> sasc4)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);

            using (conn)
            {
                conn.Open();

                try
                {
                    var dtdailySales = Common.ConvertToDataTable<DSLDailySales>(sasc4);


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[DailySales_TMP]",
                        BatchSize = 1000
                    };
                    bulkCopy.WriteToServer(dtdailySales);
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


        private static List<SearchStoreCode> SearchStore()
        {
            List<SearchStoreCode> result = new List<SearchStoreCode>();

            string strMsgSelect = "select store, storecode  from StoreTmp ";

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);

            using (conn)
            {            
                result = conn.Query<SearchStoreCode>(strMsgSelect).ToList();            
            }


            return result;
        }



    }
}
