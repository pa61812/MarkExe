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
    class ECSalesSupplyServices
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
            bool Issuccess = true;
            int i = 0;
            string now = DateTime.Now.ToString("yyyyMMdd");
            string newpath = "";
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


                if (filename.Contains("ECOrderDetail"))
                {
                    if (Issuccess)
                    {
                        //Issuccess = StartFind(filepath, filename);
                        Issuccess = StartInsert(filepath, filename);
                    }
                    if (Issuccess)
                    {
                        //移至FileLocation
                        Common.WriteLog("移至FileLocation");
                        Issuccess = Common.MoveFile(filepath, OutToPath, filename);
                    }
                    continue;
                }
            }
        }

        //path 完整路徑  finame檔名
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<ECSalesSupply> ecSales = new List<ECSalesSupply>();
            List<ECSalesSupply> ErroecSales = new List<ECSalesSupply>();
            ECSalesSupply ecSale = new ECSalesSupply();
            try
            {
                Common.WriteLog("開始切割 " + finame);
                FileStream fileStream = File.OpenRead(path);

                //using (StreamReader reader = new StreamReader(fileStream, Encoding.Default))
                using (StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(0)))
                {
                    line = reader.ReadLine();
                    while (line != null)
                    {
                        i++;
                        if (i == 1)
                        {
                            line = reader.ReadLine();
                        }
                        line = new Regex("[\\s]+").Replace(line, " ");
                        if (line == " " || line == "")
                        {
                            line = reader.ReadLine();
                            continue;
                        }
                        ecSale = splitstrng(line, i);
                        ecSales.Add(ecSale);

                        if (i % 100000 == 0)
                        {
                            result = InsertToSql(ecSales);
                            ecSales = new List<ECSalesSupply>();
                            if (!result)
                            {
                                return false;
                            }
                        }
                      

                        line = reader.ReadLine();


                    }

                    reader.Close();
                }




                Common.WriteLog("新增DB");
                result = InsertToSql(ecSales);

              



                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");

                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("EC_DetailServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;
            }


        }


        //切割字串
        public static ECSalesSupply splitstrng(string input, int i)
        {
            ECSalesSupply ecSales = new ECSalesSupply();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");
                string[] sArray = null;
                sArray = input.Split('|');

                ecSales = new ECSalesSupply
                {
                    OrderId=sArray[0],
                    OrderNo = sArray[1],
                    DeliveryDate = sArray[2],
                    DeliveryTime = sArray[3],
                    StoreCode = sArray[4],
                    OutStore = sArray[5],
                    InputStore = sArray[6],
                    OrderType = sArray[7],
                    Media = sArray[8],
                    DeliveryBy = sArray[9],
                    OrderStatus = sArray[10],
                    LogisticStatus = sArray[11],
                    MemberEmail = sArray[12],
                    Consignee = sArray[13],
                    ConsigneeTel = sArray[14],
                    ConsigneePhone = sArray[15],
                    ZipCode = sArray[16],
                    OrderCounty = sArray[17],
                    OrderArea = sArray[18],
                    DeliveryAddress = sArray[19],
                    OrderRemarks = sArray[20],
                    OrderDate = sArray[21],
                    PaymentBy = sArray[22],
                    PaymentStatus = sArray[23],
                    OrderAmt = sArray[24],
                    ItemFullCode = sArray[25],
                    ItemMainCode = sArray[26],
                    ItemCName = sArray[27],
                    SalesQty = sArray[28],
                    SalesPrice = sArray[29],
                    ISGiveAway = sArray[30],
                    AddOn = sArray[31],
                    LangCulture = sArray[32],
                    OrderAttributes = sArray[33],
                    StoreOrdertype = sArray[34],
                    OrderPrice = sArray[35],
                    Temperature = sArray[36],
                    IsMember = sArray[37],
                    MemberAccount = sArray[38],
                    CardNo = sArray[39],
                    DiscountCode = sArray[40],
                    DiscountType = sArray[41],
                    InvoiceDate = sArray[42],
                    PromotionActivityID = sArray[43],
                    ActivityID = sArray[44],
                    CompnayOrderNo = sArray[45]
                };
                   




                return ecSales;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return ecSales;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<ECSalesSupply> ecSales)
        {
            bool result = true;

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);


            using (conn)
            {
                conn.Open();
                try
                {

                    var ec = Common.ConvertToDataTable<ECSalesSupply>(ecSales);


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[ECSalesSupply]",
                        BatchSize = 10000
                    };
                    bulkCopy.WriteToServer(ec);
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
            string strsql = "truncate table  ECSalesSupply ";

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
