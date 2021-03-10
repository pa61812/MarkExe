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
    class EC_DetailServices
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
            bool Issuccess=true;
            int i = 0;
            string now = DateTime.Now.ToString("yyyyMMdd");
            string newpath = "";
            //刪除DB，成功在進行Insert
             Common.WriteLog("刪除DB");
            //Issuccess = DeleteTable();

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

             
                if (filename.Contains("ECDetail"))
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

        public static bool StartFind(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<ECSales> ecSales = new List<ECSales>();
            List<ECSales> ErroecSales = new List<ECSales>();
            ECSales ecSale = new ECSales();
            try
            {
                Common.WriteLog("開始切割 " + finame);
                FileStream fileStream = File.OpenRead(path);
                using (StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(0)))
                //using (StreamReader reader = new StreamReader(fileStream, Encoding.Default))
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
                         splitstrng(line, i);


                        line = reader.ReadLine();


                    }

                    reader.Close();
                }

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


        //path 完整路徑  finame檔名
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            int i = 0;
            string line;
            List<ECSales> ecSales = new List<ECSales>();
            List<ECSales> ErroecSales = new List<ECSales>();
            ECSales ecSale = new ECSales();
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
                        if (i==1)
                        {
                            line = reader.ReadLine();
                        }
                        line = new Regex("[\\s]+").Replace(line, " ");
                        if (line == " " || line == "")
                        {
                            line = reader.ReadLine();
                            continue;
                        }
                        ecSale = splitstrng(line,i);

                        //如果資料沒問題
                        if (ecSale.Address!=null && ecSale.Address !="" && ecSale.SalesQty!="ERRO" &&ecSale.PaymentBy!= "ERRO")
                        {
                            ecSales.Add(ecSale);
                            if (i%100000 ==0)
                            {
                                result = InsertToSql(ecSales);
                                ecSales= new List<ECSales>();
                                if (!result)
                                {
                                    return false;
                                }
                            }
                        }

                        //如果資料有問題
                        if (ecSale.Address != null && ecSale.Address != "" && (ecSale.SalesQty == "ERRO" || ecSale.PaymentBy == "ERRO"))
                        {
                            ErroecSales.Add(ecSale);
                            if (i % 100000 == 0)
                            {
                                result = InsertToErroSql(ErroecSales);
                                ErroecSales = new List<ECSales>();
                                if (!result)
                                {
                                    return false;
                                }
                            }
                        }

                        line = reader.ReadLine();
                       

                    }

                    reader.Close();
                }




                Common.WriteLog("新增DB");
                result = InsertToSql(ecSales);

                //如果有資料有問題再INSERT
                if (ErroecSales.Count >= 1)
                {
                    result = InsertToErroSql(ErroecSales);
                }
               


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
        public static ECSales splitstrng(string input,int i)
        {
            ECSales ecSales = new ECSales();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");
                string[] sArray = null;
                sArray = input.Split('|');


                string SalesQty = "";
                string SalesPrice = "";
                string SalesQtyUnit = "";
                string SalesAmt = "";
                string PaymentBy = "";
                if (sArray.Length < 26)
                {
                    //Console.WriteLine(input);
                    //Console.WriteLine("第 " + i + " 行");
                    Common.WriteLog("切割失敗 " + input + " ");
                    Common.WriteLog("第 " + i + " 行");

                    return ecSales;
                }
                if (sArray.Length<27 && int.TryParse(sArray[22], out i)==false)
                {
                    PaymentBy = sArray[16];
                    //SalesQty = sArray[22].Substring(sArray[22].Length-1,1);
                    SalesQty = "ERRO";
                    SalesPrice = sArray[23];
                    SalesQtyUnit = sArray[24];
                    SalesAmt = sArray[25];

                    ecSales = new ECSales
                    {
                        Store = sArray[0],
                        StoreCName = sArray[1],
                        StoreZIP = sArray[2],
                        OrderDate = sArray[3],
                        OrderNo = sArray[4],
                        Media = sArray[5],
                        DeliveryDate = sArray[6],
                        InvoiceDate = sArray[7],
                        InvoiceDate_Actual = sArray[8],
                        DeliveryBy = sArray[9],
                        OrderStatus = sArray[10],
                        LogisticStatus = sArray[11],
                        ZIPCode1 = sArray[12],
                        County = sArray[13],
                        District = sArray[14],
                        Address = sArray[15],
                        PaymentBy = PaymentBy,
                        PaymentStatus = sArray[17],
                        OrderAmt = sArray[18],
                        FullCode = sArray[19],
                        Unit = sArray[20],
                        Input = sArray[21],
                        ItemCName = sArray[22],
                        SalesQty = SalesQty,
                        SalesPrice = SalesPrice,
                        SalesQtyUnit = SalesQtyUnit,
                        SalesAmt = SalesAmt
                    };

                }
                else if (int.TryParse(sArray[22], out i))
                {
                    PaymentBy = "ERRO";
                    SalesQty = sArray[22] ;
                    SalesPrice = sArray[23];
                    SalesQtyUnit = sArray[24];
                    SalesAmt = sArray[25];

                    ecSales = new ECSales
                    {
                        Store = sArray[0],
                        StoreCName = sArray[1],
                        StoreZIP = sArray[2],
                        OrderDate = sArray[3],
                        OrderNo = sArray[4],
                        Media = sArray[5],
                        DeliveryDate = sArray[6],
                        InvoiceDate = sArray[7],
                        InvoiceDate_Actual = sArray[8],
                        DeliveryBy = sArray[9],
                        OrderStatus = sArray[10],
                        LogisticStatus = sArray[11],
                        ZIPCode1 = sArray[12],
                        County = sArray[13],
                        District = sArray[14],
                        Address = sArray[15],
                        PaymentBy = PaymentBy,
                        PaymentStatus = sArray[16],
                        OrderAmt = sArray[17],
                        FullCode = sArray[18],
                        Unit = sArray[19],
                        Input = sArray[20],
                        ItemCName = sArray[21],
                        SalesQty = SalesQty,
                        SalesPrice = SalesPrice,
                        SalesQtyUnit = SalesQtyUnit,
                        SalesAmt = SalesAmt
                    };
                }
                else
                {
                    PaymentBy = sArray[16];
                    SalesQty = sArray[23];
                    SalesPrice = sArray[24];
                    SalesQtyUnit = sArray[25];
                    SalesAmt = sArray[26];

                    ecSales = new ECSales
                    {
                        Store = sArray[0],
                        StoreCName = sArray[1],
                        StoreZIP = sArray[2],
                        OrderDate = sArray[3],
                        OrderNo = sArray[4],
                        Media = sArray[5],
                        DeliveryDate = sArray[6],
                        InvoiceDate = sArray[7],
                        InvoiceDate_Actual = sArray[8],
                        DeliveryBy = sArray[9],
                        OrderStatus = sArray[10],
                        LogisticStatus = sArray[11],
                        ZIPCode1 = sArray[12],
                        County = sArray[13],
                        District = sArray[14],
                        Address = sArray[15],
                        PaymentBy = PaymentBy,
                        PaymentStatus = sArray[17],
                        OrderAmt = sArray[18],
                        FullCode = sArray[19],
                        Unit = sArray[20],
                        Input = sArray[21],
                        ItemCName = sArray[22],
                        SalesQty = SalesQty,
                        SalesPrice = SalesPrice,
                        SalesQtyUnit = SalesQtyUnit,
                        SalesAmt = SalesAmt
                    };
                }



                

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
        private static bool InsertToSql(List<ECSales> ecSales)
        {
            bool result = true;

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);


            using (conn)
            {
                conn.Open();
                try
                {
            
                    var ec = Common.ConvertToDataTable<ECSales>(ecSales);

                   
                        var bulkCopy = new SqlBulkCopy(conn)
                        {
                            DestinationTableName = "[dbo].[ECSales_TMP]",
                            BatchSize = 1000
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

        ////進DB
        //private static bool InsertToSql(List<ECSales> ecSales)
        //{
        //    bool result = true;

        //    string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

        //    SqlConnection conn = new SqlConnection(connectionStrings);


        //    using (conn)
        //    {
        //        conn.Open();
        //        try
        //        {
        //            DapperPlusManager.Entity<ECSales>().Table("ECSales");
        //            conn.BulkInsert(ecSales);
        //            conn.Close();
        //            Common.WriteLog("新增成功");
        //            return result;
        //        }
        //        catch (Exception e)
        //        {

        //            Common.WriteLog("新增失敗");
        //            Common.WriteLog(e.ToString());
        //            conn.Close();
        //            result = false;
        //            return result;
        //        }

        //    }
        //}




        //進DB
        private static bool InsertToErroSql(List<ECSales> ecSales)
        {
            bool result = true;

            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);


            using (conn)
            {
                conn.Open();
                try
                {

                    var ec = Common.ConvertToDataTable<ECSales>(ecSales);


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[ECSales_ERRO]",
                        BatchSize = 1000
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
            string strsql = "truncate table  ECSales ";

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
