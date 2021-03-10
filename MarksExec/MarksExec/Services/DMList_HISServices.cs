using MarksExec.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using Dapper;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Data.OleDb;


namespace MarksExec.Services
{
    class DMList_HISServices
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

                if (filename.Contains("DMList"))
                {
                    Issuccess = StartInsert(filepath, filename);
                    if (Issuccess)
                    {
                        File.Move(filepath, Path.Combine(OutToPath, filename), true);
                    }

                }
            }
        }
        //開始新增
        public static bool StartInsert(string path, string finame)
        {
            bool result = true;
            string sql = "select * from DMList";
            try
            {
                Common.WriteLog("開始讀取 " + finame);
                DataTable dt = GetOleDbDataTable(path, sql);

               // var DmList = DataTableToList(dt);




                Common.WriteLog("新增DB");
                result = InsertToSql(dt);



                Common.WriteLog(finame + "讀取完成，一共 " + dt.Rows.Count + " 筆");
                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("DMList_HISServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;

            }



        }


        //連線ACCESS參考項目
        public static OleDbConnection OleDbOpenConn(string Database)
        {
            string cnstr = string.Format("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" + Database);
            OleDbConnection icn = new OleDbConnection();
            icn.ConnectionString = cnstr;
            if (icn.State == ConnectionState.Open) icn.Close();
            icn.Open();
            return icn;
        }


        //連線ACCESS，取得資料表:

        public static DataTable GetOleDbDataTable(string Database, string OleDbString)
        {
            DataTable myDataTable = new DataTable();
            OleDbConnection icn = OleDbOpenConn(Database);
            OleDbDataAdapter da = new OleDbDataAdapter(OleDbString, icn);
            DataSet ds = new DataSet();
            ds.Clear();
            da.Fill(ds);
            myDataTable = ds.Tables[0];
            if (icn.State == ConnectionState.Open) icn.Close();
            return myDataTable;
        }



        //進DB
        private static bool InsertToSql(DataTable DmList)
        {
            bool result = true;


            string connectionStrings = ConfigurationManager.ConnectionStrings["Sasc4ConnectionString"].ConnectionString;

            SqlConnection conn = new SqlConnection(connectionStrings);

            using (conn)
            {
                conn.Open();

                try
                {


                    var bulkCopy = new SqlBulkCopy(conn)
                    {
                        DestinationTableName = "[dbo].[DMList_HIS]",
                        BatchSize = 1000
                    };
                    bulkCopy.WriteToServer(DmList);
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
            string strsql = "truncate table DMList_HIS ";

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



        //public static List<DMList_HIS> DataTableToList(DataTable dt)
        //{
        //    List<DMList_HIS> DmList = new List<DMList_HIS>();
        //    int i = 0;
           
        //    foreach (DataRow dr in dt.Rows)
        //    {

               

        //        DMList_HIS doc = new DMList_HIS();

        //        doc.Year = dr.Field<string>("Year");

        //        doc.NewDMCode = dr.Field<string>("NewDMCode");

        //        doc.NewDMScheduleDMFrom = Convert.IsDBNull(dt.Rows[i]["NewDMScheduleDMFrom"]) ? "" : dr.Field<DateTime>("NewDMScheduleDMFrom").ToString("yyyyMMdd");

        //        doc.NewDMScheduleDMTo = Convert.IsDBNull(dt.Rows[i]["NewDMScheduleDMTo"]) ? "" : dr.Field<DateTime>("NewDMScheduleDMTo").ToString("yyyyMMdd");

        //        doc.Festival = dr.Field<string>("Festival");

        //        doc.DM_Code = dr.Field<string>("DM Code");

        //        doc.Divison = dr.Field<string>("Divison");

        //        doc.FullCode = dr.Field<string>("FullCode");

        //        doc.Dept = dr.Field<string>("Dept");

        //        doc.NationalTGNo = dr.Field<string>("NationalTGNo");

        //        doc.Dpno_Key = dr.Field<string>("Dpno_Key");

        //        doc.ItemCode = dr.Field<string>("ItemCode");

        //        doc.SubCode = dr.Field<string>("SubCode");

        //        doc.UnitCode = dr.Field<string>("UnitCode");

        //        doc.TGFrom = Convert.IsDBNull(dt.Rows[i]["TGFrom"]) ? "" : dr.Field<DateTime>("TGFrom").ToString("yyyyMMdd");

        //        doc.TGTo = Convert.IsDBNull(dt.Rows[i]["TGTo"]) ? "":dr.Field<DateTime>("TGTo").ToString("yyyyMMdd");
                  
        //        doc.PPP1 = Convert.IsDBNull(dt.Rows[i]["PPP1"]) ? "0" : dr.Field<double>("PPP1").ToString();

        //        doc.PPP1_from = Convert.IsDBNull(dt.Rows[i]["PPP1_from"]) ? "" : dr.Field<DateTime>("PPP1_from").ToString("yyyyMMdd");

        //        doc.PPP1_to = Convert.IsDBNull(dt.Rows[i]["PPP1_to"]) ? "" : dr.Field<DateTime>("PPP1_to").ToString("yyyyMMdd");

        //        doc.PPP2 = Convert.IsDBNull(dt.Rows[i]["PPP2"]) ? "0" : dr.Field<double>("PPP2").ToString();

        //        doc.PurchasePeriodFrom = Convert.IsDBNull(dt.Rows[i]["PurchasePeriodFrom"]) ? "" : dr.Field<DateTime>("PurchasePeriodFrom").ToString("yyyyMMdd");

        //        doc.PurchasePeriodTo = Convert.IsDBNull(dt.Rows[i]["PurchasePeriodTo"]) ? "" : dr.Field<DateTime>("PurchasePeriodTo").ToString("yyyyMMdd");

        //        doc.NormalSellingPrice = Convert.IsDBNull(dt.Rows[i]["NormalSellingPrice"]) ? "0" :dr.Field<double>("NormalSellingPrice").ToString();

        //        doc.PSP_from = Convert.IsDBNull(dt.Rows[i]["PSP_from"]) ? "" : dr.Field<DateTime>("PSP_from").ToString("yyyyMMdd");

        //        doc.PSP_to = Convert.IsDBNull(dt.Rows[i]["PSP_to"]) ? "": dr.Field<DateTime>("PSP_to").ToString("yyyyMMdd");

        //        doc.DC = dr.Field<string>("DC");

        //        doc.Remark = dr.Field<string>("Remark");

        //        doc.CDes = dr.Field<string>("CDes");

        //        doc.EDes = dr.Field<string>("EDes");

        //        doc.QtyBox = dr.Field<string>("QtyBox");

        //        doc.CName = dr.Field<string>("CName");

        //        doc.EName = dr.Field<string>("EName");

        //        doc.Capacity = dr.Field<string>("Capacity");

        //        doc.MainSupplier = dr.Field<string>("MainSupplier");

        //        doc.RealSupplier = dr.Field<string>("RealSupplier");

        //        doc.SupCName = dr.Field<string>("SupCName");

        //        doc.SupEName = dr.Field<string>("SupEName");

        //        doc.Media = dr.Field<string>("Media");

        //        doc.TGEName = dr.Field<string>("TGEName");

        //        doc.ForecastQty = Convert.IsDBNull(dt.Rows[i]["ForecastQty"]) ? "0" : dr.Field<double>("ForecastQty").ToString();

        //        doc.Big = dr.Field<string>("Big");

        //        doc.DM_Page = Convert.IsDBNull(dt.Rows[i]["DM_Page"]) ? "0" : dr.Field<Int16>("DM_Page").ToString();

        //        doc.sf_ranking = dr.Field<string>("sf_ranking");

        //        doc.Hyper_store_G1 = dr.Field<string>("Hyper_store_G1");

        //        doc.Super_store_G1 = dr.Field<string>("Super_store_G1");

        //        doc.Store_G1_mapping = dr.Field<string>("Store_G1_mapping");

        //        doc.Hyper_store_G2 = dr.Field<string>("Hyper_store_G2");

        //        doc.Super_store_G2 = dr.Field<string>("Super_store_G2");

        //        doc.Store_G2_mapping = dr.Field<string>("Store_G2_mapping");

        //        doc.Hyper_store_G3 = dr.Field<string>("Hyper_store_G3");

        //        doc.Super_store_G3 = dr.Field<string>("Super_store_G3");

        //        doc.Store_G3_mapping = dr.Field<string>("Store_G3_mapping");

        //        doc.Limitation = dr.Field<string>("Limitation");

        //        doc.LimitationQty = Convert.IsDBNull(dt.Rows[i]["LimitationQty"]) ? "0" : dr.Field<double>("LimitationQty").ToString();

        //        doc.MainUnitCode = dr.Field<string>("MainUnitCode");

        //        doc.Loyalty = dr.Field<string>("Loyalty");

        //        doc.LoyaltyPrice = Convert.IsDBNull(dt.Rows[i]["LoyaltyPrice"]) ? "0" : dr.Field<int>("LoyaltyPrice").ToString();

        //        doc.ExtraPoints = Convert.IsDBNull(dt.Rows[i]["ExtraPoints"]) ? "0" : dr.Field<double>("ExtraPoints").ToString();

        //        doc.Priority = Convert.IsDBNull(dt.Rows[i]["Priority"]) ? "0" : dr.Field<double>("Priority").ToString();

        //        doc.ForecastSales = Convert.IsDBNull(dt.Rows[i]["ForecastSales"]) ? "0" : dr.Field<double>("ForecastSales").ToString();

        //        doc.CategoryManager = dr.Field<string>("CategoryManager");

        //        doc.MD_Push = dr.Field<string>("MD_Push");
        //        doc.NOcode = dr.Field<string>("NOcode");

        //        doc.PCD = dr.Field<string>("PCD");

        //        doc.MajorBarcode = dr.Field<string>("MajorBarcode");

        //        doc.normal_selling_price = dr.Field<string>("normal selling price");

        //        doc.promo_selling_price = dr.Field<string>("promo selling price");

        //        doc.DM_Activities_remark = dr.Field<string>("DM Activities remark");

        //        doc.Promotion_Effect_Type = dr.Field<string>("Promotion Effect Type");

        //        DmList.Add(doc);
        //        i++;
        //    }
        //    return DmList;
        //}


    }
}
