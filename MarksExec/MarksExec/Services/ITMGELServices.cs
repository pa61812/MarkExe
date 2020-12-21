using Dapper;
using MarksExec.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;


namespace MarksExec.Services
{
    class ITMGELServices
    {


        //path 路徑  finame檔名

        public static bool StartInsert(string path, string finame)
        {
            bool result = true;

            int i = 0;
            string line;
            List<ITEMGEL> iTEMGELs = new List<ITEMGEL>();
            ITEMGEL iTEMGEL = new ITEMGEL();
            try
            {
                Common.WriteLog("開始切割 " + finame);
                FileStream fileStream = File.OpenRead(path);
                using (StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(0)))
                {
                    line = reader.ReadLine();
                    //去除多空白

                    while (line != null)
                    {
                        line = new Regex("[\\s]+").Replace(line, " ");
                        if (line == " " || line == "")
                        {
                            line = reader.ReadLine();
                            continue;
                        }

                        iTEMGEL = splitstrng(line);
                        iTEMGELs.Add(iTEMGEL);
                        line = reader.ReadLine();

                        i++;

                    }

                    reader.Close();
                }

                Common.WriteLog("新增DB");
                result = InsertToSql(iTEMGELs);

                Common.WriteLog(finame + "切割完成，一共 " + i + " 筆");
                return result;
            }
            catch (Exception e)
            {
                Common.WriteLog("ITMGELServices失敗");
                Common.WriteLog(e.ToString());
                result = false;
                return result;

            }
          
           
        }


        //切割字串
        public static ITEMGEL splitstrng(string input)
        {



            ITEMGEL iTEMGEL = new ITEMGEL();
            try
            {
                //切割符號
                //string mark = ConfigurationManager.AppSettings["Mark"];
                input = new Regex("[\\s]+").Replace(input, " ");

                //string[] sArray = Regex.Split(input, @"|");

                string[] sArray = input.Split('|');

                iTEMGEL = new ITEMGEL
                {
                    Dept = sArray[0],
                    ItemCode = sArray[1],
                    CollectionNo = sArray[2],
                    SpecificItem = sArray[3],
                    LocalName = sArray[4],
                    EnglishName = sArray[5],
                    ShortName = sArray[6],
                    Capacity = sArray[7],
                    StockUnit = sArray[8],
                    MainSupplier = sArray[9],
                    DCSupplier = sArray[10],
                    VATRate = sArray[11],
                    InputDate = sArray[12],
                    Weight = sArray[13],
                    OnScale = sArray[14],
                    StopStartDate = sArray[15],
                    StopEndDate = sArray[16],
                    StopItemReason = sArray[17],
                    SellingCapacityUnit = sArray[18],
                    DisplayCapacityUnit = sArray[1],
                    SellingCapacity = sArray[20],
                    CapacityMultiplier = sArray[21],
                    DisplayCapacity = sArray[22],
                    Grade = sArray[23],
                    CountryofOrigin = sArray[24],
                    SPType = sArray[25],
                    TypeofSales = sArray[26],
                    SeasonCode = sArray[27],
                    Sensitiveness = sArray[28],
                    QtyPack = sArray[29],
                    RebateRate = sArray[30],
                    OrderDay = sArray[31],
                    OrderPeriod = sArray[32],
                    LeadTime = sArray[33],
                    DeliveryDays = sArray[34],
                    AttributeCodeCONS = sArray[35],
                    Status = sArray[36],
                    SubstituteTAX = sArray[37],
                    ItemGeneral = sArray[38],
                    ShelfLifeValue = sArray[39],
                    ShelfLifeUOM = sArray[40],
                    ShelfLifeLimitValue = sArray[41],
                    ShelfLifeLimitUOM = sArray[42],
                    ReturnableNonstopItem = sArray[43],
                    ReturnableTemporaryStop = sArray[44],
                    ReturnablePermanentStop = sArray[45],
                    UsebyDate = sArray[46],
                    UsebyDateUOM = sArray[47],
                    ReturnMixPacking = sArray[48],
                    WarehouseCheckingMark = sArray[49],
                    LogisticBoxDefine = sArray[50],
                    GoodExchangeCriteriaFlag = sArray[51],
                    ItemCategory = sArray[52],
                    TemporaryType = sArray[53],
                    LinkType = sArray[54],
                    LinkItemBarcode = sArray[55],
                    NonItemDescription = sArray[56],
                    BrandEnglish = sArray[57],
                    SubBrandEnglish = sArray[58],
                    BrandLocal = sArray[59],
                    SubBrandLocal = sArray[60],
                    ItemLevel = sArray[61],
                    FreshIDCardType = sArray[62],
                    FreshIDCard = sArray[63],
                    Manufacturer = sArray[64],
                    FoodAttribute = sArray[65],
                    AirproofAttribute = sArray[66],
                    TaxCategory = sArray[67],
                    SubTaxCategory = sArray[68],
                    QtyBox = sArray[69],
                    PackBox = sArray[70],
                    CNCode = sArray[71],
                    ItemTypologyFlag = sArray[72],
                    CountryofOrigin1 = sArray[73],
                    ItemTaxIdentifier = sArray[74]
                };

                return iTEMGEL;
            }
            catch (Exception ex)
            {
                Common.WriteLog("切割失敗 " + input + " ");
                Common.WriteLog(ex.ToString());
                return iTEMGEL;
                throw;
            }



        }

        //進DB
        private static bool InsertToSql(List<ITEMGEL> iTEMGELs)
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
                        string strsql = "Insert into  ITMGEL_TMP " +
                               "values(@Dept,@ItemCode,@CollectionNo,@SpecificItem,@LocalName,@EnglishName,@ShortName"
                               + ",@Capacity,@StockUnit,@MainSupplier,@DCSupplier,@VATRate,@InputDate,@Weight,@OnScale"
                               + ",@StopStartDate,@StopEndDate"
                               + ",@StopItemReason,@SellingCapacityUnit,@DisplayCapacityUnit,@SellingCapacity"
                               + ",@CapacityMultiplier,@DisplayCapacity,@Grade,@CountryofOrigin,@SPType,@TypeofSales"
                               + ",@SeasonCode,@Sensitiveness,@QtyPack,@RebateRate,@OrderDay,@OrderPeriod"
                               + ",@LeadTime,@DeliveryDays,@AttributeCodeCONS,@Status"
                               + ",@SubstituteTAX,@ItemGeneral,@ShelfLifeValue,@ShelfLifeUOM,@ShelfLifeLimitValue"
                               + ",@ShelfLifeLimitUOM,@ReturnableNonstopItem,@ReturnableTemporaryStop,@ReturnablePermanentStop"
                               + ",@UsebyDate,@UsebyDateUOM,@ReturnMixPacking,@WarehouseCheckingMark,@LogisticBoxDefine"
                               + ",@GoodExchangeCriteriaFlag,@ItemCategory,@TemporaryType,@LinkType"
                               + ",@LinkItemBarcode,@NonItemDescription,@BrandEnglish,@SubBrandEnglish"
                               + ",@BrandLocal,@SubBrandLocal,@ItemLevel"
                               + ",@FreshIDCardType,@FreshIDCard,@Manufacturer,@FoodAttribute,@AirproofAttribute"
                               + ",@TaxCategory,@SubTaxCategory,@QtyBox,@PackBox,@CNCode,@ItemTypologyFlag,@CountryofOrigin1"
                               + ",@ItemTaxIdentifier)";


                        conn.Execute(strsql, iTEMGELs, transaction);


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
