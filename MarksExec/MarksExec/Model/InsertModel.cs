using System;
using System.Collections.Generic;
using System.Text;

namespace MarksExec.Model
{
    public class ITEMGEL
    {
        public string Dept     { get; set; }
        public string ItemCode { get; set; }
        public string CollectionNo { get; set; }
        public string SpecificItem { get; set; }
        public string LocalName { get; set; }
        public string EnglishName { get; set; }
        public string ShortName { get; set; }
        public string Capacity { get; set; }
        public string StockUnit { get; set; }
        public string MainSupplier { get; set; }
        public string DCSupplier { get; set; }
        public string VATRate { get; set; }
        public string InputDate { get; set; }
        public string Weight { get; set; }
        public string OnScale { get; set; }
        public string StopStartDate { get; set; }
        public string StopEndDate { get; set; }
        public string StopItemReason { get; set; }
        public string SellingCapacityUnit { get; set; }
        public string DisplayCapacityUnit { get; set; }
        public string SellingCapacity { get; set; }
        public string CapacityMultiplier { get; set; }
        public string DisplayCapacity { get; set; }
        public string Grade { get; set; }
        public string CountryofOrigin { get; set; }
        public string SPType { get; set; }
        public string TypeofSales { get; set; }
        public string SeasonCode { get; set; }
        public string Sensitiveness { get; set; }
        public string QtyPack { get; set; }
        public string RebateRate { get; set; }
        public string OrderDay { get; set; }
        public string OrderPeriod { get; set; }
        public string LeadTime { get; set; }
        public string DeliveryDays { get; set; }
        public string AttributeCodeCONS { get; set; }
        public string Status { get; set; }
        public string SubstituteTAX { get; set; }
        public string ItemGeneral { get; set; }
        public string ShelfLifeValue { get; set; }
        public string ShelfLifeUOM { get; set; }
        public string ShelfLifeLimitValue { get; set; }
        public string ShelfLifeLimitUOM { get; set; }
        public string ReturnableNonstopItem { get; set; }
        public string ReturnableTemporaryStop { get; set; }
        public string ReturnablePermanentStop { get; set; }
        public string UsebyDate { get; set; }
        public string UsebyDateUOM { get; set; }
        public string ReturnMixPacking { get; set; }
        public string WarehouseCheckingMark { get; set; }
        public string LogisticBoxDefine { get; set; }
        public string GoodExchangeCriteriaFlag { get; set; }
        public string ItemCategory { get; set; }
        public string TemporaryType { get; set; }
        public string LinkType { get; set; }
        public string LinkItemBarcode { get; set; }
        public string NonItemDescription { get; set; }
        public string BrandEnglish { get; set; }
        public string SubBrandEnglish { get; set; }
        public string BrandLocal { get; set; }
        public string SubBrandLocal { get; set; }
        public string ItemLevel { get; set; }
        public string FreshIDCardType { get; set; }
        public string FreshIDCard { get; set; }
        public string Manufacturer { get; set; }
        public string FoodAttribute { get; set; }
        public string AirproofAttribute { get; set; }
        public string TaxCategory { get; set; }
        public string SubTaxCategory { get; set; }
        public string QtyBox { get; set; }
        public string PackBox { get; set; }
        public string CNCode { get; set; }
        public string ItemTypologyFlag { get; set; }
        public string CountryofOrigin1 { get; set; }
        public string ItemTaxIdentifier { get; set; }
    }



    public class STM_TMP
    {
        public string Dept { get; set; }
        public string StructureCode { get; set; }
        public string Cname { get; set; }
        public string Ename { get; set; }
        public string StructureQty { get; set; }
        public string UsedQty { get; set; }
        public string LogisticBoxRequire { get; set; }
        public string MultiTaxRate { get; set; }
        public string LicenseYN { get; set; }
        public string LicenseType { get; set; }
        public string FoodOrNonFood { get; set; }

    }



    public class SUPATT
    {
        public string StoreCode { get; set; }
        public string AttributeClassCode { get; set; }
        public string AttributeClassDescription { get; set; }
        public string AttributeCode { get; set; }
        public string AttributeCodeDescription { get; set; }
        public string AlphanumericValue { get; set; }
        public string NumberValue { get; set; }
        public string Date { get; set; }
        public string Time { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }

    }


    public class ITMSUB
    {
        public string Dept { get; set; }
        public string ItemCode { get; set; }
        public string SubID { get; set; }
        public string SubCode { get; set; }
        public string ItemID { get; set; }
        public string SubCodeEDes { get; set; }
        public string SubCodeCDes { get; set; }
        public string StartStopDate { get; set; }
        public string StopItemReasonCode { get; set; }
        public string ProductCharacteristic { get; set; }
        public string EndStopDate { get; set; }
        public string ReturnableforNonStopItem { get; set; }
        public string ReturnableforTemporaryStop { get; set; }
        public string ReturnableforPermanentStop { get; set; }
        public string WHCheckingMark { get; set; }
        public string ItemLevel { get; set; }
        public string UsebyDate { get; set; }
        public string CNCode { get; set; }

    }


    public class ITMSUPGEXCRTNC
    {
        public string DC { get; set; }
        public string dpno { get; set; }
        public string fullcode { get; set; }
        public string SupplierCode { get; set; }
        public string stopreasoncode { get; set; }
        public string Cname { get; set; }
        public string mainsupp { get; set; }
        public string realsupp { get; set; }
        public string suppCname { get; set; }
        public string typeofsales { get; set; }
        public string Grade { get; set; }
        public string spec { get; set; }
        public string seasoncode { get; set; }
        public string Permanentstop { get; set; }
        public string Supstop { get; set; }
        public string Supstopdate { get; set; }
        public string Goodexchangeflag { get; set; }
        public string OnlyBreakage { get; set; }
        public string Venderneedcheckproduct { get; set; }
        public string Goodexchangeflagforstopitem { get; set; }
        public string OnlyBreakage1 { get; set; }
        public string Venderneedcheckproduct1 { get; set; }
        public string ReturnFlagNonStopItem { get; set; }
        public string NoReturn { get; set; }
        public string AcceptBreakage { get; set; }
        public string OnlyGoodProduct { get; set; }
        public string OnlyBreakage_incl { get; set; }
        public string OnlyBreakage_ProductDefective { get; set; }
        public string OnlyIntactStockProduct { get; set; }
        public string MDSpecialReutrn { get; set; }
        public string SupplierSpecialDemands { get; set; }
        public string VenderNeedCheckProduct2 { get; set; }
        public string CompleteOriginalAccessaryDocumentA { get; set; }
        public string OriginalCoverBox { get; set; }
        public string OriginalProductPackage { get; set; }
        public string Return_Flag_TemporaryStopItem { get; set; }
        public string NoReturn1 { get; set; }
        public string Accept_BreakageandGood_Product1 { get; set; }
        public string OnlyGoodProduct1 { get; set; }
        public string OnlyBreakage_Near_Over_ExpireDat1 { get; set; }
        public string OnlyBreakage_ProductDefective1 { get; set; }
        public string OnlyIntactStockProduct1 { get; set; }
        public string MDSpecialReutrn1 { get; set; }
        public string SupplierSpecialDemands_RefertoLogis1 { get; set; }
        public string VenderNeedCheckProduct3 { get; set; }
        public string CompleteOriginalAccessaryDocument { get; set; }
        public string OriginalCoverBox1 { get; set; }
        public string OriginalProductPackage1 { get; set; }
        public string ReturnFlag_PermanentStopItem { get; set; }
        public string NoReturn2 { get; set; }
        public string AcceptBreakageandGoodProduct2 { get; set; }
        public string OnlyGoodProduct2 { get; set; }
        public string OnlyBreakage_Near_Over_ExpireDat2 { get; set; }
        public string OnlyBreakage_ProductDefective2 { get; set; }
        public string OnlyIntactStockProduct2 { get; set; }
        public string MDSpecialReutrn2 { get; set; }
        public string SupplierSpecialDemandsRefertoLogis2 { get; set; }
        public string VenderNeedCheckProduct4 { get; set; }
        public string CompleteOriginalAccessaryDocumentA2 { get; set; }
        public string OriginalCoverBox2 { get; set; }
        public string OriginalProductPackage2 { get; set; }


    }

    public class BAR
    {
        public string Dept { get; set; }
        public string ItemCode { get; set; }
        public string SubCode { get; set; }
        public string UnitCode { get; set; }
        public string Internal_ExternalBarcode { get; set; }
        public string EAN13Barcode { get; set; }
        public string OriginalBarcodeType { get; set; }
        public string OriginalBarcodeNo { get; set; }
        public string SNISpecialTypeBarcode { get; set; }
        public string MajorBarcode { get; set; }
        public string StartStopDate { get; set; }
        public string EndStopDate { get; set; }
        public string StopItemReason { get; set; }
        public string StopReasonDescription { get; set; }
        public string ItemID { get; set; }
        public string ItemSubcodeID { get; set; }
        public string Length { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string NotForSellFlag { get; set; }
        public string StoreCode { get; set; }
        public string NotForSellReason { get; set; }
        public string ItemCodeEnglishDescription { get; set; }
        public string ItemCodeChineseDescription { get; set; }
        public string Status { get; set; }

    }

    public class DailySales
    {
        public string StoreCode { get; set; }

        public string DepartmentCode { get; set; }

        public string ItemCode { get; set; }

        public string SubCode { get; set; }

        public string SalesDate { get; set; }

        public string UnitCode { get; set; }

        public string StatusPromotion { get; set; }

        public string SPType { get; set; }

        public string VatRate { get; set; }

        public string SalesQty { get; set; }

        public string SalesAmount { get; set; }

        public string SalesPrice { get; set; }

        public string PurchasePrice { get; set; }

        public string Rebate { get; set; }

        public string VatAmount { get; set; }

        public string OriginalSellingPrice { get; set; }


        public string CSTAmount { get; set; }


    }

    public class DailyStock
    {
        public string StoreCode { get; set; }
        public string DepartmentCode { get; set; }
        public string ItemCode { get; set; }
        public string SubCode { get; set; }
        public string TransactionDate { get; set; }
        public string SalesStatus { get; set; }
        public string BalanceQty { get; set; }
        public string Normal_OrderQty { get; set; }
        public string Promotion_OrderQty { get; set; }
        public string Normal_FreeGoods { get; set; }
        public string Promotion_FreeGoods { get; set; }
        public string Normal_ReceivedQty { get; set; }
        public string Promotion_ReceivedQty { get; set; }
        public string Normal_ReturnedQty { get; set; }
        public string Promotion_ReturnedQty { get; set; }
        public string SalesQty { get; set; }
        public string SalesAmount { get; set; }
        public string PurchaseAmount { get; set; }
        public string AdjustedQty { get; set; }
        public string BalanceForward { get; set; }
        public string Balance_Forward_OrderNotYetReceivedBasicgoods { get; set; }
        public string Balance_Forward_OrderNotYetReceivedFreegoods { get; set; }
        public string Rebate { get; set; }
        public string AvgCost { get; set; }
        public string AvgAmount { get; set; }
        public string BFAvgAmount { get; set; }
        public string PurchasePrice { get; set; }
        public string MainSuppliercode { get; set; }
        public string DSSuppliercode { get; set; }
        public string StopSubmonth { get; set; }
        public string StopSubyear { get; set; }
        public string StopSubreason { get; set; }

    }


    public class BigSales
    {
        public string YYYYMM { get; set; }
        public string Business_Date { get; set; }
        public string StoreName { get; set; }
        public string FullCode { get; set; }
        public string Item_Sales_Qty_Unit { get; set; }
        public string Item_Sales_Amount { get; set; }
        public string ItemCName { get; set; }
        public string Input { get; set; }

    }


    public class DSLDailySales
    {
        public string StoreCode { get; set; }

        public string DepartmentCode { get; set; }

        public string ItemCode { get; set; }

        public string SubCode { get; set; }

        public string SalesDate { get; set; }

        public string UnitCode { get; set; }

        public string StatusPromotion { get; set; }

        public string SPType { get; set; }

        public string VatRate { get; set; }

        public string SalesQty { get; set; }

        public string SalesAmount { get; set; }

        public string SalesPrice { get; set; }

        public string PurchasePrice { get; set; }

        public string Rebate { get; set; }

        public string VatAmount { get; set; }
        public string Input { get; set; }
        public string StoreName { get; set; }

        public string OriginalSellingPrice { get; set; }


        public string CSTAmount { get; set; }


    }

    public class SearchStoreCode
    {
        public string Store { get; set; }

        public string StoreCode { get; set; }
    }


    public class EMP_DATA
    {
        public string Emp_No { get; set; }
        public string Login_Name { get; set; }
        public string Emp_Name { get; set; }
        public string Mail_External { get; set; }
        public string Mail_Internal { get; set; }
        public string Entry_Date { get; set; }
        public string Dept_Code { get; set; }
        public string Org_code { get; set; }
        public string Org_NAME { get; set; }
        public string Job_Desc_Eng { get; set; }
        public string Job_Desc_Cht { get; set; }
        public string Direct_Manager { get; set; }
        public string Region { get; set; }
        public string REGION_NAME { get; set; }
        public string TW_sTORE_ID { get; set; }
        public string TW_sTORE_Name { get; set; }
        public string bu { get; set; }
        public string bu_NAME { get; set; }
        public string Job_code { get; set; }
        public string bank_no { get; set; }
        public string on_job  {get; set;}
        public string NIGHT_SHIFT { get; set; }
        public string SPEC_TYPE { get; set; }
        public string JOB_TITLE { get; set; }
        public string FT_PT { get; set; }
        public string JOB_LEVEL { get; set; }
        public string STORE_ID { get; set; }
        public string DEPT_ID { get; set; }
        public string MOD_TIME { get; set; }
        public string MOD_USER { get; set; }
        public string MOD_PGM { get; set; }
        public string MOD_WS { get; set; }
        public string STORE_NAME_E { get; set; }
        public string STORE_NAME_C { get; set; }
        public string ID_NO { get; set; }
        public string OUT_DATE { get; set; }
        public string SEX { get; set; }
        public string Birthday { get; set; }

    }

}
