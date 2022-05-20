using System;
using System.Collections.Generic;
using System.Text;

namespace CIMSClass
{
    public class GlobalString
    {
        /// <summary>
        /// 日志類型
        /// </summary>
        public static class LogType
        {
            public const string ErrorCategory = "ErrLog";
            public const string OpLogCategory = "OpLog";
            public const string DebugCategory = "DebugLog";
        }
        public static class SessionAndCookieKeys
        {
            public const string CULTURE_CODE = "PreferredCulture";
            public const string OUT = "Out";
            public const string USER = "User";
            public const string AGNET = "Agnet";
            public const string GROUPS = "Groups";
            public const string ACTIONS = "Actions";
            public const string RC_CODES = "RC_Codes";
            public const string USER_ID = "UserID";
            public const string MAINVIEW = "MainView";
            public const string LANG = "Lang";

        }
        public static class RST
        {
            public const string ACTIVED = "A";
            public const string UPDATE = "U";
            public const string DELETE = "D";
        }
        public static class RCU
        {
            public const string ACTIVED = "BATCH";
            
        }
        public static class RUU
        {
            public const string ACTIVED = "BATCH";
           
        }
        public static class FileSplit
        {
            //文件分割符號
            public const char Split = '^';
        }
        /// <summary>
        /// 批次狀態

        /// </summary>
        public static class BatchStatus
        {
            public const string Run = "Y";
            public const string Stop = "N";
        }
        /// <summary>
        /// 觸發時間類型
        /// </summary>
        public static class TriggerTimeType
        {
            public const string Year = "Y"; //滿足格式 MMdd HH:mm:ss
            public const string Month = "M"; //滿足格式 dd HH:mm:ss
            public const string Day = "D"; //滿足格式 HH:mm:ss
        }

        public static class YNType
        {
            public const string Yes = "Y";
            public const string No = "N";
        }

        //設定計算公式因子操作符的編號
        public static class Operation
        {
            //+
            public const string Add_RID = "+";
            //-
            public const string Del_RID = "-";
        }
        public static class TriggerTime
        {
            public static string BatchOne = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchOneTimeSpan = GlobalString.TriggerTimeType.Day;
            public static string BatchTwo = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchTwoTimeSpan = GlobalString.TriggerTimeType.Day;
            public static string BatchThree = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchThreeTimeSpan = GlobalString.TriggerTimeType.Day;
            public static string BatchFour = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchFourTimeSpan = GlobalString.TriggerTimeType.Day;
            public static string BatchFive = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchFiveTimeSpan = GlobalString.TriggerTimeType.Day;
            public static string BatchSix = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchSixTimeSpan = GlobalString.TriggerTimeType.Day;
            //200907CR
            public static string BatchSeven = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchSevenTimeSpan = GlobalString.TriggerTimeType.Day;
            public static string BatchEight = DateTime.Now.ToString("HH:mm:ss");
            public static string BatchEightTimeSpan = GlobalString.TriggerTimeType.Day;

        }
        /// <summary>
        /// 批次是否可運行標記

        /// </summary>
        public static class BatchRunFlag
        {
            public static string BatchOne = "N";
            public static string BatchTwo = "N";
            public static string BatchThree = "N";
            public static string BatchFour = "N";
            public static string BatchFive = "N";
            public static string BatchSix = "N";
            //200907CR
            public static string BatchSeven = "N";
            public static string BatchEight = "N";

        }
        public static class BatchRID
        {
            public const int BatchOne = 1;
            public const int BatchTwo = 2;
            public const int BatchThree = 3;
            public const int BatchFour = 4;
            public const int BatchFive = 5;
            public const int BatchSix = 6;
            //200907CR
            public const int BatchSeven = 7;
            public const int BatchEight = 8;
        }
        public static class ActFlag
        {
            public const string Do = "Y";
            public const string DoNot = "N";
        }
        public static class MailErroFlag
        {
            public const int BatchOne = 1;
            public const int BatchTwo = 2;
            public const int BatchThree = 3;
            public const int BatchFour = 4;
            public const int BatchFive = 5;
            public const int BatchSix = 6;
            //200907CR
            public const int BatchSeven = 7;
            public const int BatchEight = 8;
        }     
        /// <summary>
        /// 卡片參數子系列
        /// </summary>
        public static class cardparamType
        {
           // public const string card = "cardparam";

            //換卡百分比
            public const string Percent = "15";
            // 每月監控
            public const string NType = "7";
            public const string X1 = "8";
            public const string X2 = "9";
            public const string X3 = "10";
            // 每日監控
            public const string YType = "11";
            public const string Y1 = "12";
            public const string Y2 = "13";
            public const string Y3 = "14";
        }

        //設定計算公式的RID
        public static class Expression
        {
            //製成卡數
            public const int Made_RID = 1;
            //消耗卡數
            public const int Used_RID = 2;
            //耗損卡數
            public const int Waste_RID = 3;
            //入(出)庫
            public const int InOut_RID = 4;
            //特殊卡數
            public const int Special_RID = 5;
        }
        public static class Parameter
        {
            //制卡
            public const string Type = "use2";
            //賬務
            public const string Finance = "use1";
        }

        public static class ParameterType
        {
            //用途類參數
            public const string Use = "use";
            //賬務群組參數
            public const string Use1 = "use1";
            //卡片狀態
            public const string Status = "status";
            //卡別級距
            public const string CardType = "cardType";
            //物料類別
            //public const string MaterielType = "materielType";
            //物料類別縮寫
            public const string MatType = "matType";
            //物料類別
            public const string MatType1 = "matType1";
            //安全存量
            //public const string SafeType = "safeType";
            //送審狀態
            public const string SendCheck = "SendCheck";
            //類別
            public const string Type = "Type";
            //年份
            //public const string Year = "Year";
            //預估過去X月的月平均 
            //public const string xYear = "xYear";
            //使用者設定的希望庫存可以支援的可用月數參數類型標識
            //public const string NType = "NType";
            //廠商庫存異動匯入狀況代碼
            public const string ChangeStatus = "Change_Status";
            //放行狀態
            public const string PassStatus = "passStatus";
            //緊急程度
            //public const string EmergencyLevel = "emergencyLevel";

            //查詢所有參數類別時"ParamType_Code"的取值
            public const string Total = "total";
            //公式名稱
            public const string FormulaName = "FormulaName";
            //Operation
            public const string Operation = "Operation";
            //卡片管理參數設定
            public const string CardParam = "CardParam";
            //卡片成本帳務期間設定
            public const string CardCost = "CardCost";
            // 代製費用帳務異動項目設定 
            public const string Finance = "Finance";
            //操作類型
            public const string OprType = "OprType";
        }

        public static class SafeType
        {
            //最低安全庫存
            public const string storage = "1";
            //安全天數
            public const string days = "2";
            //用完為止
            public const string over = "3";
        }

        /// <summary>
        /// 定義不同FTP SERVER的前綴
        /// </summary>
        public static class FtpString
        {
            //小記檔
            public const string SUBTOTAL = "SubTotal";
            //廠商異動檔
            public const string CARDMODIFY = "CardModify";
            //廠商異動檔(替換前版面)
            public const string CARDMODIFYREPLACE = "CardModifyReplace";
            //年度換卡預測檔
            public const string YEARREPLACE = "YearReplace";
            //次月換卡預測檔
            public const string MONTHREPLACE = "MonthReplace";
            //廠商物料異動檔
            public const string MATERIAL = "Material";
            //特殊代制項目
            public const string SPECIAL = "Special";
        }

        public static class WarningType
        {
            public const string EditBugdet = "1";
            public const string BudgetAmtLower = "2";
            public const string BudgetCardLower = "3";
            public const string BudgetDateLower = "4";
            public const string EditAgreement = "5";
            public const string AgreementCardLower = "6";
            public const string AgreementDateLower = "7";
            public const string CardTypeEdit = "8";
            public const string CardTypeAdd = "9";
            public const string Perso_CardTypeEdit = "10";
            public const string AddPurchase = "21";
            public const string PlsAskFinance = "22";
            public const string SafeWarning = "16";
            public const string OrderFormCommit = "17";
            public const string OrderBudget = "18";
            public const string OrderAgreement = "19";
            public const string DepositoryBudget = "23";
            public const string DepositoryAgreement = "24";
            public const string CancelBudget = "25";
            public const string CancelAgreement = "26";
            public const string RestockBudget = "27";
            public const string RestockAgreement = "28";
            public const string MaterialDataIn = "11";
            public const string MaterialAutoDataIn = "12";
            public const string MaterialDataInMiss = "13";
            public const string MaterialDataInSafe = "14";
            public const string MaterialDataInLost = "30";
            public const string DayMonitory = "15";
            public const string MonthMonitory = "29";
            public const string PersoProjectSapAskMoney = "33";
            public const string PersoProjectChange = "34";
            public const string SubTotalDataIn = "35";
            public const string FactoryStocksChange = "36";
            public const string CardTypeNotEnough = "37";
            public const string BatchCompareNotPass = "38";
            public const string YearChangeCardForeCast = "39";
            public const string MonthChangeCardForeCast = "40";
            public const string MaterialBuget = "41";
            public const string SurplusMaterialBuget = "42";
            public const string SubtotalMaterialInMiss = "43";
            public const string SubtoalMaterialInSafe = "44";
            public const string PersoChangeMaterialInMiss = "45";
            public const string PersoChangeMaterialInSafe = "46";
            public const string PersoChangeCardInMiss = "47";
            public const string BatchPersoChangeMaterialInMiss = "48";
            public const string BatchPersoChangeMaterialInSafe = "49";
            public const string BatchPersoChangeCardInMiss = "50";
            public const string AskFinanceBudget = "51";
            public const string AskFinanceAgree = "52";
            public const string FactoryStocksChangeReplace = "53";
            public const string PersoChangeMaterialInMissReplace = "54";
            public const string PersoChangeMaterialInSafeReplace = "55";
            public const string PersoChangeCardInMissReplace = "56";
            public const string BatchPersoChangeMaterialInMissReplace = "57";
            public const string BatchPersoChangeMaterialInSafeReplace = "58";
            public const string BatchPersoChangeCardInMissReplace = "59";
            public const string BatchCompareFactoryReplace = "60";

        }
       

           
    }

}
