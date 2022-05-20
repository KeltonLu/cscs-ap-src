using System;
using System.Data;
using System.Configuration;


namespace CIMSClass
{
    /// <summary>
    /// BussinessMessage 的摘要描述
    /// </summary>
    public class BizMessage
    {
        public BizMessage()
        {
            //
            // TODO: 在此加入建構函式的程式碼
            //
        }

        public static class BizMsg
        {
            public const string ALT_BASEINFO_001_01 = "預算{0}的金額過低，請重新填寫！";

            public const string ALT_BASEINFO_001_02 = "預算{0}的卡數過低，請重新填寫！";

            public const string ALT_BASEINFO_002_01 = "合約{0}的卡數過低，請重新填寫！";

            public const string ALT_BASEINFO_002_02 = "合約已被使用，有效期間（起）必須小於{0}！";

            public const string ALT_BASEINFO_002_03 = "合約已被使用，有效期間（迄）必須大於{0}！";

            public const string ALT_BASEINFO_041_01 = "該角色已賦予權限或被指定，無法刪除";

            public const string ALT_BASEINFO_042_01 = "用戶已被賦予權限，無法刪除";

            public const string ALT_CARDTYPE_001_01 = "該卡種群組已經存在";

            public const string ALT_CARDTYPE_006_01 = "製卡類別已經存在";

            public const string ALT_CARDTYPE_007_01 = "主機檔案名稱已存在";

            public const string ALT_DEPOSITORY_003_01 = "匯入文件格式不正確";

            public const string ALT_DEPOSITORY_003_02 = "訂單流水編號不存在";

            public const string ALT_DEPOSITORY_003_03 = "不可使用該空白卡廠";

            public const string ALT_DEPOSITORY_003_04 = "不可使用該版面簡稱";

            public const string ALT_DEPOSITORY_003_05 = "晶片不存在";

            public const string ALT_DEPOSITORY_003_06 = "批號不存在";

            public const string ALT_DEPOSITORY_003_07 = "匯入重複的流水編號";

            public const string ALT_DEPOSITORY_003_08 = "不可使用該Person廠商";

            public const string ALT_DEPOSITORY_003_09 = "ID過長";

            public const string ALT_DEPOSITORY_003_10 = "訂單流水編號未放行";

            public const string ALT_DEPOSITORY_003_11 = "訂單流水編號已結案";

            public const string ALT_DEPOSITORY_003_12 = "入庫量不能為負數";

            public const string ALT_DEPOSITORY_003_13 = "訂單流水編號{0},預算金額(卡數)餘額不足,不可入庫";

            public const string ALT_DEPOSITORY_003_14 = "訂單流水編號{0},合約卡數餘額不足,不可入庫";

            public const string ALT_DEPOSITORY_010_01 = "Perso廠不一致";

            public const string ALT_DEPOSITORY_010_02 = "匯入文件排序錯誤";

            public const string ALT_DEPOSITORY_010_03 = "檔案格式錯誤，[品名]只能有一筆結餘。";

            public const string ALT_DEPOSITORY_010_04 = "檔案格式錯誤，[品名]（進貨/退貨/銷毀）日期xxxx/xx/xx大於結餘日期xxxx/xx/xx。";

            public const string ALT_DEPOSITORY_010_05 = "檔案格式錯誤，[品名]結餘日期或異動日期不能小于上次結餘日期。";

            public const string ALT_DEPOSITORY_010_06 = "檔案格式錯誤，紙品物料結餘日期不一致";

            public const string ALT_DEPOSITORY_010_07 = "檔案格式錯誤，[結餘日期]尚未日結。";

            public const string ALT_DEPOSITORY_010_08 = "的庫存結餘大與廠商結餘";

            public const string ALT_DEPOSITORY_010_09 = "刪除廠商結餘日期不是最後一筆。";

            public const string ALT_DEPOSITORY_010_10 = "匯入資料時出錯誤";

            public const string ALT_DEPOSITORY_004_01 = "記錄的退貨量已超過現有的庫存量，請重新填寫";
            public const string ALT_DEPOSITORY_004_02 = "該退貨記錄的退貨量已超過現有的庫存量，請重新填寫";
            public const string ALT_DEPOSITORY_004_03 = "合約剩余數量不足";
            public const string ALT_DEPOSITORY_004_04 = "預算剩余數量不足";
            public const string ALT_DEPOSITORY_004_05 = "預算剩余金額不足";

        }

        public static class BizCommMsg
        {
            public const string ALT_CMN_CanntDel = "不可刪除該筆資料";

            public const string ALT_CMN_CanntModify = "不可修改該筆資料";

            public const string ALT_CMN_DeleteFail = "刪除失敗";

            public const string ALT_CMN_DeleteSucc = "刪除成功";

            public const string ALT_CMN_GetModelFail = "獲取模型失敗";

            public const string ALT_CMN_InitPageFail = "初始化頁面失敗";

            public const string ALT_CMN_SaveSucc = "儲存成功";

            public const string ALT_CMN_SaveFail = "儲存失敗，請稍後再按一次儲存";

            public const string ALT_CMN_SearchFail = "查詢失敗";

            public const string ALT_CMN_CannotDel = "{0}已被使用，無法刪除！";
        }

        public static class BizPublicMsg
        {
            public const string ALT_GetIDFail = "獲取流水號失敗";

            public const string ALT_IDTooLong = "ID超過檢索範圍";

            public const string ALT_GetMenuFail = "獲取Menu資料失敗";

            public const string ALT_GetUserFunctionFail = "查詢用戶可用的功能編碼異常";

            public const string ALT_ReadPageActionFail = "讀取頁面權限組態文件失敗";

            public const string ALT_PicUpdSucess = "圖片上傳成功";

            public const string ALT_PicTooLarge = "上傳圖片不能大於10M";

            public const string ALT_PicUploadFail = "上傳圖片失敗";

            public const string ALT_PicWrongFormat = "上傳圖片格式必須為bmp、jpg、gif、tif";

            public const string ALT_HasNoAction = "用戶沒有訪問本頁面的權限";

            public const string ALT_HasNoAuth = "您還未獲得授權使用本系統";

            public const string ALT_NotLogin = "Session到期，請重新登錄";

            public const string ALT_NoCtrlAction = "控制項權限組台未設定行爲碼";




        }
        public static class GlobalStringResource
        {
            public const string Alert_CanntDelete = "不可刪除該筆資料";
            public const string Alert_CanntModify = "不可修改該筆資料";
            public const string Alert_CardGroupExist = "該卡種群組已經存在！";
            public const string Alert_CardType_Used = "尚有卡種使用中";
            public const string Alert_ComposeColumn_Duplited = "新增欄位名稱不可重複";
            public const string Alert_ConNotLinkLDAP = "連接LDAP Server失敗";
            public const string Alert_CtrlNoneAction = "控製項權限組台未設定行爲碼";
            public const string alert_CustomerYesORNO = "查無此客戶資料";
            public const string Alert_DangerousInputCode = "您輸入了危險字符，本次操作無效！";
            public const string Alert_DataExist = "該資料已經存在，如需變更請使用修改功能。";
            public const string Alert_DeleteFailErr = "刪除失敗";
            public const string Alert_DeleteSuccess = "刪除成功";
            public const string Alert_FactoryCancelMarkDelete = "沒有選擇取消標記刪除的對象";
            public const string Alert_FactoryCustomerID = "授信用戶編號：";
            public const string Alert_FactoryCustomerName = "授信用戶名稱：";
            public const string Alert_FactoryDelete = "沒有選擇要刪除的對象!";
            public const string Alert_FactoryMarkDelete = "沒有選項擇標記刪除的對象";
            public const string Alert_FactoryMessage = "有海外工廠資料";
            public const string Alert_FileTypeErr = "文件內容不正確，請查明。";
            public const string Alert_FileUploadErr = "文件上傳失敗";
            public const string Alert_GetList_ImportHistoryErr = "查詢資料導入信息";
            public const string Alert_Group_AddError = "添加群組失敗";
            public const string Alert_HasNotGroup = "您還未獲得授權使用本系統";
            public const string Alert_HasNotPageAction = "用戶沒有訪問本頁面的權限";
            public const string Alert_Info_DateInputErr = "輸入的日期不合邏輯";
            public const string Alert_InitPageFailErr = "初始化頁面失敗";
            public const string Alert_Input_CustomerID = "請輸入“授信戶統編/企業團編號”";
            public const string Alert_Input_RemitterName = "請輸入“匯款人名稱”";
            public const string Alert_InputReason = "請輸入維護原因";
            public const string Alert_LoginLDAPErr = "LDAP登錄失敗";
            public const string Alert_LoginLogTimeErr = "記錄用戶登錄時間失敗";
            public const string Alert_LostLogin = "用戶驗證未通過";
            public const string Alert_ManualCheckLogMessage = "有查核資料內容 ";
            public const string Alert_MarkDeleteSuccess = "標記刪除成功";
            public const string Alert_MustInput = "所有欄位不可空白";
            public const string Alert_NOCustomerID = "請輸入授信用戶編號!";
            public const string Alert_NoneAction = "對不起，此帳號未被設定任何功能權限，請查明。";
            public const string Alert_NoneAgent = "對不起，要代理的用戶不存在。";
            public const string Alert_NoneAgentAction = "對不起，要代理的用戶沒有被設定任何權限，請查明。";
            public const string Alert_NoneRC_Code = "您還沒有區域中心的設置";
            public const string Alert_NoneResourceItem = "資源中沒有{0}的鍵";
            public const string Alert_NoneUser = "對不起，用戶不存在。";
            public const string Alert_NoPic = "請上傳圖片";
            public const string Alert_NotLogin = "用戶未登錄";
            public const string Alert_Numberic_Required = "必須是數字";
            public const string Alert_PassWordErr = "對不起，密碼錯誤。";
            public const string Alert_PicTooLarge = "圖片上傳失敗,大小不能超過10M";
            public const string Alert_PicUploadFail = "圖片上傳失敗";
            public const string Alert_PicUploadSucess = "圖片上傳成功";
            public const string Alert_PicWrongFormat = "圖片上傳失敗，格式應該為gif jpg jpeg bmp";
            public const string alert_RightCustomerName = "請輸入正確的用戶編號";
            public const string Alert_SalesCustomerMessage = "有銷售廠商資料";
            public const string Alert_SaveFailErr = "儲存失敗，請稍後再按一次儲存";
            public const string Alert_SaveSuccess = "儲存成功";
            public const string Alert_SearchFailErr = "查詢失敗";
            public const string Alert_SQL_NotSelect = "只能輸入查詢語句";
            public const string Alert_SQL_ValidateErr = "SQL語句錯誤";
            public const string Alert_SQL_ValidateSuccess = "SQL語句正確";
            public const string Alert_StockCustomerMessage = "有進貨廠商資料";
            public const string Alert_TailorReport_Exist = "報表名稱已存在";
            public const string Alert_TailorReport_GetDataErr = "獲得頁面資料失敗";
            public const string Alert_TailorReport_NotExist = "嘗試查詢不存在的報表";
            public const string Alert_UserAcountISUseing = "您使用的帳號已經被其它地方佔用，請查明";
            public const string AlertRoleCantDel = "該角色已賦予權限或被指定，無法刪除";
            public const string AlertUserCannotDel = "用戶已被賦予權限，無法刪除";
            public const string ArrayLengthErr = "比較的長度不同";
            public const string BTN_UseAndFill = "套用";
            public const string DataSettingErr = "系統參數缺少記錄";
            public const string Format_Date = "/";
            public const string Info_All = "全部";
            public const string Info_ButtonAdd = "新增";
            public const string Info_ButtonUpdate = "更新";
            public const string Info_CancelMarkDeleteHint = "是否取消標記刪除此筆記錄";
            public const string Info_ContainsUserErr = "檢查用戶是否存在，失敗";
            public const string Info_DeleteHint = "是否刪除此筆訊息";
            public const string Info_ISBeingRecord = "已存在該筆資料";
            public const string Info_MarkDeleteHint = "是否標記刪除此筆記錄";
            public const string TradeType_Value1 = "All";
            public const string TradeType_Value2 = "InCome";
            public const string TradeType_Value3 = "Trans";

        }

        public static class BusinessBaseResource
        {
            public const string Alert_GetIDFailErr = "獲取流水號失敗";
            public const string Alert_GetMenuDataErr = "獲取Menu資料失敗";
            public const string Alert_GetUserFunctionErr = "查詢用戶可用的功能編碼異常";
            public const string Alert_NoneResourceItem = "資源中沒有{0}的鍵";
            public const string Alert_ReadPageActionConfigErr = "讀取頁面權限組態文件失敗";
            public const string ALT_041_CanDel = "該角色已賦予權限或被指定，無法刪除";
            public const string ALT_042_CanDel = "用戶已被賦予權限，無法刪除";
            public const string ALT_CMN_CanntDel = "不可刪除該筆資料";
            public const string ALT_CMN_CanntModify = "不可修改該筆資料";
            public const string ALT_CMN_DeleteFail = "刪除失敗";
            public const string ALT_CMN_GetModelFail = "獲取模型失敗";
            public const string ALT_CMN_InitPageFail = "初始化頁面失敗";
            public const string ALT_CMN_SaveFail = "儲存失敗，請稍後再按一次儲存";
            public const string ALT_CMN_SearchFail = "查詢失敗";
        }
    }
}
