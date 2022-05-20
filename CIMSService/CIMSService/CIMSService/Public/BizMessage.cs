using System;
using System.Data;
using System.Configuration;


namespace CIMSBatch
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
    }
}
