//*****************************************
//*  作    者：GaoAi
//*  功能說明：廠商庫存異動匯入
//*  創建日期：2008-11-19
//*  修改日期：
//*  修改記錄：
//*****************************************

//**************************using System;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Text.RegularExpressions;
using System.IO;
using CIMSBatch.Model;
using CIMSBatch.FTP;
using CIMSBatch.Mail;
using CIMSBatch.Public;
using CIMSClass.Business;

namespace CIMSBatch.Business
{
    class InOut002BL : BaseLogic
    {
        #region SQL
        //選擇所有的廠商資料！
        public const string SEL_FACTORY = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN "
                    + "FROM FACTORY AS F "
                    + "WHERE F.RST = 'A' AND F.Is_Perso = 'Y' order by RID";




        public const string SEL_CARDTYPE_PERSO_FACTORY = "SELECT * FROM FACTORY "
                                                        + "WHERE RST = 'A' AND Is_Perso='Y' AND Factory_ShortName_EN = @Factory_ShortName_EN";
        public const string CON_CARDTYPE_SURPLUS = "SELECT COUNT(*) FROM CARDTYPE_STOCKS WHERE RST='A' AND Stock_Date>=@Stock_Date";
        public const string CON_IMPORT_CARDTYPE_CHANGE_CHECK = "SELECT FCI.* FROM FACTORY_"
                        + "CHANGE_IMPORT FCI LEFT JOIN FACTORY F ON FCI.Perso_Factory_RID = F.RID AND F.RST = 'A'"
                             + " WHERE FCI.RST='A' AND F.Factory_ID = @Factory_ID AND (CONVERT(char(10), FCI.Date_Time, 111) = @Date_Time)";
        public const string CON_CARDTYPE_STATUS = "SELECT COUNT(*) FROM CARDTYPE_STATUS WHERE RST='A'";
        public const string SEL_CARD_TYPE = "SELECT * FROM CARD_TYPE WHERE RST='A'";
        public const string SEL_FACTORY_CHANGE_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC INNER JOIN WARNING_USER WU "
            + "ON WC.RID = WU.Warning_RID LEFT JOIN USERS U ON  WU.UserID = U.UserID "
            + "WHERE WC.RST = 'A' AND WC.RID =";
        public const string SEL_CHECK_DATE = "SELECT RID "
                                   + "FROM CARDTYPE_STOCKS "
                                   + "WHERE RST='A' AND Stock_Date = @check_date";
        public const string SEL_FACTORY_CHANGE_IMPORT_ALL = "SELECT FCI.Space_Short_Name,CS.Status_Code,FCI.Perso_Factory_RID,FCI.Date_Time "
                                  + "FROM FACTORY_CHANGE_IMPORT FCI LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID "
                                  + "WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @FactoryRID AND "
                                  + " FCI.Date_Time >= @Import_Date_Start AND FCI.Date_Time <= @Import_Date_End";
        public const string SEL_CARDTYPE_STATUS = "SELECT RID,Status_Code,Status_Name "
                                 + "FROM CARDTYPE_STATUS "
                                 + "WHERE RST='A' ";
        public const string SEL_CARDTYPE_End_Time = "SELECT TYPE,AFFINITY,PHOTO,Name,End_Time,Is_Using "
                             + "FROM CARD_TYPE "
                             + "WHERE RST='A'";
        public const string SEL_FACTORY_RID = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN "
                                        + "FROM FACTORY AS F "
                                        + "WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND Factory_ID = @factory_id order by RID";
        // 物料消耗報警
        public const string SEL_MADE_CARD_WARNNING = "SELECT FCI.Perso_Factory_RID,FCI.CareType_RID as RID,FCI.Status_RID,SUM(FCI.Number) AS Number "
                            + "FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID "
                            + "WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time "
                            + "GROUP BY FCI.Perso_Factory_RID,FCI.CareType_RID,FCI.Status_RID";
        public const string SEL_EXPRESSIONS_DEFINE_WARNNING = "SELECT ED.Operate,CS.RID "
                            + "FROM EXPRESSIONS_DEFINE ED INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND ED.Type_RID = CS.RID "
                            + "WHERE ED.RST = 'A' AND ED.Expressions_RID = 1";

        public const string DEL_TEMP_MADE_CARD = "DELETE FROM TEMP_MADE_CARD "
                            + "WHERE Perso_Factory_RID = @perso_factory_rid";

        public const string INSERT_INTO_TEMP_MADE_CARD = "INSERT INTO TEMP_MADE_CARD(Perso_Factory_RID,CardType_RID,Number)values("
                            + "@Perso_Factory_RID,@CardType_RID,@Number)";

        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD = " SELECT EI.Serial_Number AS EI_Number,CE.Serial_Number as CE_Number,TMC.Perso_Factory_RID,TMC.Number " +
                            "FROM TEMP_MADE_CARD TMC " +
                            "INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND TMC.CardType_RID = CT.RID " +
                            "INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID " +
                            "INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID " +
                            "WHERE TMC.Perso_Factory_RID = @perso_factory_rid";
        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD_DM = " SELECT DI.Serial_Number DI_Number,A.Perso_Factory_RID,A.Number " +
                            "FROM TEMP_MADE_CARD A " +
                            "INNER JOIN DM_CARDTYPE DCT ON DCT.RST = 'A' AND A.CardType_RID = DCT.CardType_RID " +
                            "INNER JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND DCT.DM_RID = DI.RID " +
                            "WHERE A.Perso_Factory_RID = @perso_factory_rid";
        public const string SEL_LAST_WORK_DATE = "SELECT TOP 1 Date_Time " +
                            "FROM WORK_DATE " +
                            "WHERE Date_Time < @date_time AND Is_WorkDay='Y' " +
                            "ORDER BY Date_Time DESC";
        public const string SEL_MATERIEL_STOCKS_MANAGER = "SELECT Top 1 MSM.Stock_Date,MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number," +
                            "CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME " +
                            "FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number " +
                                "LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number " +
                                "LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number " +
                            "WHERE Type = '4' AND MSM.Perso_Factory_RID = @perso_factory_rid AND MSM.Serial_Number = @serial_number " +
                          "ORDER BY Stock_Date Desc";
        public const string SEL_MATERIEL_USED = "SELECT SUM(Number) as Number FROM MATERIEL_STOCKS_USED " +
                            "WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Serial_Number = @serial_number " +
                            " AND Stock_Date>@from_stock_date AND Stock_Date<=@end_stock_date ";

        public const string SEL_LAST_SURPLUS_DAY = "SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS WHERE RST = 'A' ORDER BY  Stock_Date DESC";
        public const string SEL_ENVELOPE_INFO = "SELECT * "
                                    + "FROM ENVELOPE_INFO "
                                    + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_CARD_EXPONENT = "SELECT * "
                                        + "FROM CARD_EXPONENT "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_DMTYPE_INFO = "SELECT * "
                                        + "FROM DMTYPE_INFO "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_MATERIEL_STOCKS_USED = "select * from MATERIEL_STOCKS_USED where rst='A' "
                                        + "AND Serial_Number=@Serial_Number "
                                        + "AND Perso_Factory_RID=@Perso_Factory_RID "
                                        + "AND Stock_Date > @lastSurplusDateTime "
                                        + "AND Stock_Date <= @thisSurplusDateTime";
        #endregion

        #region 構造函數
        public DataTable dtPeso;
        public string strErr;
        //public ArrayList erro;
        //參數

        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        public InOut002BL()
        {
            //
            // TODO: 在此加入建構函式的程式碼
            //
            string ftpPath = ConfigurationManager.AppSettings["FTPCardModify"];
            string localPath = ConfigurationManager.AppSettings["FTPCardModifyPath"];
        }
        #endregion

        #region 下載檔案
        public ArrayList DownLoadModify(string FactoryPath)
        {
            ArrayList FileNameList = new ArrayList();
            try  // 加try catch add by judy 2018/03/28
            {
                string FolderYear = DateTime.Now.ToString("yyyy");
                string FolderDate = DateTime.Now.ToString("MMdd");
                string ftpPath = ConfigurationManager.AppSettings["FTPCardModify"] + "/" + FactoryPath;
                string localPath = ConfigurationManager.AppSettings["FTPCardModifyPath"];
                string FolderName = "";
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.CARDMODIFY);
                string[] fileList;
                string[] fileMethod;
                bool returnFlag;

                fileList = ftp.GetFileList(ftpPath);
                if (fileList != null)
                {
                    foreach (string FileName in fileList)
                    {
                        if (!CheckFile(FileName)) //檢查要下載的檔案規則，不滿足則跳過
                        {
                            continue;
                        }
                        fileMethod = FileName.Split('-');
                        if (fileMethod != null)
                        {
                            FolderName = localPath + "\\" + fileMethod[1] + "\\";

                            returnFlag = ftp.Download(ftpPath, FileName, FolderName, FileName);
                            if (returnFlag)
                            {
                                string[] FList = new string[2];
                                FList[0] = FolderName;
                                FList[1] = FileName;
                                FileNameList.Add(FList);
                                // Legend 2017/11/28 將此處刪檔注釋做UAT測試, 上線是再解開 todo
                                returnFlag = ftp.Delete(ftpPath, FileName);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogFactory.Write("匯入廠商異動檔下載FTP檔案DownLoadModify方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
           
            return FileNameList;

        }
        #endregion

        #region 物料消耗報警
        /// <summary>
        /// 根據匯入的小計檔，生成物料消耗記錄，并判斷物料庫存是否在安全水位，
        /// 如果不在安全水位，報警。
        /// </summary>
        /// <param name="strFactory_RID"></param>
        /// <param name="importDate"></param>
        //public void Material_Used_Warnning(string strFactory_RID,
        //    DateTime importDate)
        //{
        //    try
        //    {
        //        // 取最後日結日期。
        //        DateTime TheLastestSurplusDate = getLastSurplusDate();

        //        #region 計算從最後一個日結日期的下一天到資料匯入日期的卡片製成數，保存到臨時表(TEMP_MADE_CARD)
        //        dirValues.Clear();
        //        dirValues.Add("Perso_Factory_RID", strFactory_RID);
        //        dirValues.Add("From_Date_Time", TheLastestSurplusDate.ToString("yyyy/MM/dd 23:59:59"));
        //        dirValues.Add("End_Date_Time", importDate.ToString("yyyy/MM/dd 23:59:59"));
        //        DataSet dsMade_Card = dao.GetList(SEL_MADE_CARD_WARNNING, dirValues);
        //        DataSet dsEXPRESSIONS_DEFINE = dao.GetList(SEL_EXPRESSIONS_DEFINE_WARNNING);

        //        //卡種消耗表
        //        DataTable dtUSE_CARDTYPE = new DataTable();
        //        dtUSE_CARDTYPE.Columns.Add("Perso_Factory_RID");
        //        dtUSE_CARDTYPE.Columns.Add("CardType_RID");
        //        dtUSE_CARDTYPE.Columns.Add("Number");

        //        //按Perso廠、卡種的計算消耗量（循環加總各種狀況的消耗數量）
        //        int Card_Type_Rid = 0;
        //        int Perso_Factory_RID = 0;
        //        int Number = 0;
        //        //todo 此循環可以改進為存儲過程
        //        foreach (DataRow dr in dsMade_Card.Tables[0].Rows)
        //        {
        //            if ((Convert.ToInt32(dr["RID"]) != Card_Type_Rid) ||
        //                (Convert.ToInt32(dr["Perso_Factory_RID"]) != Perso_Factory_RID))
        //            {
        //                if (Card_Type_Rid != 0 && Perso_Factory_RID != 0 && Number != 0)
        //                {
        //                    DataRow drow = dtUSE_CARDTYPE.NewRow();
        //                    drow["Number"] = Number.ToString();
        //                    drow["Perso_Factory_RID"] = Perso_Factory_RID.ToString();
        //                    drow["CardType_RID"] = Card_Type_Rid.ToString();
        //                    dtUSE_CARDTYPE.Rows.Add(drow);
        //                }

        //                #region 取消耗卡公式,計算消耗卡數
        //                Number = 0;
        //                DataRow[] drEXPRESSIONS = dsEXPRESSIONS_DEFINE.Tables[0].Select("RID = " + dr["Status_RID"].ToString());
        //                if (drEXPRESSIONS.Length > 0)
        //                {
        //                    if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Add_RID)
        //                    {
        //                        Number += Convert.ToInt32(dr["Number"]);
        //                        Card_Type_Rid = Convert.ToInt32(dr["RID"]);
        //                        Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                    }
        //                    else if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Del_RID)
        //                    {
        //                        Number -= Convert.ToInt32(dr["Number"]);
        //                        Card_Type_Rid = Convert.ToInt32(dr["RID"]);
        //                        Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                    }
        //                }
        //                #endregion
        //            }
        //            else
        //            {
        //                #region 取消耗卡公式,計算消耗卡數
        //                DataRow[] drEXPRESSIONS = dsEXPRESSIONS_DEFINE.Tables[0].Select("RID = " + dr["Status_RID"].ToString());
        //                if (drEXPRESSIONS.Length > 0)
        //                {
        //                    if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Add_RID)
        //                    {
        //                        Number += Convert.ToInt32(dr["Number"]);
        //                    }
        //                    else if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Del_RID)
        //                    {
        //                        Number -= Convert.ToInt32(dr["Number"]);
        //                    }
        //                }
        //                #endregion
        //            }
        //        }
        //        if (Card_Type_Rid != 0 && Perso_Factory_RID != 0 && Number != 0)
        //        {
        //            DataRow drow = dtUSE_CARDTYPE.NewRow();
        //            drow["Number"] = Number.ToString();
        //            drow["Perso_Factory_RID"] = Perso_Factory_RID.ToString();
        //            drow["CardType_RID"] = Card_Type_Rid.ToString();
        //            dtUSE_CARDTYPE.Rows.Add(drow);
        //        }

        //        // 刪除臨時表中的數據
        //        dirValues.Clear();
        //        dirValues.Add("perso_factory_rid", strFactory_RID);
        //        dao.ExecuteNonQuery(DEL_TEMP_MADE_CARD, dirValues);

        //        foreach (DataRow dr in dtUSE_CARDTYPE.Rows)
        //        {
        //            dirValues.Clear();
        //            dirValues.Add("Perso_Factory_RID", dr["Perso_Factory_RID"].ToString());
        //            dirValues.Add("CardType_RID", dr["CardType_RID"].ToString());
        //            dirValues.Add("Number", dr["Number"].ToString());
        //            dao.ExecuteNonQuery(INSERT_INTO_TEMP_MADE_CARD, dirValues);
        //        }

        //        #endregion 計算當天製成卡數

        //        // 根據製成卡數，計算物料消耗
        //        //200908IR耗用量=小計檔數量*（1+耗損率） add  by 楊昆 2009/09/11 start
        //        InOut001BL BL001 = new InOut001BL();
        //        DataTable dtMATERIAL_USED = BL001.getMaterialUsed(strFactory_RID, importDate);
        //        //200908IR耗用量=小計檔數量*（1+耗損率） add  by 楊昆 2009/09/11 end

        //        // 計算物料剩余數量并警示
        //        getMaterielStocks(TheLastestSurplusDate,
        //            strFactory_RID,
        //            importDate,
        //            dtMATERIAL_USED);
        //    }
        //    catch (Exception ex)
        //    {
        //        //strReturn = "Error";
        //        LogFactory.Write(ex.ToString(), GlobalString.LogType.ErrorCategory);
        //        BatchBL Bbl = new BatchBL();
        //        Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
        //   }
        //}

        /// <summary>
        /// 計算物料剩余數量并警示
        /// </summary>
        /// <param name="strFactory_RID"></param>
        /// <param name="importDate"></param>
        /// <param name="dtMATERIAL_USED"></param>
        //public void getMaterielStocks(DateTime dtLastWorkDate,
        //        string strFactory_RID,
        //        DateTime importDate,
        //        DataTable dtMATERIAL_USED)
        //{
        //    try
        //    {
        //        Depository010BL bl010 = new Depository010BL();

        //        #region 根據前一天的庫存及今天的庫存。計算物料剩餘數量，判斷是否報警
        //        foreach (DataRow drMATERIAL_USED in dtMATERIAL_USED.Rows)
        //        {
        //            dirValues.Clear();
        //            dirValues.Add("perso_factory_rid", strFactory_RID);
        //            dirValues.Add("serial_number", drMATERIAL_USED["Serial_Number"].ToString());
        //            DataSet dsMaterielStocksManager = dao.GetList(SEL_MATERIEL_STOCKS_MANAGER, dirValues);
        //            if (null != dsMaterielStocksManager &&
        //                dsMaterielStocksManager.Tables.Count > 0 &&
        //                dsMaterielStocksManager.Tables[0].Rows.Count > 0)
        //            {
        //                // 從盤整日到日結日，耗用
        //                dirValues.Clear();
        //                dirValues.Add("perso_factory_rid", strFactory_RID);
        //                dirValues.Add("serial_number", drMATERIAL_USED["Serial_Number"].ToString());
        //                dirValues.Add("from_stock_date", Convert.ToDateTime(dsMaterielStocksManager.Tables[0].Rows[0]["Stock_Date"]).ToString("yyyy/MM/dd 23:59:59"));
        //                dirValues.Add("end_stock_date", dtLastWorkDate.ToString("yyyy/MM/dd 23:59:59"));
        //                DataSet dsUsedMaterial = dao.GetList(SEL_MATERIEL_USED, dirValues);
        //                if (null != dsUsedMaterial &&
        //                    dsUsedMaterial.Tables.Count > 0 &&
        //                    dsUsedMaterial.Tables[0].Rows.Count > 0)
        //                {
        //                    // 盤整時的庫存
        //                    int intLastStockNumber = Convert.ToInt32(dsMaterielStocksManager.Tables[0].Rows[0]["Number"].ToString());
        //                    // 從盤整日到最結餘日的消耗
        //                    int intUsedMaterialFront = 0;
        //                    if (dsUsedMaterial.Tables[0].Rows[0]["Number"] != DBNull.Value)
        //                        intUsedMaterialFront = Convert.ToInt32(dsUsedMaterial.Tables[0].Rows[0]["Number"]);

        //                    // 最後結餘日后的消耗
        //                    int intUsedMaterialAfter = Convert.ToInt32(drMATERIAL_USED["Number"]);

        //                    // 庫存為0時，顯示庫存不足
        //                    if (intLastStockNumber <= 0)
        //                    {
        //                        if (bl010.DmNotSafe_Type(drMATERIAL_USED["Serial_Number"].ToString()))
        //                        {
        //                            // 庫存不足
        //                            string[] arg = new string[1];
        //                            arg[0] = dsMaterielStocksManager.Tables[0].Rows[0]["Name"].ToString();
        //                            Warning.SetWarning(GlobalString.WarningType.SubtotalMaterialInMiss, arg);
        //                        }
        //                    }
        //                    // 如果前一天的庫存小余今天的消耗
        //                    else if (intLastStockNumber < (intUsedMaterialFront + intUsedMaterialAfter))
        //                    {
        //                        if (bl010.DmNotSafe_Type(drMATERIAL_USED["Serial_Number"].ToString()))
        //                        {
        //                            // 庫存不足
        //                            string[] arg = new string[1];
        //                            arg[0] = dsMaterielStocksManager.Tables[0].Rows[0]["Name"].ToString();
        //                            Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInMiss, arg);
        //                            Warning.SetWarning(GlobalString.WarningType.SubtotalMaterialInMiss, arg);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        // 取物料的安全庫存訊息
        //                        DataSet dtMateriel = this.GetMateriel(drMATERIAL_USED["Serial_Number"].ToString());
        //                        if (null != dtMateriel &&
        //                            dtMateriel.Tables.Count > 0 &&
        //                            dtMateriel.Tables[0].Rows.Count > 0)
        //                        {
        //                            // 最低安全庫存
        //                            if (GlobalString.SafeType.storage == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                            {
        //                                // 廠商結餘低於最低安全庫存數值時
        //                                if (intLastStockNumber - intUsedMaterialAfter - intUsedMaterialFront <
        //                                    Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]))
        //                                {
        //                                    string[] arg = new string[1];
        //                                    arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                                    Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInSafe, arg);
        //                                    Warning.SetWarning(GlobalString.WarningType.SubtoalMaterialInSafe, arg);
        //                                }
        //                                // 安全天數
        //                            }
        //                            else if (GlobalString.SafeType.days == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                            {
        //                                // 檢查庫存是否充足
        //                                if (!this.CheckMaterielSafeDays(drMATERIAL_USED["Serial_Number"].ToString(),
        //                                                        Convert.ToInt32(drMATERIAL_USED["Perso_Factory_RID"].ToString()),
        //                                                        Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]),
        //                                                        intLastStockNumber - intUsedMaterialFront - intUsedMaterialAfter))
        //                                {
        //                                    string[] arg = new string[1];
        //                                    arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                                    Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInSafe, arg);
        //                                    Warning.SetWarning(GlobalString.WarningType.SubtoalMaterialInSafe, arg);
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //        #endregion 根據前一天的庫存及今天的庫存。計算物料剩餘數量，判斷是否報警
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        /// <summary>
        /// 取物料的物品、RID等訊息
        /// </summary>
        /// <param name="Serial_Number">品名編號</param>
        /// <returns><DataTable>物料DataTable</returns>
        public DataSet GetMateriel(string Serial_Number)
        {
            DataSet dtsMateriel = null;
            try
            {
                // 取物料的品名
                dirValues.Clear();
                dirValues.Add("serial_number", Serial_Number);

                // 信封
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                // 寄卡單
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_CARD_EXPONENT, dirValues);
                // DM
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_DMTYPE_INFO, dirValues);
                return dtsMateriel;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 檢查物料的庫存是否安全（安全天數）
        /// </summary>
        /// <param name="Materiel_RID"></param>
        /// <param name="Materiel_Type"></param>
        /// <param name="Factory_RID"></param>
        /// <param name="Days"></param>
        /// <returns></returns>
        //public bool CheckMaterielSafeDays(string Serial_Number,
        //                                int Factory_RID,
        //                                int Days,
        //                                int Stock_Number)
        //{
        //    bool blCheckMaterielSafeDays = true;
        //    Days = Days + 1;   // 為了適應匯入時的函數需要，需要多減一天
        //    DateTime dtStartTime = DateTime.Now.AddDays(-Days);
        //    DataTable dtblSubtotal_Import = MaterielUsedCount(Factory_RID,
        //                                        Serial_Number,
        //                                        dtStartTime,
        //                                        DateTime.Now);

        //    int intMaterielWear = 0;
        //    if (null != dtblSubtotal_Import &&
        //        dtblSubtotal_Import.Rows.Count > 0)
        //    {
        //        // 前N天的耗用量
        //        for (int intRow = 0; intRow < dtblSubtotal_Import.Rows.Count; intRow++)
        //        {
        //            intMaterielWear += Convert.ToInt32(dtblSubtotal_Import.Rows[intRow]["System_Num"]);
        //        }

        //        // 如果庫存小於前N天的耗用量
        //        if (Stock_Number < intMaterielWear)
        //        {
        //            blCheckMaterielSafeDays = false;
        //        }
        //    }

        //    return blCheckMaterielSafeDays;
        //}

        /// <summary>
        /// 計算物料庫存消耗檔
        /// </summary>
        /// <param name="Factory_RID">Perso廠商RID</param>
        /// <param name="Serial_Number">物料編號</param>    
        /// <param name="lastSurplusDateTime">最近一次的結餘日期</param>
        /// <param name="thisSurplusDateTime">本次結餘日期</param>
        /// <returns>DataTable<物料使用記錄></returns>
        public DataTable MaterielUsedCount(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            DataTable dtSubtotal_Import = null;
            try
            {
                dirValues.Clear();
                dirValues.Add("Perso_Factory_RID", Factory_RID);
                dirValues.Add("Serial_Number", Serial_Number);
                dirValues.Add("lastSurplusDateTime", lastSurplusDateTime);
                dirValues.Add("thisSurplusDateTime", thisSurplusDateTime);
                DataSet dstSTOCKS_USED = dao.GetList(SEL_MATERIEL_STOCKS_USED, dirValues);
                if (null != dstSTOCKS_USED && dstSTOCKS_USED.Tables.Count > 0 &&
                                dstSTOCKS_USED.Tables[0].Rows.Count > 0)
                {
                    dtSubtotal_Import = dstSTOCKS_USED.Tables[0];
                    dtSubtotal_Import.Columns.Add(new DataColumn("System_Num", Type.GetType("System.Int32")));
                    for (int intRow = 0; intRow < dtSubtotal_Import.Rows.Count; intRow++)
                    {
                        // 取物品的損耗率(關聯到物品表，取物品表的損耗率）
                        Decimal dWear_Rate = GetWearRate(Serial_Number);
                        // 系統耗用量
                        dtSubtotal_Import.Rows[intRow]["System_Num"] = Convert.ToInt32(dtSubtotal_Import.Rows[intRow]["Number"]) * (dWear_Rate / 100 + 1);
                    }
                }
                return dtSubtotal_Import;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 取物品的損耗率
        /// </summary>    
        /// <param name="Serial_Number">物品編號 1：信封；2：寄卡單；3：DM</param>
        /// <returns>Decimal<物品的耗用率></returns>
        public Decimal GetWearRate(string Serial_Number)
        {
            Decimal dWearRate = 0;
            DataSet dstWearRate = null;

            try
            {
                dirValues.Clear();
                dirValues.Add("Serial_Number", Serial_Number);
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())// 信封
                {
                    dstWearRate = dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                }
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())// 卡單
                {
                    dstWearRate = dao.GetList(SEL_CARD_EXPONENT, dirValues);
                }
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())// DM
                {
                    dstWearRate = dao.GetList(SEL_DMTYPE_INFO, dirValues);
                }

                if (null != dstWearRate &&
                        dstWearRate.Tables.Count > 0 &&
                        dstWearRate.Tables[0].Rows.Count > 0)
                {
                    // 取損耗率
                    dWearRate = Convert.ToDecimal(dstWearRate.Tables[0].Rows[0]["Wear_Rate"]);
                }

                return dWearRate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 取最后一次日結日期
        /// </summary>
        /// <returns></returns>
        public DateTime getLastSurplusDate()
        {
            DateTime dtLastSurplusDate = Convert.ToDateTime("1900-01-01");
            try
            {
                DataSet dsLAST_SURPLUS_DAY = dao.GetList(SEL_LAST_SURPLUS_DAY);
                if (dsLAST_SURPLUS_DAY != null
                        && dsLAST_SURPLUS_DAY.Tables.Count > 0
                        && dsLAST_SURPLUS_DAY.Tables[0].Rows.Count > 0)
                {
                    dtLastSurplusDate = Convert.ToDateTime(dsLAST_SURPLUS_DAY.Tables[0].Rows[0]["Stock_Date"].ToString());
                }
                return dtLastSurplusDate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 用小計檔生成卡片對應的物料耗用記錄
        /// </summary>
        /// <returns></returns>
        //public DataTable getMaterialUsed(string strFactory_RID, DateTime importDate)
        //{
        //    DataTable dtUSE_CARDTYPE = new DataTable();
        //    dtUSE_CARDTYPE.Columns.Add("Stock_Date", Type.GetType("System.DateTime"));
        //    dtUSE_CARDTYPE.Columns.Add("Number", Type.GetType("System.Int32"));
        //    dtUSE_CARDTYPE.Columns.Add("Serial_Number", Type.GetType("System.String"));
        //    dtUSE_CARDTYPE.Columns.Add("Perso_Factory_RID", Type.GetType("System.Int32"));

        //    try
        //    {
        //        dirValues.Clear();
        //        dirValues.Add("perso_factory_rid", strFactory_RID);
        //        //取信封和寄卡單耗用記錄，DataSet<物料耗用記錄>
        //        DataSet dsMATERIAL_BY_SUBTOTAL = dao.GetList(SEL_MATERIAL_BY_TEMP_MADE_CARD, dirValues);
        //        foreach (DataRow dr in dsMATERIAL_BY_SUBTOTAL.Tables[0].Rows)
        //        {
        //            if (dr["CE_Number"].ToString() != "")
        //            {
        //                DataRow[] drSelect = dtUSE_CARDTYPE.Select("Serial_Number = '" + dr["CE_Number"].ToString() + "'");
        //                if (drSelect.Length > 0)
        //                {
        //                    drSelect[0]["Number"] = Convert.ToInt32(drSelect[0]["Number"]) + Convert.ToInt32(dr["Number"]);
        //                }
        //                else
        //                {
        //                    DataRow drNewCARD_EXPONENT = dtUSE_CARDTYPE.NewRow();
        //                    drNewCARD_EXPONENT["Stock_Date"] = importDate;
        //                    drNewCARD_EXPONENT["Number"] = Convert.ToInt32(dr["Number"]);
        //                    drNewCARD_EXPONENT["Serial_Number"] = dr["CE_Number"].ToString();
        //                    drNewCARD_EXPONENT["Perso_Factory_RID"] = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                    dtUSE_CARDTYPE.Rows.Add(drNewCARD_EXPONENT);
        //                }
        //            }

        //            if (dr["EI_Number"].ToString() != "")
        //            {
        //                DataRow[] drSelect = dtUSE_CARDTYPE.Select("Serial_Number = '" + dr["EI_Number"].ToString() + "'");
        //                if (drSelect.Length > 0)
        //                {
        //                    drSelect[0]["Number"] = Convert.ToInt32(drSelect[0]["Number"]) + Convert.ToInt32(dr["Number"]);
        //                }
        //                else
        //                {
        //                    DataRow drNewENVELOPE_INFO = dtUSE_CARDTYPE.NewRow();
        //                    drNewENVELOPE_INFO["Stock_Date"] = importDate;
        //                    drNewENVELOPE_INFO["Number"] = Convert.ToInt32(dr["Number"]);
        //                    drNewENVELOPE_INFO["Serial_Number"] = dr["EI_Number"].ToString();
        //                    drNewENVELOPE_INFO["Perso_Factory_RID"] = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                    dtUSE_CARDTYPE.Rows.Add(drNewENVELOPE_INFO);
        //                }
        //            }
        //        }

        //        //取DM耗用記錄，DataSet<DM物料耗用記錄>
        //        DataSet MATERIAL_BY_SUBTOTAL_DM = dao.GetList(SEL_MATERIAL_BY_TEMP_MADE_CARD_DM, dirValues);
        //        foreach (DataRow dr in MATERIAL_BY_SUBTOTAL_DM.Tables[0].Rows)
        //        {
        //            if (dr["DI_Number"].ToString() != "")
        //            {
        //                DataRow[] drSelect = dtUSE_CARDTYPE.Select("Serial_Number = '" + dr["DI_Number"].ToString() + "'");
        //                if (drSelect.Length > 0)
        //                {
        //                    drSelect[0]["Number"] = Convert.ToInt32(drSelect[0]["Number"]) + Convert.ToInt32(dr["Number"]);
        //                }
        //                else
        //                {
        //                    DataRow drNewDMTYPE_INFO = dtUSE_CARDTYPE.NewRow();
        //                    drNewDMTYPE_INFO["Stock_Date"] = importDate;
        //                    drNewDMTYPE_INFO["Number"] = Convert.ToInt32(dr["Number"]);
        //                    drNewDMTYPE_INFO["Serial_Number"] = dr["DI_Number"].ToString();
        //                    drNewDMTYPE_INFO["Perso_Factory_RID"] = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                    dtUSE_CARDTYPE.Rows.Add(drNewDMTYPE_INFO);
        //                }
        //            }
        //        }

        //        return dtUSE_CARDTYPE;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        #endregion 物料消耗報警

        /// <summary>
        ///下載資料檢核
        /// </summary>
        public DataSet ImportCheck(string strPath,  string Date, string filename)
        {
            StreamReader sr = null;
            //DataSet dsCheck_Date = null;
            DataSet dst = new DataSet();
            FACTORY_CHANGE_IMPORT fciModel = new FACTORY_CHANGE_IMPORT();

            //判斷狀況代碼數據集
            DataSet dsFileStatus = null;
            ArrayList FileStatus = new ArrayList();

            //判斷卡種有效期和是否停用數據集
            DataSet dsCARD_TYPE_EndTime = null;

            //異動信息數據集
            DataSet dsFACTORY_CHANGE_IMPORT = null;
            DataTable dtblFileImp = CreatTable();

            //取出所有卡種卡況字段的信息
            dsFileStatus = CheckFileStatus();
            //取出所有有效期和是否停用字段信息
            dsCARD_TYPE_EndTime = CheckCARD_TYPE_EndTime();

            string[] perso=filename.Split('.');
            DataTable dtblFileImp2 = GetPerso(perso[0].Trim());

            //取出所有廠商異動信息表字段信息判斷是否重復匯入
            dsFACTORY_CHANGE_IMPORT = CheckFACTORY_CHANGE_IMPORT(Date, dtblFileImp2.Rows[0]["Factory_ID"].ToString().Trim());


            sr = new StreamReader(strPath, System.Text.Encoding.Default);
            string[] strLine;
            string strReadLine = "";
            int count = 1;
            strErr = "";
            try
            {

                while ((strReadLine = sr.ReadLine()) != null)
                {
                    if (count == 1)
                    {
                        strLine = new string[2];

                        if (strReadLine.Length != 13)
                        {
                            throw new Exception("Error");
                            //strErr += "匯入文件的格式不正確!\n";
                        }
                        strLine[0] = strReadLine.Substring(0, 8);
                        strLine[1] = strReadLine.Substring(8, 5);
                        if (strLine[1].Trim() != dtblFileImp2.Rows[0]["Factory_ID"].ToString().Trim())
                        {
                            throw new Exception("Error");
                            //strErr += "匯入文件的Perso廠商不正確!\n";
                        }
                        for (int i = 0; i < strLine.Length; i++)
                        {
                            int num = i + 1;
                            if (StringUtil.IsEmpty(strLine[i]))
                            {
                                if (count == 1 && num == 1)
                                {
                                    throw new Exception("Error");
                                    //strErr += "匯入文件的匯入日期不能為空！\n";
                                }
                                if (count == 1 && num == 2)
                                {
                                    throw new Exception("Error");
                                    //strErr += "匯入文件的Perso廠商不能為空！\n";
                                }
                            }
                            else
                            {
                                
                                strErr += CheckFileOneColumn(strLine[i], num, count);
                            }

                        }

                        if (!StringUtil.IsEmpty(strErr))
                            throw new Exception("Error");
                        count++;
                    }
                    else
                    {
                        if (!StringUtil.IsEmpty(strReadLine))
                        {
                            if (StringUtil.GetByteLength(strReadLine) != 50)//列數量檢查
                            {
                                throw new Exception("Error");
                                //strErr += "第" + count.ToString() + "行列數不正確。\n";
                            }
                            else
                            {
                                string[] strLine1 = new string[3];
                                Depository010BL bl003 = new Depository010BL();
                                int nextBegin = 0;
                                strLine1[0] = bl003.GetSubstringByByte(strReadLine, nextBegin, 9, out nextBegin).Trim();
                                strLine1[1] = bl003.GetSubstringByByte(strReadLine, nextBegin, 30, out nextBegin).Trim();
                                strLine1[2] = bl003.GetSubstringByByte(strReadLine, nextBegin, 11, out nextBegin).Trim();

                                strLine = new string[6];

                                strLine[0] = strLine1[0].Substring(0, 3);// Type
                                strLine[1] = strLine1[0].Substring(3, 4);// Affinity
                                strLine[2] = strLine1[0].Substring(7, 2);// Photo
                                strLine[3] = strLine1[1];// Name
                                strLine[4] = strLine1[2].Substring(0, 2);// 狀況代碼
                                strLine[5] = strLine1[2].Substring(2, 9);// 數量


                                DataRow drRegex = dtblFileImp.NewRow();//作為數據驗證
                                DataRow dr = dtblFileImp.NewRow();//作為插入數據庫
                                for (int i = 0; i < strLine.Length; i++)
                                {
                                    int num = i + 1;
                                    if (StringUtil.IsEmpty(strLine[i]))
                                        throw new Exception("Error");
                                    //strErr += "第" + (count - 1).ToString() + "行第" + num.ToString() + "列為空;\n";
                                }


                                for (int i = 0; i < strLine.Length; i++)
                                {
                                    int num = i + 1;
                                    if (i == 0)
                                    {
                                        DataTable dtcard = GetCardType();
                                        if (dtcard.Rows.Count == 0)
                                        {
                                            throw new Exception("Error");
                                            //strErr += "第" + (count - 1).ToString() + "行第" + num.ToString() + "行異動信息的卡種信息在系統中非法，請檢查！\n";
                                        }

                                        DataRow[] drs = dtcard.Select("TYPE = '" + strLine[0] + "'" +
                                                            " AND AFFINITY = '" + strLine[1] + "'" +
                                            //Mark By Jacky On 2009/04/17，不檢查版面簡稱是否一致！
                                                            //" AND Name = '" + strLine[3] + "'" +
                                                            " AND PHOTO = '" + strLine[2] + "'");
                                        if (drs.Length == 0)
                                        {
                                            // 存在檢查
                                            throw new Exception("Error");
                                            //strErr += "第" + count.ToString() + "行的卡種不存在;\\n";
                                        }


                                    }
                                    if (i == 3)
                                    {
                                        strLine[3] = strLine[3].Trim();//去掉卡種不滿30位填充的空格
                                    }
                                    if (i == 4)
                                    {
                                        if (dsFileStatus.Tables[0].Rows.Count != 0)
                                        {
                                            for (int j = 0; j < dsFileStatus.Tables[0].Rows.Count; j++)
                                            {
                                                FileStatus.Add(dsFileStatus.Tables[0].Rows[j]["Status_Code"].ToString());
                                            }
                                            //判斷狀況代碼是否存在
                                            if (FileStatus.Contains(strLine[4]) == false)
                                            {
                                                throw new Exception("Error");
                                                //strErr += "第" + (count - 1).ToString() + "行第" + num.ToString() + "列異動信息狀況代碼內容不正確，請檢查！;\n";
                                            }
                                            else
                                            {
                                                DataRow[] drowLevels = dsFileStatus.Tables[0].Select("Status_Code='" + strLine[4] + "'");
                                                strLine[4] = drowLevels[0]["RID"].ToString();
                                            }
                                        }                                       
                                    }

                                    drRegex[i] = strLine[i];
                                    dr[i] = strLine[i];
                                }


                                bool IsDate = false;//false在有效期內
                                bool IsN = false;//false不存在停用
                                bool IsCard_type = false;//true存在非法卡種
                                if (dsCARD_TYPE_EndTime.Tables[0].Rows.Count != 0)
                                {
                                    for (int i = 0; i < dsCARD_TYPE_EndTime.Tables[0].Rows.Count; i++)
                                    {
                                        if (dsCARD_TYPE_EndTime.Tables[0].Rows[i]["TYPE"].ToString() == drRegex.ItemArray[0].ToString()
                                            && dsCARD_TYPE_EndTime.Tables[0].Rows[i]["AFFINITY"].ToString() == drRegex.ItemArray[1].ToString()
                                            && dsCARD_TYPE_EndTime.Tables[0].Rows[i]["PHOTO"].ToString() == drRegex.ItemArray[2].ToString()
                                            && dsCARD_TYPE_EndTime.Tables[0].Rows[i]["Name"].ToString() == drRegex.ItemArray[3].ToString())
                                        {
                                            string time = dsCARD_TYPE_EndTime.Tables[0].Rows[i]["End_Time"].ToString();
                                            DateTime time2 = Convert.ToDateTime(dsCARD_TYPE_EndTime.Tables[0].Rows[i]["End_Time"].ToString());
                                            string time3 = Convert.ToDateTime(dsCARD_TYPE_EndTime.Tables[0].Rows[i]["End_Time"].ToString()).ToString("yyyyMMdd");
                                            DateTime dtime = DateTime.ParseExact(time3, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                                            //DateTime dtime2 = Convert.ToDateTime(Date);

                                            if (DateTime.ParseExact(time3, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo) <= DateTime.ParseExact(Date, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                                                && time3 != "19000101")
                                            {
                                                IsDate = true;
                                                throw new Exception("Error");
                                                //strErr += "第" + count.ToString() + "行異動信息的卡種不在有效期，請檢查！\n";
                                            }
                                            if (dsCARD_TYPE_EndTime.Tables[0].Rows[i]["Is_Using"].ToString() == "N")
                                            {
                                                IsN = true;
                                                throw new Exception("Error");
                                                //strErr += "第" + count.ToString() + "行異動信息的卡種已經停用，請檢查！\n";
                                            }
                                        }


                                    }


                                    IsDate = false;
                                    IsN = false;
                                    IsCard_type = false;
                                }


                                bool IsFACTORY_CHANGE_IMPORT = false;//false不存在重復廠商異動信息
                                if (dsFACTORY_CHANGE_IMPORT.Tables[0].Rows.Count != 0)
                                {
                                    for (int i = 0; i < dsFACTORY_CHANGE_IMPORT.Tables[0].Rows.Count; i++)
                                    {
                                        if (dsFACTORY_CHANGE_IMPORT.Tables[0].Rows[i]["Space_Short_Name"].ToString() == drRegex.ItemArray[3].ToString()
                                            && Convert.ToInt16(dsFACTORY_CHANGE_IMPORT.Tables[0].Rows[i]["Status_Code"].ToString()) == Convert.ToInt16(drRegex.ItemArray[4])
                                            && dsFACTORY_CHANGE_IMPORT.Tables[0].Rows[i]["Perso_Factory_RID"].ToString() == GetFactoryList_RID(dtblFileImp2.Rows[0]["Factory_ID"].ToString().Trim()))
                                        {
                                            IsFACTORY_CHANGE_IMPORT = true;//存在重復廠商異動信息
                                        }
                                    }
                                }
                                if (IsFACTORY_CHANGE_IMPORT)
                                {
                                    throw new Exception("Error");
                                    //strErr += "第" + (count - 1).ToString() + "行異動信息已匯入，請檢查！\n";
                                }
                                if (dtblFileImp.Select("TYPE='" + strLine[0] + "' and AFFINITY='" + strLine[1] + "' and PHOTO='" + strLine[2] + "' and Status_RID='" + strLine[4] + "'").Length > 0)
                                {
                                    //sbErr.Append("第" + count.ToString() + "行的廠商庫存異動資訊不能重複匯入;\\n");
                                    throw new Exception("Error");
                                }
                                IsFACTORY_CHANGE_IMPORT = false;
                                dr[strLine.Length] = dtblFileImp2.Rows[0]["RID"].ToString();
                                dtblFileImp.Rows.Add(dr);

                            }
                        }
                        count++;
                    }
                }

                if (!StringUtil.IsEmpty(strErr))
                {
                    // 格式不正確，警示
                    object[] arg = new object[2];
                    arg[0] = dtblFileImp2.Rows[0]["Factory_ShortName_EN"].ToString();
                    arg[1] = strErr;
                    Warning.SetWarning(GlobalString.WarningType.FactoryStocksChange, arg);
                }

                dst.Tables.Add(dtblFileImp);


            }
            catch
            {
                object[] arg = new object[2];
                arg[0] = dtblFileImp2.Rows[0]["Factory_ShortName_EN"].ToString();
                arg[1] = strErr;
                Warning.SetWarning(GlobalString.WarningType.FactoryStocksChange, arg);
            }
            finally
            {
                sr.Close();
            }
            return dst;
        }
        public string ImportCardTypeChange(DataSet dst, string Date)
        {
            string strReturn = "";
            try
            {
                dao.OpenConnection();
                FACTORY_CHANGE_IMPORT fciModel = new FACTORY_CHANGE_IMPORT();

                DataTable dtblFileImp = dst.Tables[0];

                foreach (DataRow drowFileImp in dtblFileImp.Rows)
                {
                    fciModel.TYPE = drowFileImp["TYPE"].ToString();
                    fciModel.AFFINITY = drowFileImp["AFFINITY"].ToString();
                    fciModel.PHOTO = drowFileImp["PHOTO"].ToString();
                    fciModel.Space_Short_Name = drowFileImp["Name"].ToString();
                    fciModel.Status_RID = Convert.ToInt32(drowFileImp["Status_RID"]);
                    fciModel.Number = Convert.ToInt32(drowFileImp["Number"].ToString().Replace(",", ""));
                    fciModel.Date_Time = DateTime.ParseExact(Date, "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                    fciModel.Perso_Factory_RID = Convert.ToInt32(drowFileImp["Factory_RID"].ToString());
                    fciModel.Is_Auto_Import = "Y";
                    dao.Add<FACTORY_CHANGE_IMPORT>(fciModel, "RID");
                }

                //事務提交
                dao.Commit();

                this.SendWarningPerso(dtblFileImp, fciModel.Perso_Factory_RID.ToString());
                // 物料消耗報警
                //this.Material_Used_Warnning(fciModel.Perso_Factory_RID.ToString(),DateTime.Now);
                //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 ADD BY 楊昆 2009/08/31 start
                InOut000BL BL000 = new InOut000BL();
                BL000.Material_Used_Warnning(fciModel.Perso_Factory_RID.ToString(), DateTime.Now, "2");
                //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 ADD BY 楊昆 2009/08/31 end
            }
            catch (Exception ex)
            {
                strReturn = "Error";
                dao.Rollback();
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("ImportCardTypeChange報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                BatchBL Bbl = new BatchBL();
                Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
                //throw ex;
            }
            finally
            {
                dao.CloseConnection();
            }
            return strReturn;

        }

        /// <summary>
        /// 根據Perso廠商的記錄計算是否有不足的卡種！
        /// </summary>
        /// <param name="dtImport"></param>
        private void SendWarningPerso(DataTable dtImport, string sFactoryRid)
        {
            try
            {
                DataTable dtCardType = this.GetCardType();
                DataTable dtFactory = this.GetFactoryList().Tables[0];
                int iNum = 0;

                DataTable dtCard = new DataTable();
                dtCard.Columns.Add("card");
                dtCard.Columns.Add("factory");

                DataTable dtblXuNi = dao.GetList("select CardType_RID from dbo.GROUP_CARD_TYPE a inner join CARD_GROUP b on a.Group_rid=b.rid where b.Group_Name = '虛擬卡'").Tables[0];


                foreach (DataRow dr in dtImport.Rows)
                {
                    DataRow[] drCardType = dtCardType.Select("TYPE='" + dr["TYPE"].ToString() + "' and AFFINITY='"
                        + dr["AFFINITY"].ToString() + "' and PHOTO='" + dr["PHOTO"].ToString() + "'");


                    if (drCardType.Length < 0)
                        continue;

                    int CardRID = 0;

                    if (dr["Status_RID"].ToString() == "4")
                    {
                        if (drCardType[0]["Change_Space_RID"].ToString() != "0")
                            CardRID = int.Parse(drCardType[0]["Change_Space_RID"].ToString());
                        else
                        {
                            if (drCardType[0]["Replace_Space_RID"].ToString() != "0")
                                CardRID = int.Parse(drCardType[0]["Replace_Space_RID"].ToString());
                            else
                                CardRID = int.Parse(drCardType[0]["RID"].ToString());
                        }
                    }
                    else if (dr["Status_RID"].ToString().ToUpper() == "1" || dr["Status_RID"].ToString().ToUpper() == "2" || dr["Status_RID"].ToString().ToUpper() == "3")
                    {
                        if (drCardType[0]["Replace_Space_RID"].ToString() != "0")
                            CardRID = int.Parse(drCardType[0]["Replace_Space_RID"].ToString());
                        else
                            CardRID = int.Parse(drCardType[0]["RID"].ToString());
                    }
                    else
                    {
                        CardRID = int.Parse(drCardType[0]["RID"].ToString());
                    }


                    DataRow[] drFactory = dtFactory.Select("RID='" + sFactoryRid + "'");

                    if (dtCard.Select("card='" + CardRID.ToString() + "' and factory='" + sFactoryRid.ToString() + "'").Length > 0)
                        continue;

                    if (dtblXuNi.Rows.Count > 0)
                    {
                        if (dtblXuNi.Select("CardType_RID = '" + CardRID.ToString() + "'").Length > 0)
                            continue;
                    }

                    DataRow drcard = dtCard.NewRow();
                    drcard[0] = CardRID.ToString();
                    drcard[1] = sFactoryRid.ToString();
                    dtCard.Rows.Add(drcard);

                    CardTypeManager ctm = new CardTypeManager();
                    iNum = ctm.getCurrentStockPerso(Convert.ToInt32(sFactoryRid), CardRID, DateTime.Now.Date.AddDays (1).AddSeconds (-1));

                    //如果庫存小於零，則發送警訊！
                    if (iNum < 0)
                    {
                        object[] arg = new object[2];
                        arg[0] = drFactory[0]["Factory_Shortname_CN"];

                        DataRow[] drCardType1 = dtCardType.Select("RID=" + CardRID.ToString());

                        if (drCardType1.Length > 0)
                        {
                            arg[1] = drCardType1[0]["NAME"];
                        }
                        else
                        {
                            arg[1] = "";
                        }
                        Warning.SetWarning(GlobalString.WarningType.PersoChangeCardInMiss, arg);
                    }

                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("根據Perso廠商的記錄計算是否有不足的卡種！, SendWarningPerso報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
        }


        /// <summary>
        ///初始化DT
        /// </summary>
        public DataTable CreatTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("TYPE", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("AFFINITY", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("PHOTO", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Name", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Status_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Number", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Factory_RID", System.Type.GetType("System.String")));
            return dt;
        }
        /// <summary>
        /// 驗證匯入字段是否滿足格式
        /// </summary>
        /// <param name="strColumn"></param>
        /// <param name="num"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private string CheckFileColumn(string strColumn, int num, int count)
        {
            string strErr = "";
            string Pattern = "";
            MatchCollection Matches;
            switch (num)
            {
                case 1:
                    Pattern = @"^\d{1}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + (count - 1).ToString() + "行第" + num.ToString() + "列格式必須為1位數字;\n";
                    }
                    break;
                case 2:
                    Pattern = @"^\d{9}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + (count - 1).ToString() + "行第" + num.ToString() + "列格式必須為9位數字;\n";
                    }
                    break;
                case 3:
                    Pattern = @"^\d{1,5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + (count - 1).ToString() + "行第" + num.ToString() + "列格式必須為5位以內的數字;\n";
                    }
                    break;
                default:
                    break;
            }

            return strErr;
        }
        /// <summary>
        /// 檢查狀況代碼字段對應的參數在[系統參數檔]中是否存在，并正常使用
        /// </summary>
        private DataSet CheckFileStatus()
        {
            DataSet dsCARDTYPE_STATUS = null;

            dirValues.Clear();
            try
            {
                dsCARDTYPE_STATUS = dao.GetList(SEL_CARDTYPE_STATUS);
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查狀況代碼字段對應的參數在[系統參數檔]中是否存在，并正常使用, CheckFileStatus報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dsCARDTYPE_STATUS;
        }
        /// <summary>
        ///查詢當前異動信息的時間是否在有效期內并且使用狀態是否是停用，如果無查詢結果，則拋出異常，不繼續執行
        /// </summary>
        private DataSet CheckCARD_TYPE_EndTime()
        {
            DataSet dsCARD_TYPE = null;
            dirValues.Clear();
            try
            {
                dsCARD_TYPE = dao.GetList(SEL_CARDTYPE_End_Time);
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("查詢當前異動信息的時間是否在有效期內并且使用狀態是否是停用, CheckCARD_TYPE_EndTime報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dsCARD_TYPE;
        }
        /// <summary>
        ///檢查當前異動信息是否已匯入過（當版面簡稱+狀況代碼+匯入日期+Perso廠代號重復，則表示當前異動信息已匯入過）
        /// </summary>
        ///<returns></returns>
        private DataSet CheckFACTORY_CHANGE_IMPORT(string Import_Date , string FactoryRID)
        {
            DataSet dsFACTORY_CHANGE_IMPORT = null;
            try
            {
                dirValues.Clear();
                dirValues.Add("FactoryRID", FactoryRID);
                dirValues.Add("Import_Date_Start", Import_Date + " 00:00:00");
                dirValues.Add("Import_Date_End", Import_Date + " 23:59:59");
                dsFACTORY_CHANGE_IMPORT =dao.GetList(SEL_FACTORY_CHANGE_IMPORT_ALL,dirValues);
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查當前異動信息是否已匯入過, CheckFACTORY_CHANGE_IMPORT報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dsFACTORY_CHANGE_IMPORT;
        }
        /// <summary>
        /// 驗證匯入字段第一行是否滿足格式
        /// </summary>
        /// <param name="strColumn"></param>
        /// <param name="num"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private string CheckFileOneColumn(string strColumn, int num, int count)
        {
            string strErr = "";
            string Pattern = "";
            MatchCollection Matches;
            switch (num)
            {
                case 1:
                    Pattern = @"^\d{8}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {

                        strErr = "匯入文件的匯入日期必須為8位數字\n";
                    }
                    else
                    {
                        try
                        {
                            string date = strColumn.Substring(0, 4) + "/" + strColumn.Substring(4, 2) + "/" + strColumn.Substring(6, 2);

                            Convert.ToDateTime(date);
                        }
                        catch (Exception ex)
                        {
                            strErr = "匯入文件的匯入日期格式不正確\n";

                        }
                    }
                    break;
                case 2:
                    Pattern = @"^\d{5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {

                        strErr = "匯入文件的Perso廠商必須為5位數字\n";
                    }
                    break;
                default:
                    break;
            }

            return strErr;
        }
        /// <summary>
        /// 獲得Perso廠商RID
        /// </summary>
        /// <returns>DataSet[Perso廠商]</returns>
        public string GetFactoryList_RID(string Factory_ID)
        {
            DataSet dstFactory_RID = null;
            string RID = "";
            try
            {
                dirValues.Clear();
                dirValues.Add("factory_id", Factory_ID);
                dstFactory_RID = dao.GetList(SEL_FACTORY_RID, dirValues);
                if (dstFactory_RID.Tables[0].Rows.Count != 0)
                {
                    RID = dstFactory_RID.Tables[0].Rows[0]["RID"].ToString();
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("獲得Perso廠商RID, GetFactoryList_RID報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return RID;
        }
        /// <summary>
        /// 檢查FTP檔案規則
        /// </summary>
        /// <param name="FileName">檔案名稱(全名)</param>
        /// <returns></returns>
        private bool CheckFile(string FileName)
        {
            string[] fileSplit = CheckFileNameFormat(FileName);
            if (fileSplit == null)
            {
                return false;
            }
            if (fileSplit[0].ToString().ToUpper() != "CARD")
            {
                return false;
            }

            string EnName = fileSplit[2].Substring(0, fileSplit[2].Length - 4);
            if (!CheckEnPersoExist(EnName))
            {
                return false;
            }
            
            string ImportEDate = fileSplit[1].ToString();

            BatchBL bl = new BatchBL();
            if (!(bl.CheckWorkDate(Convert.ToDateTime(ImportEDate.Substring(0, 4) + "/" + ImportEDate.Substring(4, 2) + "/" + ImportEDate.Substring(6, 2))))) //非工作日直接返回，不執行批次
            {
                return false;
            }

            if (CheckImportDate(ImportEDate))
            {
                return false;
            }
            if (CheckImportFile(FileName))
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 檢查檔案名稱
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns>檔案名稱正確返回字符串數組</returns>
        private string[] CheckFileNameFormat(string FileName)
        {
            string[] fileSplit = FileName.Split('-');
            if (fileSplit == null || fileSplit.Length != 3)
                return null;
            else
            {
                return fileSplit;
            }
        }
        /// <summary>
        /// 檢查Perso英文簡稱
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>英文簡稱存在返回true</returns>
        private bool CheckEnPersoExist(string EnPerso)
        {

            try
            {
                dtPeso = new DataTable();
                dirValues.Clear();
                dirValues.Add("Factory_ShortName_EN", EnPerso);
                dtPeso = dao.GetList(SEL_CARDTYPE_PERSO_FACTORY, dirValues).Tables[0];
                if (dtPeso.Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查Perso英文簡稱, CheckEnPersoExist報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }
        /// <summary>
        /// 檢查日期是否已日結
        /// </summary>
        /// <param name="ImportSDate">區間日期起始</param>
        /// <param name="ImportEDate">區間日期結束</param>
        /// <returns>已日結過返回true</returns>
        private bool CheckImportDate(string ImportEDate)
        {

            DataSet dst = new DataSet();
            dirValues.Clear();
            ImportEDate = ImportEDate.Substring(0, 4) + "/" + ImportEDate.Substring(4, 2) + "/" + ImportEDate.Substring(6, 2);
            dirValues.Add("Stock_Date", ImportEDate);
            dst = dao.GetList(CON_CARDTYPE_SURPLUS, dirValues);

            if (dst.Tables[0].Rows[0][0].ToString() == "0")
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// 檢查是否已匯入過
        /// </summary>
        /// <param name="ImportFileName"></param>
        /// <returns>已匯入過返回true</returns>
        private bool CheckImportFile(string ImportFileName)
        {

            string[] sr = ImportFileName.Split('-');
            DataSet dst = new DataSet();
            dirValues.Clear();
            dirValues.Add("Factory_ID", dtPeso.Rows[0]["Factory_ID"].ToString());
            dirValues.Add("Date_Time", sr[1].ToString());
            dst = dao.GetList(CON_IMPORT_CARDTYPE_CHANGE_CHECK, dirValues);
            if (dst.Tables[0].Rows.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        /// <summary>
        /// 檢查是否有紀錄數返回
        /// </summary>
        /// <param name="commondSring">SQL指令</param>
        /// <param name="dirValues">參數</param>
        /// <returns></returns>
        private bool CheckValue(string commondSring, Dictionary<string, object> dirValues)
        {
            try
            {
                object returnObj = dao.ExecuteScalar(commondSring, dirValues);
                if (Convert.ToInt32(returnObj) == 0)
                    return false;
                else
                    return true;
            }
            catch (Exception ex)
            {
                LogFactory.Write("檢查是否有紀錄數返回, CheckValue錯誤：" + commondSring, GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        /// <summary>
        /// 取得Perso廠商信息
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>dt</returns>
        private DataTable GetPerso(string EnPerso)
        {

            try
            {
                dtPeso = new DataTable();
                dirValues.Clear();
                dirValues.Add("Factory_ShortName_EN", EnPerso);
                dtPeso = dao.GetList(SEL_CARDTYPE_PERSO_FACTORY, dirValues).Tables[0];
                if (dtPeso.Rows.Count > 0)
                {
                    return dtPeso;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取得Perso廠商信息, GetPerso報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }

        }

        /// <summary>
        /// 獲得Perso廠商
        /// </summary>
        /// <returns>DataSet[Perso廠商]</returns>
        public DataSet GetFactoryList()
        {
            DataSet dstFactory = null;
            try
            {
                dirValues.Clear();
                dstFactory = dao.GetList(SEL_FACTORY, dirValues);

                return dstFactory;
            }
            catch (Exception ex)
            {
                //ExceptionFactory.CreateCustomSaveException(BizMessage.BizCommMsg.ALT_CMN_InitPageFail, ex.Message, dao.LastCommands);
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("獲得Perso廠商, GetFactoryList報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
        }

        /// <summary>
        /// 取得Perso廠商信息
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>dt</returns>
        private DataTable GetCardType()
        {

            try
            {
                DataTable dt = new DataTable();

                dt = dao.GetList(SEL_CARD_TYPE).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取得Perso廠商信息, GetCardType報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }

        }

    }
}
