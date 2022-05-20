//*****************************************
//*  �@    �̡GGaoAi
//*  �\�໡���G�t�Ӯw�s���ʶפJ
//*  �Ыؤ���G2008-11-19
//*  �ק����G
//*  �ק�O���G
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
        //��ܩҦ����t�Ӹ�ơI
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
        // ���Ʈ��ӳ�ĵ
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

        #region �c�y���
        public DataTable dtPeso;
        public string strErr;
        //public ArrayList erro;
        //�Ѽ�

        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        public InOut002BL()
        {
            //
            // TODO: �b���[�J�غc�禡���{���X
            //
            string ftpPath = ConfigurationManager.AppSettings["FTPCardModify"];
            string localPath = ConfigurationManager.AppSettings["FTPCardModifyPath"];
        }
        #endregion

        #region �U���ɮ�
        public ArrayList DownLoadModify(string FactoryPath)
        {
            ArrayList FileNameList = new ArrayList();
            try  // �[try catch add by judy 2018/03/28
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
                        if (!CheckFile(FileName)) //�ˬd�n�U�����ɮ׳W�h�A�������h���L
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
                                // Legend 2017/11/28 �N���B�R�ɪ`����UAT����, �W�u�O�A�Ѷ} todo
                                returnFlag = ftp.Delete(ftpPath, FileName);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogFactory.Write("�פJ�t�Ӳ����ɤU��FTP�ɮ�DownLoadModify��k����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
           
            return FileNameList;

        }
        #endregion

        #region ���Ʈ��ӳ�ĵ
        /// <summary>
        /// �ھڶפJ���p�p�ɡA�ͦ����Ʈ��ӰO���A�}�P�_���Ʈw�s�O�_�b�w������A
        /// �p�G���b�w������A��ĵ�C
        /// </summary>
        /// <param name="strFactory_RID"></param>
        /// <param name="importDate"></param>
        //public void Material_Used_Warnning(string strFactory_RID,
        //    DateTime importDate)
        //{
        //    try
        //    {
        //        // ���̫�鵲����C
        //        DateTime TheLastestSurplusDate = getLastSurplusDate();

        //        #region �p��q�̫�@�Ӥ鵲������U�@�Ѩ��ƶפJ������d���s���ơA�O�s���{�ɪ�(TEMP_MADE_CARD)
        //        dirValues.Clear();
        //        dirValues.Add("Perso_Factory_RID", strFactory_RID);
        //        dirValues.Add("From_Date_Time", TheLastestSurplusDate.ToString("yyyy/MM/dd 23:59:59"));
        //        dirValues.Add("End_Date_Time", importDate.ToString("yyyy/MM/dd 23:59:59"));
        //        DataSet dsMade_Card = dao.GetList(SEL_MADE_CARD_WARNNING, dirValues);
        //        DataSet dsEXPRESSIONS_DEFINE = dao.GetList(SEL_EXPRESSIONS_DEFINE_WARNNING);

        //        //�d�خ��Ӫ�
        //        DataTable dtUSE_CARDTYPE = new DataTable();
        //        dtUSE_CARDTYPE.Columns.Add("Perso_Factory_RID");
        //        dtUSE_CARDTYPE.Columns.Add("CardType_RID");
        //        dtUSE_CARDTYPE.Columns.Add("Number");

        //        //��Perso�t�B�d�ت��p����Ӷq�]�`���[�`�U�ت��p�����Ӽƶq�^
        //        int Card_Type_Rid = 0;
        //        int Perso_Factory_RID = 0;
        //        int Number = 0;
        //        //todo ���`���i�H��i���s�x�L�{
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

        //                #region �����ӥd����,�p����ӥd��
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
        //                #region �����ӥd����,�p����ӥd��
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

        //        // �R���{�ɪ����ƾ�
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

        //        #endregion �p���ѻs���d��

        //        // �ھڻs���d�ơA�p�⪫�Ʈ���
        //        //200908IR�ӥζq=�p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/11 start
        //        InOut001BL BL001 = new InOut001BL();
        //        DataTable dtMATERIAL_USED = BL001.getMaterialUsed(strFactory_RID, importDate);
        //        //200908IR�ӥζq=�p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/11 end

        //        // �p�⪫�ƳѧE�ƶq�}ĵ��
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
        /// �p�⪫�ƳѧE�ƶq�}ĵ��
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

        //        #region �ھګe�@�Ѫ��w�s�Τ��Ѫ��w�s�C�p�⪫�ƳѾl�ƶq�A�P�_�O�_��ĵ
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
        //                // �q�L����鵲��A�ӥ�
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
        //                    // �L��ɪ��w�s
        //                    int intLastStockNumber = Convert.ToInt32(dsMaterielStocksManager.Tables[0].Rows[0]["Number"].ToString());
        //                    // �q�L����̵��l�骺����
        //                    int intUsedMaterialFront = 0;
        //                    if (dsUsedMaterial.Tables[0].Rows[0]["Number"] != DBNull.Value)
        //                        intUsedMaterialFront = Convert.ToInt32(dsUsedMaterial.Tables[0].Rows[0]["Number"]);

        //                    // �̫ᵲ�l��Z������
        //                    int intUsedMaterialAfter = Convert.ToInt32(drMATERIAL_USED["Number"]);

        //                    // �w�s��0�ɡA��ܮw�s����
        //                    if (intLastStockNumber <= 0)
        //                    {
        //                        if (bl010.DmNotSafe_Type(drMATERIAL_USED["Serial_Number"].ToString()))
        //                        {
        //                            // �w�s����
        //                            string[] arg = new string[1];
        //                            arg[0] = dsMaterielStocksManager.Tables[0].Rows[0]["Name"].ToString();
        //                            Warning.SetWarning(GlobalString.WarningType.SubtotalMaterialInMiss, arg);
        //                        }
        //                    }
        //                    // �p�G�e�@�Ѫ��w�s�p�E���Ѫ�����
        //                    else if (intLastStockNumber < (intUsedMaterialFront + intUsedMaterialAfter))
        //                    {
        //                        if (bl010.DmNotSafe_Type(drMATERIAL_USED["Serial_Number"].ToString()))
        //                        {
        //                            // �w�s����
        //                            string[] arg = new string[1];
        //                            arg[0] = dsMaterielStocksManager.Tables[0].Rows[0]["Name"].ToString();
        //                            Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInMiss, arg);
        //                            Warning.SetWarning(GlobalString.WarningType.SubtotalMaterialInMiss, arg);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        // �����ƪ��w���w�s�T��
        //                        DataSet dtMateriel = this.GetMateriel(drMATERIAL_USED["Serial_Number"].ToString());
        //                        if (null != dtMateriel &&
        //                            dtMateriel.Tables.Count > 0 &&
        //                            dtMateriel.Tables[0].Rows.Count > 0)
        //                        {
        //                            // �̧C�w���w�s
        //                            if (GlobalString.SafeType.storage == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                            {
        //                                // �t�ӵ��l�C��̧C�w���w�s�ƭȮ�
        //                                if (intLastStockNumber - intUsedMaterialAfter - intUsedMaterialFront <
        //                                    Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]))
        //                                {
        //                                    string[] arg = new string[1];
        //                                    arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                                    Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInSafe, arg);
        //                                    Warning.SetWarning(GlobalString.WarningType.SubtoalMaterialInSafe, arg);
        //                                }
        //                                // �w���Ѽ�
        //                            }
        //                            else if (GlobalString.SafeType.days == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                            {
        //                                // �ˬd�w�s�O�_�R��
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

        //        #endregion �ھګe�@�Ѫ��w�s�Τ��Ѫ��w�s�C�p�⪫�ƳѾl�ƶq�A�P�_�O�_��ĵ
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        /// <summary>
        /// �����ƪ����~�BRID���T��
        /// </summary>
        /// <param name="Serial_Number">�~�W�s��</param>
        /// <returns><DataTable>����DataTable</returns>
        public DataSet GetMateriel(string Serial_Number)
        {
            DataSet dtsMateriel = null;
            try
            {
                // �����ƪ��~�W
                dirValues.Clear();
                dirValues.Add("serial_number", Serial_Number);

                // �H��
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                // �H�d��
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
        /// �ˬd���ƪ��w�s�O�_�w���]�w���Ѽơ^
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
        //    Days = Days + 1;   // ���F�A���פJ�ɪ���ƻݭn�A�ݭn�h��@��
        //    DateTime dtStartTime = DateTime.Now.AddDays(-Days);
        //    DataTable dtblSubtotal_Import = MaterielUsedCount(Factory_RID,
        //                                        Serial_Number,
        //                                        dtStartTime,
        //                                        DateTime.Now);

        //    int intMaterielWear = 0;
        //    if (null != dtblSubtotal_Import &&
        //        dtblSubtotal_Import.Rows.Count > 0)
        //    {
        //        // �eN�Ѫ��ӥζq
        //        for (int intRow = 0; intRow < dtblSubtotal_Import.Rows.Count; intRow++)
        //        {
        //            intMaterielWear += Convert.ToInt32(dtblSubtotal_Import.Rows[intRow]["System_Num"]);
        //        }

        //        // �p�G�w�s�p��eN�Ѫ��ӥζq
        //        if (Stock_Number < intMaterielWear)
        //        {
        //            blCheckMaterielSafeDays = false;
        //        }
        //    }

        //    return blCheckMaterielSafeDays;
        //}

        /// <summary>
        /// �p�⪫�Ʈw�s������
        /// </summary>
        /// <param name="Factory_RID">Perso�t��RID</param>
        /// <param name="Serial_Number">���ƽs��</param>    
        /// <param name="lastSurplusDateTime">�̪�@�������l���</param>
        /// <param name="thisSurplusDateTime">�������l���</param>
        /// <returns>DataTable<���ƨϥΰO��></returns>
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
                        // �����~���l�Ӳv(���p�쪫�~��A�����~���l�Ӳv�^
                        Decimal dWear_Rate = GetWearRate(Serial_Number);
                        // �t�ίӥζq
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
        /// �����~���l�Ӳv
        /// </summary>    
        /// <param name="Serial_Number">���~�s�� 1�G�H�ʡF2�G�H�d��F3�GDM</param>
        /// <returns>Decimal<���~���ӥβv></returns>
        public Decimal GetWearRate(string Serial_Number)
        {
            Decimal dWearRate = 0;
            DataSet dstWearRate = null;

            try
            {
                dirValues.Clear();
                dirValues.Add("Serial_Number", Serial_Number);
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())// �H��
                {
                    dstWearRate = dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                }
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())// �d��
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
                    // ���l�Ӳv
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
        /// ���̦Z�@���鵲���
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
        /// �Τp�p�ɥͦ��d�����������ƯӥΰO��
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
        //        //���H�ʩM�H�d��ӥΰO���ADataSet<���ƯӥΰO��>
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

        //        //��DM�ӥΰO���ADataSet<DM���ƯӥΰO��>
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

        #endregion ���Ʈ��ӳ�ĵ

        /// <summary>
        ///�U������ˮ�
        /// </summary>
        public DataSet ImportCheck(string strPath,  string Date, string filename)
        {
            StreamReader sr = null;
            //DataSet dsCheck_Date = null;
            DataSet dst = new DataSet();
            FACTORY_CHANGE_IMPORT fciModel = new FACTORY_CHANGE_IMPORT();

            //�P�_���p�N�X�ƾڶ�
            DataSet dsFileStatus = null;
            ArrayList FileStatus = new ArrayList();

            //�P�_�d�ئ��Ĵ��M�O�_���μƾڶ�
            DataSet dsCARD_TYPE_EndTime = null;

            //���ʫH���ƾڶ�
            DataSet dsFACTORY_CHANGE_IMPORT = null;
            DataTable dtblFileImp = CreatTable();

            //���X�Ҧ��d�إd�p�r�q���H��
            dsFileStatus = CheckFileStatus();
            //���X�Ҧ����Ĵ��M�O�_���Φr�q�H��
            dsCARD_TYPE_EndTime = CheckCARD_TYPE_EndTime();

            string[] perso=filename.Split('.');
            DataTable dtblFileImp2 = GetPerso(perso[0].Trim());

            //���X�Ҧ��t�Ӳ��ʫH����r�q�H���P�_�O�_���_�פJ
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
                            //strErr += "�פJ��󪺮榡�����T!\n";
                        }
                        strLine[0] = strReadLine.Substring(0, 8);
                        strLine[1] = strReadLine.Substring(8, 5);
                        if (strLine[1].Trim() != dtblFileImp2.Rows[0]["Factory_ID"].ToString().Trim())
                        {
                            throw new Exception("Error");
                            //strErr += "�פJ���Perso�t�Ӥ����T!\n";
                        }
                        for (int i = 0; i < strLine.Length; i++)
                        {
                            int num = i + 1;
                            if (StringUtil.IsEmpty(strLine[i]))
                            {
                                if (count == 1 && num == 1)
                                {
                                    throw new Exception("Error");
                                    //strErr += "�פJ��󪺶פJ������ର�šI\n";
                                }
                                if (count == 1 && num == 2)
                                {
                                    throw new Exception("Error");
                                    //strErr += "�פJ���Perso�t�Ӥ��ର�šI\n";
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
                            if (StringUtil.GetByteLength(strReadLine) != 50)//�C�ƶq�ˬd
                            {
                                throw new Exception("Error");
                                //strErr += "��" + count.ToString() + "��C�Ƥ����T�C\n";
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
                                strLine[4] = strLine1[2].Substring(0, 2);// ���p�N�X
                                strLine[5] = strLine1[2].Substring(2, 9);// �ƶq


                                DataRow drRegex = dtblFileImp.NewRow();//�@���ƾ�����
                                DataRow dr = dtblFileImp.NewRow();//�@�����J�ƾڮw
                                for (int i = 0; i < strLine.Length; i++)
                                {
                                    int num = i + 1;
                                    if (StringUtil.IsEmpty(strLine[i]))
                                        throw new Exception("Error");
                                    //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�C����;\n";
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
                                            //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�沧�ʫH�����d�ثH���b�t�Τ��D�k�A���ˬd�I\n";
                                        }

                                        DataRow[] drs = dtcard.Select("TYPE = '" + strLine[0] + "'" +
                                                            " AND AFFINITY = '" + strLine[1] + "'" +
                                            //Mark By Jacky On 2009/04/17�A���ˬd����²�٬O�_�@�P�I
                                                            //" AND Name = '" + strLine[3] + "'" +
                                                            " AND PHOTO = '" + strLine[2] + "'");
                                        if (drs.Length == 0)
                                        {
                                            // �s�b�ˬd
                                            throw new Exception("Error");
                                            //strErr += "��" + count.ToString() + "�檺�d�ؤ��s�b;\\n";
                                        }


                                    }
                                    if (i == 3)
                                    {
                                        strLine[3] = strLine[3].Trim();//�h���d�ؤ���30���R���Ů�
                                    }
                                    if (i == 4)
                                    {
                                        if (dsFileStatus.Tables[0].Rows.Count != 0)
                                        {
                                            for (int j = 0; j < dsFileStatus.Tables[0].Rows.Count; j++)
                                            {
                                                FileStatus.Add(dsFileStatus.Tables[0].Rows[j]["Status_Code"].ToString());
                                            }
                                            //�P�_���p�N�X�O�_�s�b
                                            if (FileStatus.Contains(strLine[4]) == false)
                                            {
                                                throw new Exception("Error");
                                                //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�C���ʫH�����p�N�X���e�����T�A���ˬd�I;\n";
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


                                bool IsDate = false;//false�b���Ĵ���
                                bool IsN = false;//false���s�b����
                                bool IsCard_type = false;//true�s�b�D�k�d��
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
                                                //strErr += "��" + count.ToString() + "�沧�ʫH�����d�ؤ��b���Ĵ��A���ˬd�I\n";
                                            }
                                            if (dsCARD_TYPE_EndTime.Tables[0].Rows[i]["Is_Using"].ToString() == "N")
                                            {
                                                IsN = true;
                                                throw new Exception("Error");
                                                //strErr += "��" + count.ToString() + "�沧�ʫH�����d�ؤw�g���ΡA���ˬd�I\n";
                                            }
                                        }


                                    }


                                    IsDate = false;
                                    IsN = false;
                                    IsCard_type = false;
                                }


                                bool IsFACTORY_CHANGE_IMPORT = false;//false���s�b���_�t�Ӳ��ʫH��
                                if (dsFACTORY_CHANGE_IMPORT.Tables[0].Rows.Count != 0)
                                {
                                    for (int i = 0; i < dsFACTORY_CHANGE_IMPORT.Tables[0].Rows.Count; i++)
                                    {
                                        if (dsFACTORY_CHANGE_IMPORT.Tables[0].Rows[i]["Space_Short_Name"].ToString() == drRegex.ItemArray[3].ToString()
                                            && Convert.ToInt16(dsFACTORY_CHANGE_IMPORT.Tables[0].Rows[i]["Status_Code"].ToString()) == Convert.ToInt16(drRegex.ItemArray[4])
                                            && dsFACTORY_CHANGE_IMPORT.Tables[0].Rows[i]["Perso_Factory_RID"].ToString() == GetFactoryList_RID(dtblFileImp2.Rows[0]["Factory_ID"].ToString().Trim()))
                                        {
                                            IsFACTORY_CHANGE_IMPORT = true;//�s�b���_�t�Ӳ��ʫH��
                                        }
                                    }
                                }
                                if (IsFACTORY_CHANGE_IMPORT)
                                {
                                    throw new Exception("Error");
                                    //strErr += "��" + (count - 1).ToString() + "�沧�ʫH���w�פJ�A���ˬd�I\n";
                                }
                                if (dtblFileImp.Select("TYPE='" + strLine[0] + "' and AFFINITY='" + strLine[1] + "' and PHOTO='" + strLine[2] + "' and Status_RID='" + strLine[4] + "'").Length > 0)
                                {
                                    //sbErr.Append("��" + count.ToString() + "�檺�t�Ӯw�s���ʸ�T���୫�ƶפJ;\\n");
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
                    // �榡�����T�Aĵ��
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

                //�ưȴ���
                dao.Commit();

                this.SendWarningPerso(dtblFileImp, fciModel.Perso_Factory_RID.ToString());
                // ���Ʈ��ӳ�ĵ
                //this.Material_Used_Warnning(fciModel.Perso_Factory_RID.ToString(),DateTime.Now);
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� ADD BY ���� 2009/08/31 start
                InOut000BL BL000 = new InOut000BL();
                BL000.Material_Used_Warnning(fciModel.Perso_Factory_RID.ToString(), DateTime.Now, "2");
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� ADD BY ���� 2009/08/31 end
            }
            catch (Exception ex)
            {
                strReturn = "Error";
                dao.Rollback();
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("ImportCardTypeChange����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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
        /// �ھ�Perso�t�Ӫ��O���p��O�_���������d�ءI
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

                DataTable dtblXuNi = dao.GetList("select CardType_RID from dbo.GROUP_CARD_TYPE a inner join CARD_GROUP b on a.Group_rid=b.rid where b.Group_Name = '�����d'").Tables[0];


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

                    //�p�G�w�s�p��s�A�h�o�eĵ�T�I
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�ھ�Perso�t�Ӫ��O���p��O�_���������d�ءI, SendWarningPerso����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
        }


        /// <summary>
        ///��l��DT
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
        /// ���ҶפJ�r�q�O�_�����榡
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
                        strErr = "��" + (count - 1).ToString() + "���" + num.ToString() + "�C�榡������1��Ʀr;\n";
                    }
                    break;
                case 2:
                    Pattern = @"^\d{9}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "��" + (count - 1).ToString() + "���" + num.ToString() + "�C�榡������9��Ʀr;\n";
                    }
                    break;
                case 3:
                    Pattern = @"^\d{1,5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "��" + (count - 1).ToString() + "���" + num.ToString() + "�C�榡������5��H�����Ʀr;\n";
                    }
                    break;
                default:
                    break;
            }

            return strErr;
        }
        /// <summary>
        /// �ˬd���p�N�X�r�q�������ѼƦb[�t�ΰѼ���]���O�_�s�b�A�}���`�ϥ�
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�ˬd���p�N�X�r�q�������ѼƦb[�t�ΰѼ���]���O�_�s�b�A�}���`�ϥ�, CheckFileStatus����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dsCARDTYPE_STATUS;
        }
        /// <summary>
        ///�d�߷�e���ʫH�����ɶ��O�_�b���Ĵ����}�B�ϥΪ��A�O�_�O���ΡA�p�G�L�d�ߵ��G�A�h�ߥX���`�A���~�����
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�d�߷�e���ʫH�����ɶ��O�_�b���Ĵ����}�B�ϥΪ��A�O�_�O����, CheckCARD_TYPE_EndTime����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dsCARD_TYPE;
        }
        /// <summary>
        ///�ˬd��e���ʫH���O�_�w�פJ�L�]����²��+���p�N�X+�פJ���+Perso�t�N�����_�A�h��ܷ�e���ʫH���w�פJ�L�^
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�ˬd��e���ʫH���O�_�w�פJ�L, CheckFACTORY_CHANGE_IMPORT����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dsFACTORY_CHANGE_IMPORT;
        }
        /// <summary>
        /// ���ҶפJ�r�q�Ĥ@��O�_�����榡
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

                        strErr = "�פJ��󪺶פJ���������8��Ʀr\n";
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
                            strErr = "�פJ��󪺶פJ����榡�����T\n";

                        }
                    }
                    break;
                case 2:
                    Pattern = @"^\d{5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {

                        strErr = "�פJ���Perso�t�ӥ�����5��Ʀr\n";
                    }
                    break;
                default:
                    break;
            }

            return strErr;
        }
        /// <summary>
        /// ��oPerso�t��RID
        /// </summary>
        /// <returns>DataSet[Perso�t��]</returns>
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("��oPerso�t��RID, GetFactoryList_RID����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return RID;
        }
        /// <summary>
        /// �ˬdFTP�ɮ׳W�h
        /// </summary>
        /// <param name="FileName">�ɮצW��(���W)</param>
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
            if (!(bl.CheckWorkDate(Convert.ToDateTime(ImportEDate.Substring(0, 4) + "/" + ImportEDate.Substring(4, 2) + "/" + ImportEDate.Substring(6, 2))))) //�D�u�@�骽����^�A������妸
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
        /// �ˬd�ɮצW��
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns>�ɮצW�٥��T��^�r�Ŧ�Ʋ�</returns>
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
        /// �ˬdPerso�^��²��
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>�^��²�٦s�b��^true</returns>
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�ˬdPerso�^��²��, CheckEnPersoExist����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }
        /// <summary>
        /// �ˬd����O�_�w�鵲
        /// </summary>
        /// <param name="ImportSDate">�϶�����_�l</param>
        /// <param name="ImportEDate">�϶��������</param>
        /// <returns>�w�鵲�L��^true</returns>
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
        /// �ˬd�O�_�w�פJ�L
        /// </summary>
        /// <param name="ImportFileName"></param>
        /// <returns>�w�פJ�L��^true</returns>
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
        /// �ˬd�O�_�������ƪ�^
        /// </summary>
        /// <param name="commondSring">SQL���O</param>
        /// <param name="dirValues">�Ѽ�</param>
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
                LogFactory.Write("�ˬd�O�_�������ƪ�^, CheckValue���~�G" + commondSring, GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        /// <summary>
        /// ���oPerso�t�ӫH��
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���oPerso�t�ӫH��, GetPerso����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }

        }

        /// <summary>
        /// ��oPerso�t��
        /// </summary>
        /// <returns>DataSet[Perso�t��]</returns>
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("��oPerso�t��, GetFactoryList����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
        }

        /// <summary>
        /// ���oPerso�t�ӫH��
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
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���oPerso�t�ӫH��, GetCardType����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }

        }

    }
}
