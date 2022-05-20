//******************************************************************
//*  �@    �̡GRay
//*  �\�໡���G�p�p�ɶפJ�ץX�妸
//*  �Ыؤ���G2008-11-14
//*  �ק����G
//*  �ק�O���G
//*            ��2009-09-02
//*                �ק� ����
//*                      1.���J�p�p�ɸ�T(�����e)
//*                      
//*******************************************************************
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
    class InOut001BL : BaseLogic
    {
        #region SQL�w�q
        //��ܩҦ����t�Ӹ�ơI
        public const string SEL_FACTORY = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN "
                    + "FROM FACTORY AS F "
                    + "WHERE F.RST = 'A' AND F.Is_Perso = 'Y' order by RID";

        const string SEL_SUBTOTAL_FILENAME_COUNT = "SELECT count(*) FROM IMPORT_PROJECT WHERE RST= 'A' AND Type = '1' AND [File_Name] = @File_Name";
        const string SEL_SUBTOTAL_PERSO_FACTORY = "SELECT COUNT(*) FROM FACTORY WHERE RST = 'A' AND Is_Perso = 'Y' AND Factory_ShortName_EN = @PersoEnName";
        const string CON_SUBTOTAL_SURPLUS = "SELECT COUNT(*) FROM CARDTYPE_STOCKS WHERE RST = 'A' AND (Stock_Date >= @ImportSDate AND Stock_Date <= @ImportEDate) ";
        const string CON_IMPORT_SUBTOTAL_CHECK = "SELECT COUNT(*) FROM SUBTOTAL_IMPORT WHERE RST = 'A' AND Import_FileName = @Import_FileName";


        public const string SEL_CARD_TYPE = "SELECT * FROM CARD_TYPE WHERE RST='A' AND TYPE = @TYPE  AND AFFINITY = @AFFINITY AND PHOTO = @PHOTO";
        public const string SEL_CARD_TYPE_1 = "SELECT * FROM CARD_TYPE WHERE RST='A' AND RID = @Change_Space_RID";
        public const string SEL_CARD_TYPE_2 = "SELECT * FROM CARD_TYPE WHERE RST='A' AND RID = @Replace_Space_RID";
        public const string SEL_CARD_TYPE_Space_RID = "SELECT * FROM CARD_TYPE WHERE RST='A' ";
        public const string SEL_SUBTOTAL_FILENAME = "SELECT * FROM IMPORT_PROJECT WHERE RST= 'A' AND Type = '1' AND File_Name = @File_Name";
        public const string CON_SUBTOTAL_PERSO_FACTORY = "SELECT RID,Factory_ID FROM FACTORY WHERE RST = 'A' AND Is_Perso = 'Y' AND Factory_ShortName_EN = @factory_shortname_en";

        public const string SEL_CARDGROUP = "select b.cardgroup_rid from dbo.IMPORT_PROJECT a inner join MAKE_CARD_TYPE b on a.makecardtype_rid=b.rid where file_name=@file_name";

        public const string SEL_CARDGROUP_BY_CARD = "select group_rid from GROUP_CARD_TYPE where group_rid in (select rid from card_group where param_code='use2') and cardtype_rid =@cardtype_rid";

        // ���Ʈ��ӳ�ĵ       
        public const string SEL_MADE_CARD_WARNNING = "SELECT * FROM (SELECT SI.Perso_Factory_RID,CT.RID,SI.MakeCardType_RID,SUM(SI.Number) AS Number "
                            + "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                            + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                            + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time "
                            + "GROUP BY SI.Perso_Factory_RID,CT.RID,SI.MakeCardType_RID "
                            + "UNION SELECT FCI.Perso_Factory_RID,FCI.CareType_RID,FCI.Status_RID,SUM(FCI.Number) AS Number "
                            + "FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') "
                            + "WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time "
                            + "GROUP BY FCI.Perso_Factory_RID,FCI.CareType_RID,FCI.Status_RID ) A "
                            + "ORDER BY Perso_Factory_RID,RID,MakeCardType_RID ";
        //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 start
        public const string SEL_MADE_CARD_WARNNING_REPLACE = "SELECT * FROM (SELECT SI.Perso_Factory_RID,CT.RID,SI.MakeCardType_RID,SUM(SI.Number) AS Number "
                            + "FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                            + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                            + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time "
                            + "GROUP BY SI.Perso_Factory_RID,CT.RID,SI.MakeCardType_RID "
                            + "UNION SELECT FCI.Perso_Factory_RID,FCI.CareType_RID,FCI.Status_RID,SUM(FCI.Number) AS Number "
                            + "FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') "
                            + "WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time "
                            + "GROUP BY FCI.Perso_Factory_RID,FCI.CareType_RID,FCI.Status_RID ) A "
                            + "ORDER BY Perso_Factory_RID,RID,MakeCardType_RID ";
        //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 end
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
                            //"INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID " +
                            //"INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID " +
                                //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 start
                            "LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID " +
                            "LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID " +
                                //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 end
                            "WHERE TMC.Perso_Factory_RID = @perso_factory_rid ";
        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD_DM = " SELECT DI.Serial_Number DI_Number,A.Perso_Factory_RID,A.Number " +
                            "FROM TEMP_MADE_CARD A " +
                            "INNER JOIN DM_CARDTYPE DCT ON DCT.RST = 'A' AND A.CardType_RID = DCT.CardType_RID " +
                            "INNER JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND DCT.DM_RID = DI.RID " +
                            "WHERE A.Perso_Factory_RID = @perso_factory_rid ";
        public const string SEL_LAST_WORK_DATE = "SELECT TOP 1 Date_Time " +
                        "FROM WORK_DATE " +
                        "WHERE Date_Time < @date_time AND Is_WorkDay='Y' " +
                        "ORDER BY Date_Time DESC";
        public const string SEL_MATERIEL_STOCKS_MANAGER = "SELECT Top 1 MSM.Stock_Date,MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number," +
                            "CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME " +
                        "FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number " +
                            "LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number " +
                            "LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number " +
                        "WHERE MSM.Type = '4' AND MSM.Perso_Factory_RID = @perso_factory_rid AND MSM.Serial_Number = @serial_number " +
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

        public const string SEL_UsedMaterial = "select  b.Serial_Number  from SUBTOTAL_IMPORT a "
                       + " inner join "
                       + " (select a.*,b.Serial_Number from card_type a "
                       + " inner join dbo.CARD_EXPONENT b on a.Exponent_rid = b.rid) b on a.type=b.type and a.affinity=b.affinity and a.photo=b.photo"
                       + " where import_fileName = @import_fileName"
                       + " union"
                       + " select  b.Serial_Number  from SUBTOTAL_IMPORT a "
                       + " inner join "
                       + " (select a.*,b.Serial_Number from card_type a "
                       + " inner join dbo.ENVELOPE_INFO b on a.Envelope_rid = b.rid) b on a.type=b.type and a.affinity=b.affinity and a.photo=b.photo"
                       + " where import_fileName = @import_fileName"
                       + " union"
                       + " select  b.Serial_Number  from SUBTOTAL_IMPORT a "
                       + " inner join"
                       + " (select A.MakeCardType_RID,B.Serial_Number from DM_MAKECARDTYPE A inner join DMTYPE_INFO B on A.DM_RID=B.RID) b"
                       + " on a.makecardtype_rid=b.makecardtype_rid"
                       + " where import_fileName = @import_fileName";

        #endregion
        //public ArrayList erro;
        public string strErr;
        public DataTable dtCardType;

        Dictionary<string, object> dirValues = new Dictionary<string, object>();

        #region �U���p�p��

        public ArrayList DownloadSubtotal()
        {
            ArrayList FileNameList = new ArrayList();

            try  // �[try catch add by judy 2018/03/28
            {
                #region Attribute


                string FolderYear = DateTime.Now.ToString("yyyy");
                string FolderDate = DateTime.Now.ToString("MMdd");
                string FolderName = "";
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.SUBTOTAL);
                string ftpPath = ConfigurationManager.AppSettings["FTPRemoteSubtotal"];
                string locPath = ConfigurationManager.AppSettings["SubTotalFilesPath"];
                string[] fileList;
                string[] fileMethod;
                bool returnFlag;


                #endregion
                fileList = ftp.GetFileList(ftpPath);
                if (fileList != null)
                {
                    foreach (string FileName in fileList)
                    {
                        if (!CheckFile(FileName)) //�ˬd�n�U�����ɮ׳W�h�A�������h���L
                        {
                            continue;
                        }
                        fileMethod = AnalysisName(FileName, '-');
                        if (fileMethod != null)
                        {
                            FolderName = locPath + "\\" + fileMethod[0] + "\\";

                            //�p�G�ؿ����s�b�A�h�s�إߥؿ��I
                            if (!System.IO.Directory.Exists(FolderName))
                            {
                                System.IO.Directory.CreateDirectory(FolderName);
                            }

                            returnFlag = ftp.Download(ftpPath, FileName, FolderName, FileName);
                            if (returnFlag)
                            {
                                string[] FList = new string[2];
                                FList[0] = FolderName;
                                FList[1] = FileName;
                                FileNameList.Add(FList);
                                // Legend 2017/05/19 �]���� �N�R���ɮץN�X�`��   �]��I�Ѷ}�`��
                                // Legend 2017/08/31 �N���B�R�ɪ`����UAT����, �W�u�O�A�Ѷ} todo
                                returnFlag = ftp.Delete(ftpPath, FileName);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogFactory.Write("�פJ�p�p�ɤU��FTP�ɮ�DownloadSubtotal��k����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

            return FileNameList;
        }
        /// <summary>
        /// ���R�ɮצW��
        /// </summary>
        /// <param name="FileName">�ɮצW��</param>
        /// <param name="Separator">���j�Ÿ�</param>
        /// <returns></returns>
        private string[] AnalysisName(string FileName, char Separator)
        {
            string[] FileList = new string[3];

            if (!StringUtil.IsEmpty(FileName))
            {
                string[] tmpList = FileName.Split(Separator);
                FileList[0] = tmpList[0].Substring(0, 8);
                FileList[1] = tmpList[0].Substring(8);
                FileList[2] = tmpList[1];
                return FileList;
            }
            else
                return null;
        }
        #region �ˬd�ɮ׳W�h
        /// <summary>
        /// �ˬd�ɮצW�� MMDD�p�p�ɦW_Penro�^��²��
        /// </summary>
        /// <param name="FileName"></param>
        /// <returns>�ɮצW�٥��T��^�r�Ŧ�Ʋ�</returns>
        private string[] CheckFileNameFormat(string FileName)
        {
            string[] fileSplit = AnalysisName(FileName, '-');
            if (fileSplit == null || fileSplit.Length != 3)
                return null;
            else
            {
                return fileSplit;
            }
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

            if (!CheckFileNameExist(fileSplit[1]))
            {
                return false;
            }
            string EnName = fileSplit[2].Substring(0, fileSplit[2].Length - 4);
            if (!CheckEnPersoExist(EnName))
            {
                return false;
            }

            DateTime dtNow = DateTime.Now;

            string ImportSDate = dtNow.ToString("yyyy-MM-dd 00:00:00.000");
            string ImportEDate = dtNow.ToString("yyyy-MM-dd 23:59:59.999");
            if (CheckImportDate(ImportSDate, ImportEDate))
            {
                return false;
            }
            if (CheckImportFile(FileName))
            {
                return false;
            }

            BatchBL bl = new BatchBL();

            // Legend 2017/05/17 �ק�
            // 1. �P�_Batch service������O�_���u�@��A�Y���O�u�@��N������
            // 2.������FTP�U���ɮסA���ݧP�_�ɦW�O�_���u�@�骺���
            // 3.�פJ�ܸ�Ʈw���A�g�J�������������
            //if (!(bl.CheckWorkDate(Convert.ToDateTime(fileSplit[0].Substring(0, 4) + "/" + fileSplit[0].Substring(4, 2) + "/" + fileSplit[0].Substring(6, 2))))) //�D�u�@�骽����^�A������妸
            if (!(bl.CheckWorkDate(dtNow))) //�D�u�@�骽����^�A������妸
            {
                return false;
            }

            return true;
        }
        /// <summary>
        /// �ˬd�p�p�ɦW�٬O�_�s�b
        /// </summary>
        /// <param name="FileName">�p�p�ɦW��</param>
        /// <returns>�ɮצs�b��^true</returns>
        private bool CheckFileNameExist(string FileName)
        {
            dirValues.Clear();
            dirValues.Add("File_Name", FileName);
            return CheckValue(SEL_SUBTOTAL_FILENAME_COUNT, dirValues);
        }
        /// <summary>
        /// �ˬdPerso�^��²��
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>�^��²�٦s�b��^true</returns>
        private bool CheckEnPersoExist(string EnPerso)
        {
            dirValues.Clear();
            dirValues.Add("PersoEnName", EnPerso);
            return CheckValue(SEL_SUBTOTAL_PERSO_FACTORY, dirValues);
        }
        /// <summary>
        /// �ˬd����O�_�w�鵲
        /// </summary>
        /// <param name="ImportSDate">�϶�����_�l</param>
        /// <param name="ImportEDate">�϶��������</param>
        /// <returns>�w�鵲�L��^true</returns>
        private bool CheckImportDate(string ImportSDate, string ImportEDate)
        {
            dirValues.Clear();
            dirValues.Add("ImportSDate", ImportSDate);
            dirValues.Add("ImportEDate", ImportEDate);
            return CheckValue(CON_SUBTOTAL_SURPLUS, dirValues);
        }
        /// <summary>
        /// �ˬd�p�p�ɬO�_�w�פJ�L
        /// </summary>
        /// <param name="ImportFileName"></param>
        /// <returns>�w�פJ�L��^true</returns>
        private bool CheckImportFile(string ImportFileName)
        {
            dirValues.Clear();
            dirValues.Add("Import_FileName", ImportFileName);
            return CheckValue(CON_IMPORT_SUBTOTAL_CHECK, dirValues);
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
                LogFactory.Write("Ū����Ʈw���~�G" + commondSring + " ���~�H�����G" + ex.Message, GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        #endregion

        #endregion

        #region ���Ʈ��ӳ�ĵ

        /// <summary>
        /// �ھڶפJ���p�p�ɡA�ͦ����Ʈ��ӰO���A�}�P�_���Ʈw�s�O�_�b�w������A
        /// �p�G���b�w������A��ĵ�C
        /// </summary>
        /// <param name="strFactory_RID"></param>
        /// <param name="importDate"></param>
        //public void Material_Used_Warnning(string strFactory_RID,
        //    DateTime importDate, string strFileName)
        //{
        //    try
        //    {
        //        // ���̫�鵲����C
        //        DateTime TheLastestSurplusDate = getLastSurplusDate();

        //        #region �p��q�̫�@�Ѥ鵲����ĤG�Ѩ�p�p�ɶפJ��Ѫ��s���d�ơA�O�s���{�ɪ�(TEMP_MADE_CARD)
        //        dirValues.Clear();
        //        dirValues.Add("Perso_Factory_RID", strFactory_RID);
        //        dirValues.Add("From_Date_Time", TheLastestSurplusDate.ToString("yyyy/MM/dd 23:59:59"));
        //        dirValues.Add("End_Date_Time", importDate.ToString("yyyy/MM/dd 23:59:59"));
        //        DataSet dsMade_Card = dao.GetList(SEL_MADE_CARD_WARNNING, dirValues);
        //        //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 start
        //        //DataSet dsMade_Card = dao.GetList(SEL_MADE_CARD_WARNNING_REPLACE, dirValues);
        //        //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 end
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
        //                DataRow[] drEXPRESSIONS = dsEXPRESSIONS_DEFINE.Tables[0].Select("RID = " + dr["MakeCardType_RID"].ToString());
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
        //                DataRow[] drEXPRESSIONS = dsEXPRESSIONS_DEFINE.Tables[0].Select("RID = " + dr["MakeCardType_RID"].ToString());
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
        //        DataTable dtMATERIAL_USED = getMaterialUsed(strFactory_RID, importDate);

        //        // �p�⪫�ƳѧE�ƶq�}ĵ��
        //        getMaterielStocks(TheLastestSurplusDate,
        //                strFactory_RID,
        //                importDate,
        //                dtMATERIAL_USED, strFileName);
        //    }
        //    catch (Exception ex)
        //    {
        //        LogFactory.Write(ex.ToString(), GlobalString.LogType.ErrorCategory);
        //        BatchBL Bbl = new BatchBL();
        //        Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
        //    }
        //}

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
        /// �p�⪫�ƳѧE�ƶq�}ĵ��
        /// </summary>
        /// <param name="dtLastWorkDate"></param>
        /// <param name="strFactory_RID"></param>
        /// <param name="importDate"></param>
        /// <param name="dtMATERIAL_USED"></param>
        //public void getMaterielStocks(DateTime dtLastestSurplus,
        //        string strFactory_RID,
        //        DateTime importDate,
        //        DataTable dtMATERIAL_USED, string strFileName)
        //{
        //    try
        //    {
        //        Depository010BL bl010 = new Depository010BL();

        //        #region �ھګe�@�Ѫ��w�s�Τ��Ѫ��w�s�C�p�⪫�ƳѾl�ƶq�A�P�_�O�_��ĵ
        //        dirValues.Clear();
        //        dirValues.Add("import_fileName", strFileName);
        //        DataTable dtblMaterial = dao.GetList(SEL_UsedMaterial, dirValues).Tables[0];

        //        foreach (DataRow drMATERIAL_USED in dtMATERIAL_USED.Rows)
        //        {
        //            if (dtblMaterial.Rows.Count > 0)
        //            {
        //                if (dtblMaterial.Select("Serial_Number='" + drMATERIAL_USED["Serial_Number"].ToString() + "'").Length == 0)
        //                    continue;
        //            }

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
        //                dirValues.Add("end_stock_date", dtLastestSurplus.ToString("yyyy/MM/dd 23:59:59"));
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
        //                    if (intLastStockNumber < 0)
        //                    {
        //                        // �w�s����
        //                        //string[] arg = new string[1];
        //                        //arg[0] = dsMaterielStocksManager.Tables[0].Rows[0]["Name"].ToString();
        //                        //Warning.SetWarning(GlobalString.WarningType.SubtotalMaterialInMiss, arg);
        //                    }
        //                    // �p�G�e�@�Ѫ��w�s�p�E���Ѫ�����
        //                    else if (intLastStockNumber < (intUsedMaterialFront + intUsedMaterialAfter))
        //                    {
        //                        if (bl010.DmNotSafe_Type(drMATERIAL_USED["Serial_Number"].ToString()))
        //                        {
        //                            // �w�s����
        //                            string[] arg = new string[1];
        //                            arg[0] = dsMaterielStocksManager.Tables[0].Rows[0]["Name"].ToString();
        //                            //Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInMiss, arg);
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
        /// �Τp�p�ɥͦ��d�����������ƯӥΰO��
        /// </summary>
        /// <returns></returns>
        //public DataTable getMaterialUsed(string strFactory_RID, DateTime importDate)
        //{
        //    DataTable dtUSE_CARDTYPE = new DataTable();
        //    Depository010BL BL010 = new Depository010BL();
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
        //                    //drNewCARD_EXPONENT["Number"] = Convert.ToInt32(dr["Number"]);
        //                    //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 start
        //                    drNewCARD_EXPONENT["Number"] = BL010.ComputeMaterialNumber(dr["CE_Number"].ToString(), Convert.ToInt64(dr["Number"]));
        //                    //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 end
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
        //                    //drNewENVELOPE_INFO["Number"] = Convert.ToInt32(dr["Number"]);
        //                    //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 start
        //                    drNewENVELOPE_INFO["Number"] = BL010.ComputeMaterialNumber(dr["EI_Number"].ToString(), Convert.ToInt64(dr["Number"]));
        //                    //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 end
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
        //                    //drNewDMTYPE_INFO["Number"] = Convert.ToInt32(dr["Number"]);
        //                    //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 start
        //                    drNewDMTYPE_INFO["Number"] = BL010.ComputeMaterialNumber(dr["DI_Number"].ToString(), Convert.ToInt64(dr["Number"]));
        //                    //200908CR�ӥζq=���������e���p�p�ɼƶq*�]1+�ӷl�v�^ add  by ���� 2009/09/01 end
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
        #endregion ���Ʈ��ӳ�ĵ


        /// <summary>
        ///�פJ���
        /// </summary>
        public string ImportSubTotal(DataSet dst, string strPath, string file_name)
        {
            string strReturn = "";
            try
            {
                DataTable dt = dst.Tables[0];
                SUBTOTAL_IMPORT SI = new SUBTOTAL_IMPORT();
                // 200908CR���J�p�p�ɸ�T(�����e) add by ���� 2009/09/04 start
                DataTable dtr = dst.Tables[0].Copy();
                SUBTOTAL_REPLACE_IMPORT SRI = new SUBTOTAL_REPLACE_IMPORT();
                // 200908CR���J�p�p�ɸ�T(�����e) add by ���� 2009/09/04 end

                int FileNameLen = file_name.LastIndexOf('-');
                string[] str = file_name.Split('-');
                string Factory_ShortName_EN = file_name.Substring(FileNameLen + 1, file_name.Length - FileNameLen - 5);
                string strFactory_ID = GetFactory_ID(Factory_ShortName_EN);
                string MakeCardType_RID = GetMakeCardType_RID(str[0].Substring(8).ToString());

                string sdtImport = str[0].Substring(0, 4) + "-" + str[0].Substring(4, 2) + "-" + str[0].Substring(6, 2);

                // Legend 2017/05/17 �N ����  ���  �אּ��e�t�Τ��
                // 1. �P�_Batch service������O�_���u�@��A�Y���O�u�@��N������
                // 2.������FTP�U���ɮסA���ݧP�_�ɦW�O�_���u�@�骺���
                // 3.�פJ�ܸ�Ʈw���A�g�J�������������
                //DateTime dtImport =  Convert.ToDateTime (sdtImport );
                DateTime dtImport = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));

                if (strFactory_ID == "")
                {
                    LogFactory.Write("���s�b��Perso�t�I", GlobalString.LogType.ErrorCategory);
                }
                //DataSet dsSpace_RID = null;

                #region ���d�ާ@
                foreach (DataRow drowFileImp in dt.Rows)
                {
                    if (drowFileImp["Action"].ToString() == "5")
                    {
                        if (Convert.ToInt32(drowFileImp["Change_Space_RID"]) != 0)
                        {
                            for (int intLoop = 0; intLoop < dtCardType.Rows.Count; intLoop++)
                            {
                                if (Convert.ToInt32(dtCardType.Rows[intLoop]["RID"]) ==
                                    Convert.ToInt32(drowFileImp["Change_Space_RID"]))
                                {
                                    drowFileImp["TYPE"] = dtCardType.Rows[intLoop]["TYPE"];
                                    drowFileImp["AFFINITY"] = dtCardType.Rows[intLoop]["AFFINITY"];
                                    drowFileImp["PHOTO"] = dtCardType.Rows[intLoop]["PHOTO"];
                                    break;
                                }
                            }
                        }
                        else if (Convert.ToInt32(drowFileImp["Replace_Space_RID"]) != 0)
                        {
                            for (int intLoop = 0; intLoop < dtCardType.Rows.Count; intLoop++)
                            {
                                if (Convert.ToInt32(dtCardType.Rows[intLoop]["RID"]) ==
                                    Convert.ToInt32(drowFileImp["Replace_Space_RID"]))
                                {
                                    drowFileImp["TYPE"] = dtCardType.Rows[intLoop]["TYPE"];
                                    drowFileImp["AFFINITY"] = dtCardType.Rows[intLoop]["AFFINITY"];
                                    drowFileImp["PHOTO"] = dtCardType.Rows[intLoop]["PHOTO"];
                                    break;
                                }
                            }
                        }
                    }
                    else if (drowFileImp["Action"].ToString() == "1" ||
                                drowFileImp["Action"].ToString() == "2" ||
                                drowFileImp["Action"].ToString() == "3")
                    {
                        if (Convert.ToInt32(drowFileImp["Replace_Space_RID"]) != 0)
                        {
                            for (int intLoop = 0; intLoop < dtCardType.Rows.Count; intLoop++)
                            {
                                if (Convert.ToInt32(dtCardType.Rows[intLoop]["RID"]) ==
                                    Convert.ToInt32(drowFileImp["Replace_Space_RID"]))
                                {
                                    drowFileImp["TYPE"] = dtCardType.Rows[intLoop]["TYPE"];
                                    drowFileImp["AFFINITY"] = dtCardType.Rows[intLoop]["AFFINITY"];
                                    drowFileImp["PHOTO"] = dtCardType.Rows[intLoop]["PHOTO"];
                                    break;
                                }
                            }
                        }
                    }
                }
                #endregion ���d�ާ@

                //�ưȶ}�l
                dao.OpenConnection();
                // 200908CR���J�p�p�ɸ�T(�����e) add by ���� 2009/09/04 start
              
                foreach (DataRow dr in dtr.Rows)
                {
                    SRI.Action = dr["ACTION"].ToString();
                    SRI.Old_CardType_RID = Convert.ToInt32(dr["Old_CardType_RID"]);
                    SRI.TYPE = dr["TYPE"].ToString();
                    SRI.AFFINITY = dr["AFFINITY"].ToString();
                    SRI.PHOTO = dr["PHOTO"].ToString();
                    SRI.Number = Convert.ToInt32(dr["Number"].ToString());
                    SRI.Date_Time = dtImport;
                    SRI.Perso_Factory_RID = Convert.ToInt32(strFactory_ID);
                    SRI.MakeCardType_RID = Convert.ToInt32(MakeCardType_RID);
                    SRI.Replace_Space_RID = Convert.ToInt32(dr["Replace_Space_RID"]);
                    SRI.Change_Space_RID = Convert.ToInt32(dr["Change_Space_RID"]);
                    SRI.Import_FileName = file_name;
                    SRI.Is_Check = "N";
                    SRI.Check_Date = Convert.ToDateTime("1900-01-01");
                    dao.Add<SUBTOTAL_REPLACE_IMPORT>(SRI, "RID");
                }
                // 200908CR���J�p�p�ɸ�T(�����e) add by ���� 2009/09/04 end
                // ���J�p�p�ɸ�T(�����Z)
                foreach (DataRow dr in dt.Rows)
                {
                    SI.Action = dr["ACTION"].ToString();
                    SI.Old_CardType_RID = Convert.ToInt32(dr["Old_CardType_RID"]);
                    SI.TYPE = dr["TYPE"].ToString();
                    SI.AFFINITY = dr["AFFINITY"].ToString();
                    SI.PHOTO = dr["PHOTO"].ToString();
                    SI.Number = Convert.ToInt32(dr["Number"].ToString());
                    SI.Date_Time = dtImport ;
                    SI.Perso_Factory_RID = Convert.ToInt32(strFactory_ID);
                    SI.MakeCardType_RID = Convert.ToInt32(MakeCardType_RID);
                    SI.Import_FileName = file_name;
                    SI.Is_Check = "N";
                    SI.Check_Date = Convert.ToDateTime("1900-01-01");
                    dao.Add<SUBTOTAL_IMPORT>(SI, "RID");
                }

                //�ưȴ���
                dao.Commit();

                this.CheckWarnningSend(dt, strFactory_ID);

                //Material_Used_Warnning(strFactory_ID, DateTime.Now,file_name);
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� ADD BY ���� 2009/08/31 start
                InOut000BL BL000 = new InOut000BL();
                BL000.Material_Used_Warnning(strFactory_ID, DateTime.Now, "1");
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� ADD BY ���� 2009/08/31 end
            }
            catch (Exception ex)
            {

                strReturn = "error";
                dao.Rollback();
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("ImportSubTotal����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                BatchBL Bbl = new BatchBL();
                Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
            }
            finally
            {
                dao.CloseConnection();
            }
            return strReturn;
        }


        /// <summary>
        /// �ˬd�ݥ����פJ���d�جO�_�����������p�I
        /// </summary>
        /// <param name="dtImport"></param>
        private void CheckWarnningSend(DataTable dtImport , string sFactoryRid)
        {
            DataTable dtCardType = this.GetCardType();
            DataTable dtFactory = this.GetFactoryList().Tables[0];
            int iNum = 0;

            DataTable dtblXuNi = dao.GetList("select CardType_RID from dbo.GROUP_CARD_TYPE a inner join CARD_GROUP b on a.Group_rid=b.rid where b.Group_Name = '�����d'").Tables[0];

            foreach (DataRow dr in dtImport.Rows)
            {
                DataRow[] drCardType = dtCardType.Select("TYPE='" + dr["TYPE"].ToString().PadLeft (3,'0') + "' and AFFINITY='"
                    + dr["AFFINITY"].ToString().PadLeft (4,'0') + "' and PHOTO='" + dr["PHOTO"].ToString().PadLeft (2,'0') + "'");

                if (drCardType.Length == 0)
                    continue;

                if (dtblXuNi.Rows.Count > 0)
                {
                    if (dtblXuNi.Select("CardType_RID = '" + drCardType[0]["RID"].ToString() + "'").Length > 0)
                        continue;
                }

                DataRow[] drFactory = dtFactory.Select("RID='" + sFactoryRid + "'");

                CardTypeManager ctm = new CardTypeManager();
                iNum = ctm.getCurrentStock(Convert.ToInt32 (sFactoryRid), Convert.ToInt32(drCardType[0]["RID"]), DateTime.Now.AddDays (1).Date);

                //�p�G�w�s�p��s�A�h�o�eĵ�T�I
                if (iNum < 0)
                {
                    object[] arg = new object[2];
                    arg[0] = drFactory[0]["Factory_Shortname_CN"];


                    if (drCardType.Length > 0)
                    {
                        arg[1] = drCardType[0]["NAME"];
                    }
                    else
                    {
                        arg[1] = "";
                    }
                    Warning.SetWarning(GlobalString.WarningType.CardTypeNotEnough , arg);
                }

            }
        }

        /// <summary>
        ///�U������ˮ�
        /// </summary>
        public DataSet ImportCheck(string filename, string strPath)
        {
            //�Ѽ�
            StreamReader sr = null;
            sr = new StreamReader(strPath + filename, System.Text.Encoding.Default);
            string[] strLine;
            string strReadLine = "";
            int count = 1;
            strErr = "";
            DataTable dtblFileImp = CreatTable();
            DataSet dsCard_type = null;
            DataSet dst = new DataSet();
            string Date = filename.Substring(0, 8);
            dtCardType = GetCardType();

            string[] fileSplit = CheckFileNameFormat(filename);

            string strGroupRID = "";

            try
            {
                dirValues.Clear();
                dirValues.Add("file_name", fileSplit[1].ToString());
                DataTable dtbl = dao.GetList(SEL_CARDGROUP, dirValues).Tables[0];
                if (dtbl.Rows.Count > 0)
                    strGroupRID = dtbl.Rows[0]["cardgroup_rid"].ToString();

                while ((strReadLine = sr.ReadLine()) != null)
                {
                    strLine = new string[3];

                    string Pattern = @"\w+";
                    MatchCollection Matches = Regex.Matches(strReadLine.Replace(",", ""), Pattern, RegexOptions.IgnoreCase);

                    if (strReadLine.Contains("===="))
                    {
                        count++;
                        continue;
                    }
                    else if (strReadLine.Contains("PHOTO"))
                    {
                        count++;
                        continue;
                    }
                    else if (strReadLine.Contains("ACTION"))
                    {
                        count++;
                        continue;
                    }
                    else if (strReadLine.Contains("�`�d��"))
                    {
                        count++;
                        continue;
                    }
                    else if (Matches.Count != 3)
                    {
                        throw new Exception("Error");
                    }
                    else
                    {
                        for (int i = 0; i < Matches.Count; i++)
                        {
                            strLine[i] = Matches[i].ToString();
                        }

                        DataRow dr = dtblFileImp.NewRow();//�@�����J�ƾڮw
                        for (int i = 0; i < strLine.Length; i++)
                        {
                            int num = i + 1;
                            if (StringUtil.IsEmpty(strLine[i]))
                                throw new Exception("Error");
                            //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�C����;";
                            else
                            {
                                strErr += CheckFileColumn(strLine[i], num, count);
                                if (!StringUtil.IsEmpty(strErr))
                                    throw new Exception("Error");
                            }
                        }


                        for (int i = 0; i < strLine.Length; i++)
                        {
                            int num = i + 1;

                            if (i == 1)
                            {
                                dsCard_type = ChecCard_TypeExists(strLine[i]);
                                if (dsCard_type != null)
                                {
                                    string name = "";
                                    if (dsCard_type.Tables[0].Rows.Count != 0)
                                        name = dsCard_type.Tables[0].Rows[0]["Name"].ToString();

                                    if (dsCard_type.Tables[0].Rows.Count == 0)
                                        throw new Exception("Error");
                                    //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�C" + strLine[i] + name + "�������d�ؤ��s�b";


                                    Int32 iEndDate = Convert.ToInt32(Convert.ToDateTime(dsCard_type.Tables[0].Rows[0]["End_Time"]).ToString("yyyyMMdd"));
                                    if (Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd 00:00:00.000")) < Convert.ToDateTime(dsCard_type.Tables[0].Rows[0]["Begin_Time"]) || (Convert.ToInt32(Date) > iEndDate && iEndDate != 19000101))
                                        throw new Exception("Error");
                                    //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�C" + strLine[i] + name + "�������d�عL��";


                                    if (dsCard_type.Tables[0].Rows[0]["Is_Using"].ToString() == "N")
                                        throw new Exception("Error");
                                    //strErr += "��" + (count - 1).ToString() + "���" + num.ToString() + "�C" + strLine[i] + name + "�������d�ذ���";


                                    if (strGroupRID != "")
                                    {
                                        //string strCardRID = dsCard_type.Tables[0].Rows[0]["RID"].ToString();
                                        //dirValues.Clear();
                                        //dirValues.Add("cardtype_rid", int.Parse(strCardRID));
                                        //DataTable dtbl1 = dao.GetList(SEL_CARDGROUP_BY_CARD, dirValues).Tables[0];
                                        //if (dtbl1.Rows.Count > 0)
                                        //{
                                        //    if (dtbl1.Rows[0][0].ToString() != strGroupRID)
                                        //    {
                                        //        throw new Exception("Error");
                                        //    }
                                        //}
                                    }

                                }
                                else
                                {
                                    // Legend 2017/05/24 ��dsCard_type��null��, �߲��`
                                    throw new Exception("Error");
                                }

                            }

                        }
                        dr["Action"] = strLine[0]; 
                        dr["Old_CardType_RID"] = dsCard_type.Tables[0].Rows[0]["RID"];
                        dr["Type"] = strLine[1].Substring(0, 3);
                        dr["AFFINITY"] = strLine[1].Substring(3, 4);
                        dr["PHOTO"] = strLine[1].Substring(7, 2);
                        dr["Number"] = strLine[2];
                        dr["Change_Space_RID"] = dsCard_type.Tables[0].Rows[0]["Change_Space_RID"];
                        dr["Replace_Space_RID"] = dsCard_type.Tables[0].Rows[0]["Replace_Space_RID"];

                        dtblFileImp.Rows.Add(dr);
                    }

                    count++;
                }

                int FileNameLen = filename.LastIndexOf('-');
                string Factory_ShortName_EN = filename.Substring(FileNameLen + 1, filename.Length - FileNameLen - 5);
                if (!StringUtil.IsEmpty(strErr))
                {
                    // �榡�����T�Aĵ��
                    object[] arg = new object[2];
                    arg[0] = Factory_ShortName_EN;
                    arg[1] = strErr;
                    Warning.SetWarning(GlobalString.WarningType.SubTotalDataIn, arg);
                }
                dst.Tables.Add(dtblFileImp);

            }
            catch
            {
                int FileNameLen = filename.LastIndexOf('-');
                string Factory_ShortName_EN = filename.Substring(FileNameLen + 1, filename.Length - FileNameLen - 5);
                // �榡�����T�Aĵ��
                object[] arg = new object[2];
                arg[0] = Factory_ShortName_EN;
                arg[1] = strErr;
                Warning.SetWarning(GlobalString.WarningType.SubTotalDataIn, arg);
            }
            finally { sr.Close(); }
            return dst;
        }
        /// <summary>
        ///���d����
        /// </summary>
        private DataSet GetCard_Change_Space_RID(int Space_RID)
        {
            DataSet dsCard_Type = null;
            DataTable dttemp = new DataTable();
            DataRow[] dr = dtCardType.Select("RID='" + Space_RID + "'");
            for (int i = 0; i < dr.Length; i++)
            {
                dttemp.Rows.Add(dr[i]);
            }

            // Legend 2017/05/24 �P�_����null��, �A�ϥ�
            if (dsCard_Type != null)
            {
                dsCard_Type.Tables.Add(dttemp);
            }
            return dsCard_Type;
        }
        /// <summary>
        /// ����t�ӥN��
        /// </summary>
        /// <returns>DataSet[�W��]</returns>
        public string GetFactory_ID(string Factory_ShortName_EN)
        {
            DataSet dsGetFactory_ShortName_EN = null;
            string Factory_ID = "";
            try
            {
                dirValues.Clear();
                dirValues.Add("factory_shortname_en", Factory_ShortName_EN);
                dsGetFactory_ShortName_EN = dao.GetList(CON_SUBTOTAL_PERSO_FACTORY, dirValues);
                if (dsGetFactory_ShortName_EN.Tables[0].Rows.Count != 0)
                {
                    Factory_ID = dsGetFactory_ShortName_EN.Tables[0].Rows[0]["RID"].ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return Factory_ID;
        }
        public string GetMakeCardType_RID(string File_Name)
        {
            DataSet dsGetFactory_ShortName_EN = null;
            string MakeCardType_RID = "";
            try
            {
                dirValues.Clear();
                dirValues.Add("File_Name", File_Name);
                dsGetFactory_ShortName_EN = dao.GetList(SEL_SUBTOTAL_FILENAME, dirValues);
                if (dsGetFactory_ShortName_EN.Tables[0].Rows.Count != 0)
                {
                    MakeCardType_RID = dsGetFactory_ShortName_EN.Tables[0].Rows[0]["MakeCardType_RID"].ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

            return MakeCardType_RID;
        }

        /// <summary>
        /// �Ы��{�ɪ��Ӧs��p�p�ɸ��
        /// </summary>
        /// <returns></returns>
        public DataTable CreatTable()
        {
            DataTable dtRet = new DataTable();
            dtRet.Columns.Add(new DataColumn("Action", Type.GetType("System.Int32")));
            dtRet.Columns.Add(new DataColumn("Old_CardType_RID", Type.GetType("System.Int32")));
            dtRet.Columns.Add(new DataColumn("Type", Type.GetType("System.String")));
            dtRet.Columns.Add(new DataColumn("Affinity", Type.GetType("System.String")));
            dtRet.Columns.Add(new DataColumn("Photo", Type.GetType("System.String")));
            dtRet.Columns.Add(new DataColumn("Number", Type.GetType("System.Int32")));
            dtRet.Columns.Add(new DataColumn("Change_Space_RID", Type.GetType("System.Int32")));
            dtRet.Columns.Add(new DataColumn("Replace_Space_RID", Type.GetType("System.Int32")));
            return dtRet;
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
                        strErr = "��" + (count - 1).ToString() + "���" + num.ToString() + "�C�榡������1��Ʀr;";
                    }
                    break;
                case 2:
                    Pattern = @"^\d{9}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "��" + (count - 1).ToString() + "���" + num.ToString() + "�C�榡������9��Ʀr;";
                    }
                    break;
                case 3:
                    Pattern = @"^\d{1,5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "��" + (count - 1).ToString() + "���" + num.ToString() + "�C�榡������5��H�����Ʀr;";
                    }
                    break;
                default:
                    break;
            }

            return strErr;
        }
        /// <summary>
        ///�ˬd��Card Type�BAffinity�BPhoto�������d�جO�_�s�b
        /// </summary>
        private DataSet ChecCard_TypeExists(string strLine)
        {
            DataSet dsCard_Type = new DataSet();
            DataTable dttemp = new DataTable();
            dttemp = dtCardType.Clone();
            string TYPE = strLine.Substring(0, 3);
            string AFFINITY = strLine.Substring(3, 4);
            string PHOTO = strLine.Substring(7, 2);
            DataRow[] dr = dtCardType.Select("type='" + TYPE + "' and affinity='" + AFFINITY + "' and photo='" + PHOTO + "'");
            for (int i = 0; i < dr.Length; i++)
            {
                DataRow drte = dttemp.NewRow();
                drte.ItemArray = dr[i].ItemArray;
                dttemp.Rows.Add(drte);
            }
            dsCard_Type.Tables.Add(dttemp);
            return dsCard_Type;
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
        /// ���o�Ҧ����d�ثH���I
        /// </summary>
        /// <returns></returns>
        private DataTable GetCardType()
        {
            DataTable dt = new DataTable();
            try
            {
                dt = dao.GetList(SEL_CARD_TYPE_Space_RID).Tables[0];

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���o�Ҧ����d�ثH���I, GetCardType����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dt;

        }

    }
}
