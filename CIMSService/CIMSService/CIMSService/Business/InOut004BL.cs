//*****************************************
//*  �@    �̡GGaoAi
//*  �\�໡���G�~�״��d�w���ɶפJ
//*  �Ыؤ���G2008-11-25
//*  �ק����G
//*  �ק�O���G
//*****************************************

using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.IO;
using System.Text.RegularExpressions;
using CIMSBatch;
using CIMSBatch.FTP;
using CIMSBatch.Model;
using CIMSBatch.Public;

namespace CIMSBatch.Business
{
    class InOut004BL:BaseLogic
    {
        #region SQL�y�y�w�q
        const string SEL_FILE_NAME = "SELECT File_Name FROM IMPORT_PROJECT WHERE RST = 'A' AND Type = '3'";
        const string SEL_CARDTYPE = "SELECT * FROM CARD_TYPE WHERE RST = 'A' ";
        const string CON_FORE_CHANGE_CARD = "SELECT COUNT(*) FROM FORE_CHANGE_CARD AS FCC WHERE FCC.RST = 'A' ";
        //const string DEl_FORE_CHANGE_CARD = "DELETE FROM  FORE_CHANGE_CARD WHERE RST = 'A' ";
        const string SEL_CARD_TYPE = "SELECT CT.RID FROM CARD_TYPE AS CT WHERE CT.RST = 'A' ";
        const string DEL_PERSO_FORE_CHANGE_CARD = "DELETE FROM PERSO_FORE_CHANGE_CARD WHERE RST = 'A' AND Fore_RID IN (SELECT RID FROM FORE_CHANGE_CARD WHERE RST = 'A' AND Change_Date = @change_date AND Type = @type AND Affinity = @affinity AND Photo = @photo) ";
        const string DEL_FORE_CHANGE_CARD = "DELETE FROM FORE_CHANGE_CARD WHERE RST = 'A' AND Change_Date = @change_date AND Type = @type AND Affinity = @affinity AND Photo = @photo ";
        const string SEL_FORE_CHANGE_CARD = "SELECT * FROM FORE_CHANGE_CARD WHERE RST = 'A' AND Change_Date = @change_date AND Type = @type AND Affinity = @affinity AND Photo = @photo ";
        const string SEL_CARDTYPE_PERSO = "SELECT PC.*,CT.TYPE,CT.AFFINITY,CT.PHOTO FROM PERSO_CARDTYPE PC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND PC.CardType_RID = CT.RID WHERE PC.RST = 'A' AND PC.CardType_RID = @cardtype_rid ORDER BY PC.Base_Special DESC,PC.Priority ASC ";

        #endregion
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        public DataTable dtcard;
        public string strErr;
        #region �U���~�״��d�w����
        public ArrayList YearReplaceCard()
        {
            #region Attribute

            ArrayList FileNameList = new ArrayList();

            try  // �[try catch add by judy 2018/03/28
            {
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.YEARREPLACE);
                string ftpPath = ConfigurationManager.AppSettings["FTPRemoteYearReplaceCard"]; //ftp�ɮץؿ��t�m�ɫH��
                string locPath = ConfigurationManager.AppSettings["YearReplaceCardForecastFilesPath"]; //local�ɮץؿ��t�m�ɫH��
                string FolderName = ""; //���a�ؿ�
                string remFileName; //FTP�ɦW
                string locFileName; //���a�s�x�ɦW
                string[] fileList;
                string[] namelist;
                bool returnFlag;

                GetCardType(dtcard);

                fileList = GetYearReplaceCard();
                if (fileList != null)
                {
                    foreach (string FileName in fileList)
                    {
                        remFileName = "";
                        locFileName = FileName + DateTime.Now.ToString("yyyyMM") + ".txt";
                        if (ftp.GetFileList(ftpPath, remFileName) == null) //�ˬdFTP�O�_�������ɮ�
                        {
                            continue;
                        }
                        else
                        {
                            namelist = ftp.GetFileList(ftpPath, remFileName);
                        }
                        FolderName = locPath + "\\";

                        foreach (string Name in namelist)
                        {
                            if (Name != "")
                            {
                                string[] checkname = Name.Split('.');
                                if (checkname.Length == 3)
                                {
                                    returnFlag = ftp.Delete(ftpPath, Name);
                                    continue;
                                }
                                if (checkname[0].Contains(FileName))
                                {
                                    returnFlag = ftp.Download(ftpPath, Name, FolderName, locFileName);
                                    if (returnFlag)
                                    {
                                        //FileNameList.Add(FolderName + locFileName);
                                        string[] FList = new string[2];
                                        FList[0] = FolderName;
                                        FList[1] = locFileName;
                                        FileNameList.Add(FList);
                                        // Legend 2017/11/28 �N���B�R�ɪ`����UAT����, �W�u�O�A�Ѷ} todo
                                        returnFlag = ftp.Delete(ftpPath, Name);
                                    }
                                }
                                else
                                {
                                    //returnFlag = ftp.Delete(ftpPath, Name);
                                }
                            }

                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogFactory.Write("�פJ�~�״��d�w���ɤU��FTP�ɮ�YearReplaceCard��k����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
          
            return FileNameList;
        }
        #endregion
        /// <summary>
        /// ���o�~�״��d�w���ɮצW�٦C��
        /// </summary>
        /// <returns>string []</returns>
        private string[] GetYearReplaceCard()
        {
            
            try
            {
                string[] fileList;
                string FileNameDate = "";
                int index = 0;
                DataSet ds = dao.GetList(SEL_FILE_NAME);
                if (ds.Tables.Count > 0)
                {
                    fileList = new string[ds.Tables[0].Rows.Count];
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        fileList[index] = dr["File_Name"].ToString() + FileNameDate;
                        index++;
                    }
                    if (fileList.Length > 0)
                    {
                        return fileList;
                    }
                    else
                        return null;
                }
                else
                    return null;
            }
            catch(Exception ex)
            {
                LogFactory.Write("���o�~�״��d�ɦW���~:" + ex.ToString(),GlobalString.LogType.ErrorCategory );
                return null;
            }
        }

        #endregion
        /// <summary>
        ///�U������ˮ�
        /// </summary>
        public DataSet DetailCheck(string FilePath)
        {
            DataSet dtsReturn = new DataSet();
            #region ���Ҥ��
            StreamReader sr = null;
            DataSet dsCARDTYPE = null;
            try
            {
                //�s�ؼƾڪ�
                DataTable dtblFileImp = new DataTable();
                dtblFileImp.Columns.Add("Type_Code");
                dtblFileImp.Columns.Add("Affinity_Code");
                dtblFileImp.Columns.Add("StartDate");
                dtblFileImp.Columns.Add("Photo_Code");
                dtblFileImp.Columns.Add("DtDate_Number");

                dsCARDTYPE = dao.GetList(SEL_CARDTYPE);
                sr = new StreamReader(FilePath, System.Text.Encoding.Default);
                string[] strLine;
                string strReadLine = "";
                int count = 0;
                string strErr = "";
                int TableCount = 0;

                DataTable dtData_Number = new DataTable();
                dtData_Number.Columns.Add("Date");
                dtData_Number.Columns.Add("Number");

                DataRow dr = dtblFileImp.NewRow();

                string NowTime = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo);


                while ((strReadLine = sr.ReadLine()) != null)
                {
                    if (StringUtil.IsEmpty(strReadLine))
                    {
                        count++;
                        continue;
                    }

                    if (strReadLine.Contains("E-TYPE"))
                    {
                        strLine = strReadLine.Split(':');
                        if (strLine.Length != 2 || dtData_Number.Rows.Count == 0)
                            throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_003_01);

                        string strTemp = strLine[1];

                        string Pattern = @"\w+";
                        MatchCollection Matches = Regex.Matches(strTemp, Pattern, RegexOptions.IgnoreCase);
                        if (Matches.Count != 18)
                            throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_003_01);



                        if (Matches[0].Length != 2)
                            strErr += "��" + count.ToString() + "�檺E_TYPE�N�X " + Matches[0].ToString() + "�����G��;\n";


                        dr[3] = Matches[0].ToString();

                        #region ���ҥd�ت��p
                        for (int i = 1; i < Matches.Count - 1; i++)
                        {
                            if (i != Matches.Count - 1)
                                strErr += CheckNumberColumn(Matches[i].ToString(), i, count);

                            dtData_Number.Rows[i - 1][1] = Matches[i].ToString().Replace(",", "");
                        }
                        int TableNum = TableCount + 1;
                        if (dtblFileImp.Select("Type_Code='" + dr["Type_Code"].ToString() + "' AND Affinity_Code = '" + dr["Affinity_Code"].ToString() + "' AND Photo_Code = '" + dr["Photo_Code"].ToString() + "' ").Length > 0)
                        {
                            strErr += "��" + TableNum + "�� " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                            strErr += "�������d�ؤ��୫�ƶפJ!\n";
                        }
                        else
                        {

                            if (dsCARDTYPE.Tables[0].Select("TYPE='" + dr["Type_Code"].ToString() + "' AND AFFINITY = '" + dr["Affinity_Code"].ToString() + "' AND PHOTO = '" + dr["Photo_Code"].ToString() + "' ").Length == 0)
                            {
                                strErr += "��" + TableNum + "�� " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                                strErr += "�������d�ؤ��s�b!\n";
                            }
                            else
                            {

                                if (dsCARDTYPE.Tables[0].Select("Begin_Time<='" + dr["StartDate"].ToString().Substring(0, 4) + "-" + dr["StartDate"].ToString().Substring(4, 2) + "-" + "01" + "' AND (End_Time='1900-01-01' or End_Time >='" + dr["StartDate"].ToString().Substring(0, 4) + "-" + dr["StartDate"].ToString().Substring(4, 2) + "-" + "01" + "') AND TYPE='" + dr["Type_Code"].ToString() + "' AND AFFINITY = '" + dr["Affinity_Code"].ToString() + "' AND PHOTO = '" + dr["Photo_Code"].ToString() + "' ").Length == 0)
                                {
                                    strErr += "��" + TableNum + "�� " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                                    strErr += "��" + count.ToString() + "��������d�ؤ��b���Ĵ���!;\n";
                                }


                                if (dsCARDTYPE.Tables[0].Select("TYPE='" + dr["Type_Code"].ToString() + "' AND AFFINITY = '" + dr["Affinity_Code"].ToString() + "' AND PHOTO = '" + dr["Photo_Code"].ToString() + "' AND Is_Using = 'N' ").Length >= 1)
                                {
                                    strErr += "��" + TableNum + "�� " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                                    strErr += "�������d�ؤw����!\n";
                                }
                            }
                        }
                        #endregion



                        dr[4] = TableCount;
                        ++TableCount;

                        DataRow drClone = dtblFileImp.NewRow();
                        drClone[0] = dr[0];
                        drClone[1] = dr[1];
                        drClone[2] = dr[2];
                        drClone[3] = dr[3];
                        drClone[4] = dr[4];
                        dtblFileImp.Rows.Add(drClone);

                        DataTable dtblTemp = new DataTable();
                        dtblTemp = dtData_Number.Clone();
                        foreach (DataRow drow in dtData_Number.Rows)
                        {
                            DataRow drowTemp = dtblTemp.NewRow();
                            drowTemp.ItemArray = drow.ItemArray;
                            dtblTemp.Rows.Add(drowTemp);
                        }
                        dtsReturn.Tables.Add(dtblTemp);

                        count++;
                    }
                    else if (strReadLine.Contains("===="))
                    {
                        count++;
                        continue;
                    }
                    else if (strReadLine.Contains("AFFIN"))
                    {
                        dtData_Number = new DataTable();

                        dtData_Number.Columns.Add("Date");
                        dtData_Number.Columns.Add("Number");

                        strLine = strReadLine.Split(':');
                        if (strLine.Length != 2 || dr[0].ToString() == "")
                            throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_003_01);

                        string strTemp = strLine[1];

                        string Pattern = @"\w+";
                        MatchCollection Matches = Regex.Matches(strTemp, Pattern, RegexOptions.IgnoreCase);
                        if (Matches.Count != 18)
                            throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_003_01);

                        if (Matches[0].Length != 4)
                            strErr += "��" + count.ToString() + "�檺AFFIN�N�X " + Matches[0].ToString() + "�����|��;\n";

                        dr[1] = Matches[0].ToString();


                        for (int i = 1; i < Matches.Count - 1; i++)
                        {
                            if (i != Matches.Count - 1)
                                strErr += CheckDateColumn(Matches[i].ToString(), i, count);

                            if (i == 1)
                            {
                                dr[2] = Matches[i].ToString();
                            }

                            DataRow dtDn = dtData_Number.NewRow();
                            dtDn[0] = Matches[i];
                            dtDn[1] = 0;
                            dtData_Number.Rows.Add(dtDn);
                        }
                        count++;
                    }
                    else if (strReadLine.Contains("TYPE"))
                    {
                        dr = dtblFileImp.NewRow();


                        strLine = strReadLine.Split(':');
                        if (strLine.Length != 2)
                            throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_003_01);

                        dr[0] = strLine[1].Trim();

                        count++;
                    }
                    else
                    {
                        count++;
                        continue;
                    }
                }
                dtsReturn.Tables.Add(dtblFileImp);

                if (dtsReturn.Tables.Count < 2)
                    throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_003_01);

                if (!StringUtil.IsEmpty(strErr))
                {

                    // �ߥX���`
                    throw new Exception(strErr);
                }
            }
     
            catch (Exception ex)
            {
                string[] arg = new string[1];
                arg[0] = ex.Message;
                Warning.SetWarning(GlobalString.WarningType.YearChangeCardForeCast, arg);
            }
            finally
            {
                if (sr != null)
                    sr.Close();
            }
            #endregion

            return dtsReturn;
        }
        /// <summary>
        ///�פJ�~�״��d�w����
        /// </summary>
        public string In(DataSet dsFileImp)
        {
            FORE_CHANGE_CARD fccModel = null;
            DataSet dsCARDTYPE = null;

            string strReturn = "";
            try
            {
              
                dao.OpenConnection();
                foreach (DataRow dr in dsFileImp.Tables[dsFileImp.Tables.Count - 1].Rows)
                {
                    int intNum = int.Parse(dr["DtDate_Number"].ToString());
                    foreach (DataRow dr_date in dsFileImp.Tables[intNum].Rows)
                    {
                        fccModel = new FORE_CHANGE_CARD();

                        dirValues.Clear();
                        dirValues.Add("change_date", dr_date["Date"].ToString());
                        dirValues.Add("type", dr["Type_Code"].ToString());
                        dirValues.Add("affinity", dr["Affinity_Code"].ToString());
                        dirValues.Add("photo", dr["Photo_Code"].ToString());
                        Del_PERSO_FORE_CHANGE_CARD(dr["Type_Code"].ToString(), dr["Affinity_Code"].ToString(), dr["Photo_Code"].ToString(), dr_date["Date"].ToString());
                        //Del_FORE_CHANGE_CARD(dr["Type_Code"].ToString(), dr["Affinity_Code"].ToString(), dr["Photo_Code"].ToString(), dr_date["Date"].ToString());
                        //�K�[���봫�d�w���T���CDao.add(),�}���X�s�K�[�O����RID

                        int intRID = 0;

                        if (dao.GetList(SEL_FORE_CHANGE_CARD, dirValues).Tables[0].Rows.Count > 0)
                        {
                            fccModel = dao.GetModel<FORE_CHANGE_CARD>(SEL_FORE_CHANGE_CARD, dirValues);
                            fccModel.Number = Convert.ToInt64(dr_date["Number"]);
                            fccModel.IsYear = "1";
                            dao.Update<FORE_CHANGE_CARD>(fccModel, "RID");
                            intRID = fccModel.RID;
                        }
                        else
                        {
                            //�K�[���봫�d�w���T���CDao.add(),�}���X�s�K�[�O����RID
                            fccModel.Change_Date = dr_date["Date"].ToString();
                            fccModel.Type = dr["Type_Code"].ToString();
                            fccModel.Affinity = dr["Affinity_Code"].ToString();
                            fccModel.Photo = dr["Photo_Code"].ToString();
                            fccModel.Number = Convert.ToInt64(dr_date["Number"]);
                            fccModel.IsMonth = "2";
                            fccModel.IsYear = "1";
                            intRID = Convert.ToInt32(dao.AddAndGetID<FORE_CHANGE_CARD>(fccModel, "RID"));
                        }


                        //fccModel.Change_Date = dr_date["Date"].ToString();
                        //fccModel.Type = dr["Type_Code"].ToString();
                        //fccModel.Affinity = dr["Affinity_Code"].ToString();
                        //fccModel.Photo = dr["Photo_Code"].ToString();
                        //fccModel.Number = Convert.ToInt64(dr_date["Number"]);
                        //int intRID = Convert.ToInt32(dao.AddAndGetID<FORE_CHANGE_CARD>(fccModel, "RID"));

                        dsCARDTYPE = dao.GetList(SEL_CARDTYPE + "AND Type = @type AND Affinity = @affinity AND Photo = @photo", dirValues);
                        foreach (DataRow dr1 in dsCARDTYPE.Tables[0].Rows)
                        {
                            SplitToPerso(Convert.ToInt32(dr1["RID"]), intRID, Convert.ToInt64(dr_date["Number"]), dr_date["Date"].ToString());
                        }
                    }

                }
                
                dao.Commit();
                strReturn= "";
            }
            catch (Exception ex)
            {
                strReturn = "erro";
                dao.Rollback();
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�פJ�~�״��d�w����, In����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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
    /// ��Perso�t���t��h�ө��
    /// </summary>
    /// <returns></returns>
        public void SplitToPerso(int Space_RID, int intRID, long DtDate_Number, string StartDate)
        {
            PERSO_FORE_CHANGE_CARD pfccModel = new PERSO_FORE_CHANGE_CARD();
            DataSet dsCARDTYPE_PERSO = null;
            try
            {
                dirValues.Clear();
                dirValues.Add("cardtype_rid", Space_RID);
                dsCARDTYPE_PERSO = dao.GetList(SEL_CARDTYPE_PERSO, dirValues);

                if (dsCARDTYPE_PERSO.Tables[0].Rows.Count != 0)
                {
                    if (dsCARDTYPE_PERSO.Tables[0].Rows[0]["Base_Special"].ToString() == "1")
                    {
                        pfccModel.Change_Date = StartDate;
                        pfccModel.Type = dsCARDTYPE_PERSO.Tables[0].Rows[0]["Type"].ToString();
                        pfccModel.Affinity = dsCARDTYPE_PERSO.Tables[0].Rows[0]["Affinity"].ToString();
                        pfccModel.Photo = dsCARDTYPE_PERSO.Tables[0].Rows[0]["Photo"].ToString();
                        pfccModel.Perso_Factory_RID = Convert.ToInt32(dsCARDTYPE_PERSO.Tables[0].Rows[0]["Factory_RID"].ToString());
                        pfccModel.Number = DtDate_Number;
                        pfccModel.Fore_RID = intRID;
                        dao.Add<PERSO_FORE_CHANGE_CARD>(pfccModel, "RID");
                    }
                    else
                    {
                        // ���S����t�T��
                        DataRow[] drCARDTYPE_PERSO = dsCARDTYPE_PERSO.Tables[0].Select("Base_Special = '2'");
                        // ����v���t
                        if (drCARDTYPE_PERSO[0]["Percentage_Number"].ToString() == "1")
                        {
                            long intNumber = 0;
                            for (int int1 = 0; int1 < drCARDTYPE_PERSO.Length; int1++)
                            {
                                if (int1 < drCARDTYPE_PERSO.Length - 1)
                                {
                                    intNumber += Convert.ToInt64(Math.Floor(DtDate_Number * (Convert.ToDouble(drCARDTYPE_PERSO[int1]["Value"]) / 100)));
                                    pfccModel.Change_Date = StartDate;
                                    pfccModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString();
                                    pfccModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString();
                                    pfccModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString();
                                    pfccModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                    pfccModel.Number = Convert.ToInt64(Math.Floor(DtDate_Number * (Convert.ToDouble(drCARDTYPE_PERSO[int1]["Value"]) / 100)));
                                    pfccModel.Fore_RID = intRID;
                                    dao.Add<PERSO_FORE_CHANGE_CARD>(pfccModel, "RID");
                                }
                                else
                                {
                                    pfccModel.Change_Date = StartDate;
                                    pfccModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString();
                                    pfccModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString();
                                    pfccModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString();
                                    pfccModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                    pfccModel.Number = DtDate_Number - intNumber;
                                    pfccModel.Fore_RID = intRID;
                                    dao.Add<PERSO_FORE_CHANGE_CARD>(pfccModel, "RID");
                                }
                            }
                        }
                        // ���ƶq���t
                        else if (drCARDTYPE_PERSO[0]["Percentage_Number"].ToString() == "2")
                        {
                            long intNumber = 0;
                            for (int int1 = 0; int1 < drCARDTYPE_PERSO.Length; int1++)
                            {
                                if (int1 < drCARDTYPE_PERSO.Length - 1)
                                {
                                    if ((DtDate_Number - intNumber) > Convert.ToInt32(drCARDTYPE_PERSO[int1]["Value"]))
                                    {
                                        intNumber += Convert.ToInt32(drCARDTYPE_PERSO[int1]["Value"]);
                                        pfccModel.Change_Date = StartDate;
                                        pfccModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString(); ;
                                        pfccModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString(); ;
                                        pfccModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString(); ;
                                        pfccModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                        pfccModel.Number = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Value"]);
                                        pfccModel.Fore_RID = intRID;
                                        dao.Add<PERSO_FORE_CHANGE_CARD>(pfccModel, "RID");
                                    }
                                    else
                                    {
                                        pfccModel.Change_Date = StartDate;
                                        pfccModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString(); ;
                                        pfccModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString(); ;
                                        pfccModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString(); ;
                                        pfccModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                        pfccModel.Number = DtDate_Number - intNumber;
                                        pfccModel.Fore_RID = intRID;
                                        dao.Add<PERSO_FORE_CHANGE_CARD>(pfccModel, "RID");
                                        break;
                                    }
                                }
                                else
                                {
                                    pfccModel.Change_Date = StartDate;
                                    pfccModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString(); ;
                                    pfccModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString(); ;
                                    pfccModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString(); ;
                                    pfccModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                    pfccModel.Number = DtDate_Number - intNumber;
                                    pfccModel.Fore_RID = intRID;
                                    dao.Add<PERSO_FORE_CHANGE_CARD>(pfccModel, "RID");
                                }
                            }
                        }
                    }
                }
                
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("��Perso�t���t��h�ө��, SplitToPerso����:" + ex.ToString(),GlobalString.LogType.ErrorCategory);
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
            dt.Columns.Add(new DataColumn("Number", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("ChangeDate", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Change_Space_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Replace_Space_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("RID", System.Type.GetType("System.String")));
            return dt;
        }

        /// <summary>
        /// �ˬdColumn�O�_���Ʀr
        /// </summary>
        /// <param name="strColumn"></param>
        /// <param name="num"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private string CheckNumberColumn(string strColumn, int num, int count)
        {
            string strErr = "";
            string Pattern = "";
            MatchCollection Matches;
            Pattern = @"^\d+$";
            Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
            if (Matches.Count == 0)
            {
                strErr = "��" + count.ToString() + "���" + num.ToString() + "�C�榡�������Ʀr;";
            }
            return strErr;
        }
        /// <summary>
        /// ���ҶפJ�r�q�O�_�����榡
        /// </summary>
        /// <param name="strColumn"></param>
        /// <param name="num"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private string CheckDateColumn(string strColumn, int num, int count)
        {
            string strErr = "";
            if (strColumn.Length != 6)
            {
                strErr = "��" + count.ToString() + "���" + num.ToString() + "�C�榡���~;";
            }
            return strErr;
        }
        /// <summary>
        /// ���ҶפJ�r�q�O�_�����榡
        /// </summary>
        /// <param name="strColumn"></param>
        /// <param name="num"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private bool CheckDateColumn(string strColumn)
        {

            string Pattern = "";
            MatchCollection Matches;
            Pattern = @"^\d{4}$";
            Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
            if (Matches.Count == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        /// <summary>
        /// ���o�Ҧ����d�ثH��
        /// </summary>
        private void GetCardType(DataTable dt)
        {
            dt = new DataTable();

            dt = dao.GetList(SEL_CARD_TYPE).Tables[0];
        }
        /// <summary>
        /// �R��PERSO_FORE_CHANGE_CARD�ɤ��������O��
        /// </summary>
        private void Del_PERSO_FORE_CHANGE_CARD(string type, string affin, string photo,string changedate)
        {
            dirValues.Clear();
            dirValues.Add("type", type);
            dirValues.Add("affinity", affin);
            dirValues.Add("photo", photo);
            dirValues.Add("change_Date", changedate);
            dao.ExecuteNonQuery(DEL_PERSO_FORE_CHANGE_CARD, dirValues);
        }
        /// /// <summary>
        /// �R��FORE_CHANGE_CARD�ɤ��������O��
        /// </summary>
        private void Del_FORE_CHANGE_CARD(string type, string affin, string photo, string changedate)
        {
            dirValues.Clear();
            dirValues.Add("type", type);
            dirValues.Add("affinity", affin);
            dirValues.Add("photo", photo);
            dirValues.Add("change_Date", changedate);
            dao.ExecuteNonQuery(DEL_FORE_CHANGE_CARD, dirValues);
        }

    }
}
