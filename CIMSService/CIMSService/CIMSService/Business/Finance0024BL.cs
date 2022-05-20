//*****************************************
//*  �@    �̡GGaoAi
//*  �\�໡���G�S��N��ضפJ
//*  �Ыؤ���G2008-12-17
//*  �ק����G
//*  �ק�O���G
//*****************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
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

namespace CIMSBatch.Business
{
    class Finance0024BL : BaseLogic
    {
        #region SQL
        const string SEL_PERSON_BY_PERSON_CN_NAME = "SELECT * FROM FACTORY "
                    + "WHERE RST = 'A' AND Is_Perso = 'Y' AND Factory_ShortName_EN = @Factory_ShortName_EN";
        const string CON_SPECIAL_PERSO_PROJECT_IMPORT = "SELECT * FROM SPECIAL_PERSO_PROJECT_IMPORT S, Factory F"
                    + "WHERE S.RST = 'A' AND S.Project_Date = @Project_Date and  F.rid = S.Perso_Factory_RID AND F.Factory_ShortName_EN = @Factory_ShortName_EN";
        const string SEL_SPECIAL_PROJECT = "SELECT * FROM PERSO_PROJECT WHERE RST = 'A' "
                    +"AND Normal_Special = '2' AND Factory_RID = @Factory_RID";
        const string SEL_PERSON_BY_UNIT_ID = "SELECT RID FROM FACTORY WHERE RST = 'A' AND Is_Perso = 'Y'AND Factory_ID = @Factory_ID";
        const string SEL_SPECIAL_PERSO_PROJECT_IMPORT_ALL = "SELECT * FROM SPECIAL_PERSO_PROJECT_IMPORT "
                    + "WHERE RST = 'A' ORDER BY Project_Date DESC";
        private const string CHECK_FILE_IMPORT = "select * from IMPORT_HISTORY where File_Type='3' and File_Name=@File_Name";

       
        #endregion
        //�Ѽ�
        public string strErr;
        public DataTable dtPeso;
        public DataTable dtSP;
        public DataTable dtSPIA;
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        #region �U���ɮ�
        public ArrayList DownLoadModify(string FactoryPath)
        {
            ArrayList FileNameList = new ArrayList();

            try  // �[try catch add by judy 2018/03/28
            {
                string FolderYear = DateTime.Now.ToString("yyyy");
                string FolderDate = DateTime.Now.ToString("MMdd");
                string ftpPath = ConfigurationManager.AppSettings["FTPSpecialProjectFilesPath"] + "/" + FactoryPath;
                string localPath = ConfigurationManager.AppSettings["LocalSpecialProjectFilesPath"];
                string FolderName = "";
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.SPECIAL);
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
                LogFactory.Write("�פJ�t�ӯS��N��O�ΤU��FTP�ɮ�DownLoadModify��k����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }        
            return FileNameList;

        }
        #endregion
        /// <summary>
        /// �ˬdFTP�ɮ׳W�h
        /// </summary>
        /// <param name="FileName">�ɮצW��(���W)</param>
        /// <returns></returns>
        private bool CheckFile(string FileName)
        {
            try
            {
                string[] fileSplit = FileName.Split('-');
                if (fileSplit.Length != 3)
                {
                    return false;
                }
                if (fileSplit[0].ToString().ToUpper() != "PERSO")
                {
                    return false;
                }

                string EnName = fileSplit[2].Split('.')[0].Trim();
                if (!CheckEnPersoExist(EnName))
                {
                    return false;
                }

                //�Τ�����h�ˬd�N�s�O�εo�ͪ�����A�O���諸�A�]���@�Ӥ�󤤥]�A�F�@�Ӥ뤤�Ҧ��o�ͦ��S��N�s�O�Ϊ�����C
                //string ImportEDate = fileSplit[1].ToString();
                //if (CheckImportDate(ImportEDate,EnName ))
                //{
                //    return false;
                //}

                //�ˬd���פJ�O�����A�O�_�w�g���������O���I
                if (CheckImportFile(FileName))
                {
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�ˬdFTP�ɮ׳W�h, �ɦW:[" + FileName + "], CheckFile����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }


        /// <summary>
        /// �ˬd�O�_�w�פJ�L
        /// </summary>
        /// <param name="ImportFileName"></param>
        /// <returns>�w�פJ�L��^true</returns>
        private bool CheckImportFile(string ImportFileName)
        {
            DataSet dst = new DataSet();
            dirValues.Clear();
            dirValues.Add("File_Name", ImportFileName);
            dst = dao.GetList(CHECK_FILE_IMPORT, dirValues);
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
        /// �ˬdPerso�^��²��
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>�^��²�٦s�b��^true</returns>
        private bool CheckEnPersoExist(string EnPerso)
        {

            try
            {
                DataTable dtPeso = new DataTable();
                dirValues.Clear();
                dirValues.Add("Factory_ShortName_EN", EnPerso);
                dtPeso = dao.GetList(SEL_PERSON_BY_PERSON_CN_NAME, dirValues).Tables[0];
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
        /// �ˬd�@�~������O�_�w�g���S��N���
        /// </summary>

        /// <param name="ImportEDate">�@�~���</param>
        /// <returns>����^true</returns>
        private bool CheckImportDate(string ImportEDate , string Ename)
        {
            try
            {
                DataSet dst = new DataSet();
                dirValues.Clear();
                ImportEDate = ImportEDate.Substring(0, 4) + "/" + ImportEDate.Substring(4, 2) + "/" + ImportEDate.Substring(6, 2);
                dirValues.Add("Project_Date", ImportEDate);
                dirValues.Add("Factory_ShortName_EN", Ename);
                dst = dao.GetList(CON_SPECIAL_PERSO_PROJECT_IMPORT, dirValues);
                if (dst.Tables[0].Rows.Count > 0)
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
                LogFactory.Write("�ˬd�@�~������O�_�w�g���S��N���, CheckImportDate����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }

        }
        /// <summary>
        /// ���S��N��ذT��
        /// </summary>
        private void Get_SPECIAL_PROJECT(DataTable dtPeso)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("Factory_RID", dtPeso.Rows[0]["RID"].ToString());
                dtSP = dao.GetList(SEL_SPECIAL_PROJECT, dirValues).Tables[0];
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���S��N��ذT��, Get_SPECIAL_PROJECT����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// ���Ҧ��S��N��ضפJ�T��
        /// </summary>
        private void Get_SPECIAL_PERSO_PROJECT_IMPORT_ALL()
        {
            try
            {
                dtSPIA = dao.GetList(SEL_SPECIAL_PERSO_PROJECT_IMPORT_ALL).Tables[0];
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���Ҧ��S��N��ضפJ�T��, Get_SPECIAL_PERSO_PROJECT_IMPORT_ALL����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// �פJ�ɮ��ˬd
        /// </summary>

        /// <param name="strPath">�ɮ׸��|�W��</param>
        /// <param name="dtPeso">Peso�t�H��</param>
        /// <param name="dtSP">�S��N��ذT��</param>
        /// <param name="dtSPIA">�S��N��ضפJ�T��</param>
        /// <returns>dt</returns>
        public DataTable ImportCheck(string strPath,string Ename)
        {
            StreamReader sr = null;
            sr = new StreamReader(strPath, System.Text.Encoding.Default);

            try
            {
                strErr = "";

                dtPeso = new DataTable();
                dirValues.Clear();
                dirValues.Add("Factory_ShortName_EN", Ename);
                dtPeso = dao.GetList(SEL_PERSON_BY_PERSON_CN_NAME, dirValues).Tables[0];
                Get_SPECIAL_PROJECT(dtPeso);
                Get_SPECIAL_PERSO_PROJECT_IMPORT_ALL();
                string[] strLine;
                string strReadLine = "";
                int count = 1;
                DataTable dt = CreatTable();
                int PersoProject_RID;
                while ((strReadLine = sr.ReadLine()) != null)
                {
                    if (count == 1)
                    {
                        if (strReadLine.Trim() != dtPeso.Rows[0]["Factory_ID"].ToString())
                        {
                            strErr += "���ɲĤ@�椤���t�ӽs���N��Perso�t�M���ɦW����Perso�t²�٥N��Perso�t���@�P\n";
                        }
                        count++;
                    }
                    else
                    {
                        if (!StringUtil.IsEmpty(strReadLine))
                        {
                            if (StringUtil.GetByteLength(strReadLine) != 23)//�C�ƶq�ˬd
                            {
                                strErr += "��" + count.ToString() + "�榡���~�C\n";
                                //throw new AlertException("��" + count.ToString() + "��C�Ƥ����T�C");
                            }
                            else
                            {
                                // ���Φr�Ŧ�
                                int nextBegin = 0;
                                Depository010BL bl003 = new Depository010BL();
                                strLine = new string[3];
                                strLine[0] = bl003.GetSubstringByByte(strReadLine, nextBegin, 8, out nextBegin).Trim();
                                strLine[1] = bl003.GetSubstringByByte(strReadLine, nextBegin, 6, out nextBegin).Trim();
                                strLine[2] = bl003.GetSubstringByByte(strReadLine, nextBegin, 9, out nextBegin).Trim();


                                //strLine = strReadLine.Split(GlobalString.FileSplit.Split);
                                string strDate = strLine[0].Trim();
                                string SpecialCode = strLine[1].Trim();
                                string number = strLine[2].Trim();
                                DateTime Date = Convert.ToDateTime(strDate.Substring(0, 4) + "/" + strDate.Substring(4, 2) + "/" + strDate.Substring(6, 2));
                                string FRID = CheckSpecialCode(SpecialCode);
                                if (FRID == "")
                                {
                                    strErr += "��" + count.ToString() + "���S��N��ؽs�����s�b�C\n";
                                }
                                else
                                {
                                    PersoProject_RID = Convert.ToInt32(FRID);
                                }
                                if (dtSPIA.Rows.Count > 0)
                                {
                                    if (isExistSpecialCode(Convert.ToInt32(dtPeso.Rows[0]["RID"].ToString()), Date, Convert.ToInt32(FRID)))
                                    {
                                        strErr += "��" + count.ToString() + "�S��N��ضפJ�T���w�g�s�b�A���୫�_�פJ�C\n";
                                    }
                                }
                                DataRow dr = dt.NewRow();
                                dr["Perso_Factory_RID"] = dtPeso.Rows[0]["RID"].ToString();
                                dr["Project_Date"] = strDate.Substring(0, 4) + "/" + strDate.Substring(4, 2) + "/" + strDate.Substring(6, 2);
                                dr["PersoProject_RID"] = FRID;
                                dr["Number"] = number;
                                dt.Rows.Add(dr);
                            }
                        }
                        count++;
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�פJ�ɮ��ˬd, ImportCheck����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
            finally
            {
                sr.Close();
            }
        }
        /// <summary>
        ///��l��DT
        /// </summary>
        public DataTable CreatTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Perso_Factory_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Project_Date", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("PersoProject_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Number", System.Type.GetType("System.String")));
            return dt;
        }
        /// <summary>
        ///�ˬd�S��N��ؽs���O�_���T
        /// </summary>
        private string CheckSpecialCode(string spid)
        {
            DataRow[] dr = dtSP.Select("Project_Code='" + spid + "'");
            if (dr.Length > 0)
            {
                return dr[0]["RID"].ToString();
            }
            else
            {
                return "";
            }
        }
        /// <summary>
        ///�ˬd�S��N��جO�_�w�פJ
        /// </summary>
        private bool isExistSpecialCode(int persoid, DateTime Project_Date, int PersoProject_RID)
        {
            DataRow[] dr = dtSPIA.Select("Project_Date='" + Project_Date.ToShortDateString () + "' and Perso_Factory_RID='" + persoid + "' and PersoProject_RID='" + PersoProject_RID+"'");
            if (dr.Length > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool SaveSpecialIn(DataTable dt)
        {
            bool IsError = false;
            try
            {
                dao.OpenConnection();
                SPECIAL_PERSO_PROJECT_IMPORT SPPI = new SPECIAL_PERSO_PROJECT_IMPORT();
                foreach (DataRow dr in dt.Rows)
                {
                    SPPI.Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"].ToString());
                    SPPI.PersoProject_RID = Convert.ToInt32(dr["PersoProject_RID"].ToString());
                    SPPI.Project_Date = Convert.ToDateTime(dr["Project_Date"].ToString());
                    SPPI.Number = Convert.ToInt32(dr["Number"].ToString());
                    dao.Add<SPECIAL_PERSO_PROJECT_IMPORT>(SPPI, "RID");

                }
                dao.Commit();
                IsError = true;
            }
            catch (Exception ex)
            {
                dao.Rollback();
                LogFactory.Write(ex.ToString(), GlobalString.LogType.ErrorCategory);
                BatchBL Bbl = new BatchBL();
                Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
                
            }
            finally
            {
                dao.CloseConnection();
            }
            return IsError;
        }
    }
}
