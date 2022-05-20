//*****************************************
//*  作    者：GaoAi
//*  功能說明：特殊代制項目匯入
//*  創建日期：2008-12-17
//*  修改日期：
//*  修改記錄：
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
        //參數
        public string strErr;
        public DataTable dtPeso;
        public DataTable dtSP;
        public DataTable dtSPIA;
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        #region 下載檔案
        public ArrayList DownLoadModify(string FactoryPath)
        {
            ArrayList FileNameList = new ArrayList();

            try  // 加try catch add by judy 2018/03/28
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
                LogFactory.Write("匯入廠商特殊代制費用下載FTP檔案DownLoadModify方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }        
            return FileNameList;

        }
        #endregion
        /// <summary>
        /// 檢查FTP檔案規則
        /// </summary>
        /// <param name="FileName">檔案名稱(全名)</param>
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

                //用文件日期去檢查代製費用發生的日期，是不對的，因為一個文件中包括了一個月中所有發生有特殊代製費用的日期。
                //string ImportEDate = fileSplit[1].ToString();
                //if (CheckImportDate(ImportEDate,EnName ))
                //{
                //    return false;
                //}

                //檢查文件匯入記錄中，是否已經有相關的記錄！
                if (CheckImportFile(FileName))
                {
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查FTP檔案規則, 檔名:[" + FileName + "], CheckFile報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }


        /// <summary>
        /// 檢查是否已匯入過
        /// </summary>
        /// <param name="ImportFileName"></param>
        /// <returns>已匯入過返回true</returns>
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
        /// 檢查Perso英文簡稱
        /// </summary>
        /// <param name="EnPerso"></param>
        /// <returns>英文簡稱存在返回true</returns>
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
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查Perso英文簡稱, CheckEnPersoExist報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }
        /// <summary>
        /// 檢查作業日期當日是否已經有特殊代制項目
        /// </summary>

        /// <param name="ImportEDate">作業日期</param>
        /// <returns>有返回true</returns>
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

                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查作業日期當日是否已經有特殊代制項目, CheckImportDate報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }

        }
        /// <summary>
        /// 取特殊代制項目訊息
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
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取特殊代制項目訊息, Get_SPECIAL_PROJECT報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 取所有特殊代制項目匯入訊息
        /// </summary>
        private void Get_SPECIAL_PERSO_PROJECT_IMPORT_ALL()
        {
            try
            {
                dtSPIA = dao.GetList(SEL_SPECIAL_PERSO_PROJECT_IMPORT_ALL).Tables[0];
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取所有特殊代制項目匯入訊息, Get_SPECIAL_PERSO_PROJECT_IMPORT_ALL報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 匯入檔案檢查
        /// </summary>

        /// <param name="strPath">檔案路徑名稱</param>
        /// <param name="dtPeso">Peso廠信息</param>
        /// <param name="dtSP">特殊代制項目訊息</param>
        /// <param name="dtSPIA">特殊代制項目匯入訊息</param>
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
                            strErr += "文檔第一行中的廠商編號代表的Perso廠和文檔名中的Perso廠簡稱代表的Perso廠不一致\n";
                        }
                        count++;
                    }
                    else
                    {
                        if (!StringUtil.IsEmpty(strReadLine))
                        {
                            if (StringUtil.GetByteLength(strReadLine) != 23)//列數量檢查
                            {
                                strErr += "第" + count.ToString() + "格式錯誤。\n";
                                //throw new AlertException("第" + count.ToString() + "行列數不正確。");
                            }
                            else
                            {
                                // 分割字符串
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
                                    strErr += "第" + count.ToString() + "的特殊代制項目編號不存在。\n";
                                }
                                else
                                {
                                    PersoProject_RID = Convert.ToInt32(FRID);
                                }
                                if (dtSPIA.Rows.Count > 0)
                                {
                                    if (isExistSpecialCode(Convert.ToInt32(dtPeso.Rows[0]["RID"].ToString()), Date, Convert.ToInt32(FRID)))
                                    {
                                        strErr += "第" + count.ToString() + "特殊代制項目匯入訊息已經存在，不能重復匯入。\n";
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
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("匯入檔案檢查, ImportCheck報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
            finally
            {
                sr.Close();
            }
        }
        /// <summary>
        ///初始化DT
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
        ///檢查特殊代制項目編號是否正確
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
        ///檢查特殊代制項目是否已匯入
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
