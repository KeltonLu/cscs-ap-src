//*****************************************
//*  作    者：GaoAi
//*  功能說明：次月換卡預測檔匯入
//*  創建日期：2008-11-29
//*  修改日期：
//*  修改記錄：
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
    class InOut005BL : BaseLogic
    {
        #region SQL語句定義
        const string SEL_FILE_NAME = "SELECT File_Name FROM IMPORT_PROJECT WHERE RST = 'A' AND Type = '2'";
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
        #region 下載次月換卡預測檔
        public ArrayList MonthReplaceCard()
        {
            #region Attribute

            ArrayList FileNameList = new ArrayList();

            try  // 加try catch add by judy 2018/03/28
            {
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.MONTHREPLACE);
                string ftpPath = ConfigurationManager.AppSettings["FTPRemoteMonthReplaceCard"]; //ftp檔案目錄配置檔信息
                string locPath = ConfigurationManager.AppSettings["MonthReplaceCardForecastFilesPath"]; //local檔案目錄配置檔信息
                string FolderName = ""; //本地目錄
                string remFileName; //FTP檔名
                string locFileName; //本地存儲檔名
                string[] fileList;
                string[] namelist;
                bool returnFlag;

                GetCardType(dtcard);

                fileList = GetMonthReplaceCard();
                if (fileList != null)
                {
                    foreach (string FileName in fileList)
                    {

                        remFileName = "";

                        locFileName = FileName + DateTime.Now.ToString("yyyyMM") + ".txt";
                        if (ftp.GetFileList(ftpPath, remFileName) == null) //檢查FTP是否有對應檔案
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

                                        string[] FList = new string[2];
                                        FList[0] = FolderName;
                                        FList[1] = locFileName;
                                        FileNameList.Add(FList);
                                        // Legend 2017/11/28 將此處刪檔注釋做UAT測試, 上線是再解開 todo
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
                LogFactory.Write("匯入次月換卡預測檔下載FTP檔案MonthReplaceCard方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

           
            return FileNameList;
        }
            #endregion
        /// <summary>
        /// 取得次月換卡預測檔案名稱列表
        /// </summary>
        /// <returns>string []</returns>
        private string[] GetMonthReplaceCard()
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
            catch (Exception ex)
            {
                LogFactory.Write("取得次月換卡檔名錯誤:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
        }

        #endregion
        /// <summary>
        ///下載資料檢核
        /// </summary>
        public DataSet DetailCheck(string strPath)
        {
            DataSet dtsReturn = new DataSet();
            #region 驗證文件
            StreamReader sr = null;
            sr = new StreamReader(strPath, System.Text.Encoding.Default);
            DataSet dsCARDTYPE = null;

            string NowTime = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo).Replace("1900/01/01", "").ToString();

            DataTable dtblFileImp = new DataTable();
            dtblFileImp.Columns.Add("Photo_Code");
            dtblFileImp.Columns.Add("Type_Code");
            dtblFileImp.Columns.Add("Affinity_Code");
            dtblFileImp.Columns.Add("Change_Number");
            dtblFileImp.Columns.Add("Change_Date");
            try
            {
                dsCARDTYPE = dao.GetList(SEL_CARDTYPE);
                string[] strLine;
                string strReadLine = "";
                int count = 0;
                int j = 0;
                string strErr = "";
                string year_month = "";

                while ((strReadLine = sr.ReadLine()) != null)
                {
                    count++;
                    if (StringUtil.IsEmpty(strReadLine))
                        continue;

                    if (strReadLine.Contains("換卡月份"))
                    {
                        strLine = strReadLine.Split('：');
                        strLine = strLine[1].Trim().Split(' ');

                        if (strLine[0].Trim().Length != 4)
                            throw new Exception("換卡月份格式錯誤");
                        if (CheckFileColumn(strLine[0].Trim(), 3, 0) != "")
                            throw new Exception("換卡月份格式錯誤");

                        if (int.Parse(strLine[0].Trim().Substring(2)) > 12)
                            throw new Exception("換卡月份格式錯誤");

                        year_month = "20" + strLine[0].Trim().ToString();

                        continue;
                    }
                    else if (strReadLine.Contains("PHOTO TYPE"))
                    {
                        continue;
                    }
                    else if (strReadLine.Contains("======"))
                    {
                        continue;
                    }
                    else if (strReadLine.Contains("程式名稱") || strReadLine.Contains("報表編號") || strReadLine.Contains("TOTAL"))
                    {
                        continue;
                    }
                    else
                    {
                        string Pattern = @"\w+";
                        MatchCollection Matches = Regex.Matches(strReadLine.Replace(",", ""), Pattern, RegexOptions.IgnoreCase);

                        if (Matches.Count != 4)
                            continue;

                        DataRow dr = dtblFileImp.NewRow();
                        strLine = new string[4];
                        for (int i = 0; i < Matches.Count; i++)
                        {
                            strLine[i] = Matches[i].ToString();
                        }


                        for (int i = 0; i < strLine.Length; i++)
                        {
                            int num = i + 1;
                            if (StringUtil.IsEmpty(strLine[i]))
                                strErr += "第" + count.ToString() + "行第" + num.ToString() + "列為空;\n";
                            else
                                strErr += CheckFileColumn(strLine[i], num, count);
                            dr[i] = strLine[i];
                        }

                        dr[strLine.Length] = year_month;

                        if (dtblFileImp.Select("Type_Code='" + dr["Type_Code"].ToString() + "' AND Affinity_Code = '" + dr["Affinity_Code"].ToString() + "' AND Photo_Code = '" + dr["Photo_Code"].ToString() + "' ").Length > 0)
                        {
                            strErr += "第" + count.ToString() + "行 " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                            strErr += "對應的卡種已經存在,不能重複匯入!\n";
                        }

                        dtblFileImp.Rows.Add(dr);

                        if (dsCARDTYPE.Tables[0].Select("TYPE='" + dr["Type_Code"].ToString() + "' AND AFFINITY = '" + dr["Affinity_Code"].ToString() + "' AND PHOTO = '" + dr["Photo_Code"].ToString() + "' ").Length == 0)
                        {
                            strErr += "第" + count.ToString() + "行 " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                            strErr += "對應的卡種不存在!\n";
                        }
                        else
                        {

                            if (dsCARDTYPE.Tables[0].Select("TYPE='" + dr["Type_Code"].ToString() + "' AND AFFINITY = '" + dr["Affinity_Code"].ToString() + "' AND PHOTO = '" + dr["Photo_Code"].ToString() + "' AND Is_Using = 'N' ").Length >= 1)
                            {
                                strErr += "第" + count.ToString() + "行 " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                                strErr += "對應的卡種已停用!\n";
                            }


                            if (dsCARDTYPE.Tables[0].Select("Begin_Time<='" + NowTime + "' AND (End_Time='1900-01-01' or End_Time >='" + NowTime + "') AND TYPE='" + dr["Type_Code"].ToString() + "' AND AFFINITY = '" + dr["Affinity_Code"].ToString() + "' AND PHOTO = '" + dr["Photo_Code"].ToString() + "' ").Length == 0)
                            {
                                strErr += "第" + count.ToString() + "行 " + dr["Photo_Code"].ToString() + "-" + dr["Affinity_Code"].ToString() + "-" + dr["Type_Code"].ToString();
                                strErr += "第" + count.ToString() + "行對應的卡種不在有效期內!;\n";
                            }

                        }

                    }
                }

                dtsReturn.Tables.Add(dtblFileImp);

                if (!StringUtil.IsEmpty(strErr))
                {
                    throw new Exception(strErr);
                }

            }
            catch (Exception ex)
            {
                string[] arg = new string[1];
                arg[0] = ex.Message;
                Warning.SetWarning(GlobalString.WarningType.MonthChangeCardForeCast, arg);
            }
            finally
            {
                sr.Close();
            }
            #endregion
            return dtsReturn;

        }
        /// <summary>
        ///匯入年度換卡預測檔
        /// </summary>
        public string In(DataSet dsFileImp)
        {
            FORE_CHANGE_CARD fccModel = null;
            DataSet dsCARDTYPE = null;

            try
            {
                dao.OpenConnection();

                foreach (DataRow dr_date in dsFileImp.Tables[0].Rows)
                {
                    fccModel = new FORE_CHANGE_CARD();

                    Del_PERSO_FORE_CHANGE_CARD(dr_date["Type_Code"].ToString(), dr_date["Affinity_Code"].ToString(), dr_date["Photo_Code"].ToString(), dr_date["Change_Date"].ToString());
                    //Del_FORE_CHANGE_CARD(dr_date["Type_Code"].ToString(), dr_date["Affinity_Code"].ToString(), dr_date["Photo_Code"].ToString(), dr_date["Change_Date"].ToString());
                    int intRID = 0;

                    if (dao.GetList(SEL_FORE_CHANGE_CARD, dirValues).Tables[0].Rows.Count > 0)
                    {
                        fccModel = dao.GetModel<FORE_CHANGE_CARD>(SEL_FORE_CHANGE_CARD, dirValues);
                        fccModel.Number = Convert.ToInt64(dr_date["Change_Number"]);
                        fccModel.IsMonth = "1";
                        dao.Update<FORE_CHANGE_CARD>(fccModel, "RID");
                        intRID = fccModel.RID;
                    }
                    else
                    {
                        //添加次月換卡預測訊息。Dao.add(),并取出新添加記錄的RID
                        fccModel.Change_Date = dr_date["Change_Date"].ToString();
                        fccModel.Type = dr_date["Type_Code"].ToString();
                        fccModel.Affinity = dr_date["Affinity_Code"].ToString();
                        fccModel.Photo = dr_date["Photo_Code"].ToString();
                        fccModel.Number = Convert.ToInt64(dr_date["Change_Number"]);
                        fccModel.IsMonth = "1";
                        fccModel.IsYear = "2";
                        intRID = Convert.ToInt32(dao.AddAndGetID<FORE_CHANGE_CARD>(fccModel, "RID"));
                    }
                 
                    //添加次月換卡預測訊息。Dao.add(),并取出新添加記錄的RID
                    //fccModel.Change_Date = dr_date["Change_Date"].ToString();
                    //fccModel.Type = dr_date["Type_Code"].ToString();
                    //fccModel.Affinity = dr_date["Affinity_Code"].ToString();
                    //fccModel.Photo = dr_date["Photo_Code"].ToString();
                    //fccModel.Number = Convert.ToInt64(dr_date["Change_Number"]);
                    //int intRID = Convert.ToInt32(dao.AddAndGetID<FORE_CHANGE_CARD>(fccModel, "RID"));
                    dsCARDTYPE = dao.GetList(SEL_CARDTYPE + "AND Type = @type AND Affinity = @affinity AND Photo = @photo", dirValues);
                    foreach (DataRow dr1 in dsCARDTYPE.Tables[0].Rows)
                    {
                        SplitToPerso(Convert.ToInt32(dr1["RID"]), intRID, Convert.ToInt64(dr_date["Change_Number"]), dr_date["Change_Date"].ToString());
                    }
                }
                dao.Commit();
                return "";
            }
            catch (Exception ex)
            {
                dao.Rollback();
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("匯入年度換卡預測檔, In報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                BatchBL Bbl = new BatchBL();
                Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
                return "erro";
            }
            finally
            {
                dao.CloseConnection();
            }

        }
        /// <summary>
        /// 按Perso廠分配原則來拆分
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
                        // 取特殊分配訊息
                        DataRow[] drCARDTYPE_PERSO = dsCARDTYPE_PERSO.Tables[0].Select("Base_Special = '2'");
                        // 按比率分配
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
                        // 按數量分配
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
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("按Perso廠分配原則來拆分, SplitToPerso報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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
            dt.Columns.Add(new DataColumn("Number", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("ChangeDate", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Change_Space_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Replace_Space_RID", System.Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("RID", System.Type.GetType("System.String")));
            return dt;
        }

        /// <summary>
        /// 檢查Column是否為數字
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
                strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為數字;";
            }
            return strErr;
        }
        /// <summary>
        /// 驗證匯入字段是否滿足格式
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
                strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式錯誤;";
            }
            return strErr;
        }
        /// <summary>
        /// 驗證匯入字段是否滿足格式
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
        /// 取得所有的卡種信息
        /// </summary>
        private void GetCardType(DataTable dt)
        {
            dt = new DataTable();

            dt = dao.GetList(SEL_CARD_TYPE).Tables[0];
        }
        /// <summary>
        /// 刪除PERSO_FORE_CHANGE_CARD檔中對應的記錄
        /// </summary>
        private void Del_PERSO_FORE_CHANGE_CARD(string type, string affin, string photo, string changedate)
        {
            dirValues.Clear();
            dirValues.Add("type", type);
            dirValues.Add("affinity", affin);
            dirValues.Add("photo", photo);
            dirValues.Add("change_Date", changedate);
            dao.ExecuteNonQuery(DEL_PERSO_FORE_CHANGE_CARD, dirValues);
        }
        /// /// <summary>
        /// 刪除FORE_CHANGE_CARD檔中對應的記錄
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
                    Pattern = @"^\d{2}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + (count - 1).ToString() + "行第" + num.ToString() + "列格式必須為2位數字;\n";
                    }
                    break;
                case 2:
                    Pattern = @"^\d{3}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + (count - 1).ToString() + "行第" + num.ToString() + "列格式必須為3位數字;\n";
                    }

                    break;
                case 3:
                    Pattern = @"^\d{4}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + (count - 1).ToString() + "行第" + num.ToString() + "列格式必須為4位數字;\n";
                    }
                    break;
                case 4:
                    Pattern = @"^\d+$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為數字;\n";
                    }
                    break;
            }

            return strErr;
        }
    }
}
