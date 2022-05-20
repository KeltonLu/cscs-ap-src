//*****************************************
//*  作    者：陳永銘
//*  功能說明：次月下市預測表匯入
//*  創建日期：2021-03-12
//*  修改日期：
//*  修改記錄：
//*****************************************

using System;
using System.Collections.Generic;
using System.Configuration;
using System.Collections;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using CIMSBatch.FTP;
using CIMSBatch.Public;
using System.Linq;

namespace CIMSBatch.Business
{
    class InOut006BL : BaseLogic
    {
        #region SQL語句定義
        const string SEL_FILE_NAME = "SELECT File_Name FROM IMPORT_PROJECT WHERE RST = 'A' AND Type = '2'";
        const string SEL_CARDTYPE = "SELECT * FROM CARD_TYPE WHERE RST = 'A' ";
        const string SEL_IMPORT_PROJECT_RR05 = "SELECT Affinity FROM IMPORT_PROJECT_RR05 WHERE Is_Import = 'Y' ";
        const string DEL_PERSO_FORE_CHANGE_CARD_RR05 = "DELETE FROM PERSO_FORE_CHANGE_CARD_RR05 WHERE RST = 'A' AND Fore_RID IN (SELECT RID FROM FORE_CHANGE_CARD_RR05 WHERE RST = 'A' AND Delist_Date = @delist_date AND Type = @type AND Affinity = @affinity AND Photo = @photo AND Disney_Code = @disney_code) ";
        const string DEL_FORE_CHANGE_CARD_RR05 = "DELETE FROM FORE_CHANGE_CARD_RR05 WHERE RST = 'A' AND Delist_Date = @delist_date AND Type = @type AND Affinity = @affinity AND Photo = @photo AND Disney_Code = @disney_code ";
        const string SEL_FORE_CHANGE_CARD_RR05 = "SELECT * FROM FORE_CHANGE_CARD_RR05 WHERE RST = 'A' AND Delist_Date = @delist_date AND Type = @type AND Affinity = @affinity AND Photo = @photo AND Disney_Code = @disney_code ";
        const string SEL_CARDTYPE_PERSO = "SELECT PC.*,CT.TYPE,CT.AFFINITY,CT.PHOTO FROM PERSO_CARDTYPE PC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND PC.CardType_RID = CT.RID WHERE PC.RST = 'A' AND PC.CardType_RID = @cardtype_rid ORDER BY PC.Base_Special DESC,PC.Priority ASC ";

        #endregion
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        public string strErr;
        #region 下載次月下市預測檔
        public ArrayList MonthDelistCard()
        {
            #region Attribute

            ArrayList FileNameList = new ArrayList();

            try
            {
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.MONTHREPLACE);
                string ftpPath = ConfigurationManager.AppSettings["FTPRemoteNextMonthDelistCard"]; //ftp檔案目錄配置檔信息
                string locPath = ConfigurationManager.AppSettings["NextMonthDelistCardForecastFilesPath"]; //local檔案目錄配置檔信息
                string FolderName = ""; //本地目錄
                string remFileName; //FTP檔名
                string locFileName; //本地存儲檔名
                string[] fileList;
                string[] namelist;
                bool returnFlag;

                fileList = GetMonthDelistCard();
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
            catch (Exception ex)
            {
                LogFactory.Write("匯入次月下市預測檔下載FTP檔案MonthDelistCard方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }


            return FileNameList;
        }
        #endregion
        /// <summary>
        /// 取得次月下市預測檔案名稱列表
        /// </summary>
        /// <returns>string []</returns>
        private string[] GetMonthDelistCard()
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
                LogFactory.Write("取得次月下市檔名錯誤:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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
            DataSet dsIMPORTPROJECTRR05 = null;

            string NowTime = Convert.ToDateTime(DateTime.Now).ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo).Replace("1900/01/01", "").ToString();

            DataTable dtblFileImp = new DataTable();
            dtblFileImp.Columns.Add("Old_Affinity_Code");
            dtblFileImp.Columns.Add("Name");
            dtblFileImp.Columns.Add("Disney_Code");
            dtblFileImp.Columns.Add("Photo_Code");
            dtblFileImp.Columns.Add("Type_Code");
            dtblFileImp.Columns.Add("Affinity_Code");
            dtblFileImp.Columns.Add("Card_SEQ");
            dtblFileImp.Columns.Add("Unit_Code");
            dtblFileImp.Columns.Add("Delist_Number");
            dtblFileImp.Columns.Add("Delist_Date");

            try
            {
                dsCARDTYPE = dao.GetList(SEL_CARDTYPE);
                dsIMPORTPROJECTRR05 = dao.GetList(SEL_IMPORT_PROJECT_RR05);
                string[] strLine;
                string[] strIMPORTPROJECTRR05;
                string strReadLine = "";
                int count = 0;
                string strErr = "";
                string year_month = NowTime.Split('/')[0] + NowTime.Split('/')[1];

                strIMPORTPROJECTRR05 = new string[dsIMPORTPROJECTRR05.Tables[0].Rows.Count];
                for (int i = 0; i < dsIMPORTPROJECTRR05.Tables[0].Rows.Count; i++)
                {
                    strIMPORTPROJECTRR05[i] = dsIMPORTPROJECTRR05.Tables[0].Rows[i]["Affinity"].ToString();
                }

                while ((strReadLine = sr.ReadLine()) != null)
                {
                    count++;
                    if (StringUtil.IsEmpty(strReadLine))
                        continue;

                    string Pattern = @"\w+";
                    MatchCollection Matches = Regex.Matches(strReadLine.Replace(",", ""), Pattern, RegexOptions.IgnoreCase);

                    if (Matches.Count != 9)
                        continue;

                    if (!strIMPORTPROJECTRR05.Contains(Matches[0].ToString()))
                        continue;

                    DataRow dr = dtblFileImp.NewRow();
                    strLine = new string[Matches.Count];
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

                    if (dtblFileImp.Select("Type_Code='" + dr["Type_Code"].ToString() + "' AND Affinity_Code = '" + dr["Affinity_Code"].ToString() + "' AND Photo_Code = '" + dr["Photo_Code"].ToString() + "' AND Disney_Code = '" + dr["Disney_Code"].ToString() + "' ").Length > 0)
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
        ///匯入次月下市預測檔
        /// </summary>
        public string In(DataSet dsFileImp)
        {
            FORE_CHANGE_CARD_RR05 fccrModel = null;
            DataSet dsCARDTYPE = null;

            try
            {
                dao.OpenConnection();

                foreach (DataRow dr_date in dsFileImp.Tables[0].Rows)
                {
                    fccrModel = new FORE_CHANGE_CARD_RR05();

                    Del_PERSO_FORE_CHANGE_CARD_RR05(dr_date["Type_Code"].ToString(), dr_date["Affinity_Code"].ToString(), dr_date["Photo_Code"].ToString(), dr_date["Delist_Date"].ToString(), dr_date["Disney_Code"].ToString());
                    int intRID = 0;

                    if (dao.GetList(SEL_FORE_CHANGE_CARD_RR05, dirValues).Tables[0].Rows.Count > 0)
                    {
                        fccrModel = dao.GetModel<FORE_CHANGE_CARD_RR05>(SEL_FORE_CHANGE_CARD_RR05, dirValues);
                        fccrModel.Number = dr_date["Delist_Number"].ToString();
                        fccrModel.IsMonth = "1";
                        dao.Update<FORE_CHANGE_CARD_RR05>(fccrModel, "RID");
                        intRID = fccrModel.RID;
                    }
                    else
                    {
                        //添加次月下市預測訊息。Dao.add(),并取出新添加記錄的RID
                        fccrModel.Delist_Date = dr_date["Delist_Date"].ToString();
                        fccrModel.Old_Affinity = dr_date["Old_Affinity_Code"].ToString();
                        fccrModel.Name = dr_date["Name"].ToString();
                        fccrModel.Disney_Code = dr_date["Disney_Code"].ToString();
                        fccrModel.Type = dr_date["Type_Code"].ToString();
                        fccrModel.Photo = dr_date["Photo_Code"].ToString();
                        fccrModel.Affinity = dr_date["Affinity_Code"].ToString();
                        fccrModel.Card_SEQ = dr_date["Card_SEQ"].ToString();
                        fccrModel.Unit_Code = dr_date["Unit_Code"].ToString();
                        fccrModel.Number = dr_date["Delist_Number"].ToString();
                        fccrModel.IsMonth = "1";
                        fccrModel.IsYear = "2";
                        intRID = Convert.ToInt32(dao.AddAndGetID<FORE_CHANGE_CARD_RR05>(fccrModel, "RID"));
                    }

                    dsCARDTYPE = dao.GetList(SEL_CARDTYPE + "AND Type = @type AND Affinity = @affinity AND Photo = @photo", dirValues);
                    foreach (DataRow dr1 in dsCARDTYPE.Tables[0].Rows)
                    {
                        SplitToPerso(Convert.ToInt32(dr1["RID"]), intRID, Convert.ToInt64(dr_date["Delist_Number"].ToString()), dr_date["Delist_Date"].ToString(), fccrModel);
                    }
                }
                dao.Commit();
                return "";
            }
            catch (Exception ex)
            {
                dao.Rollback();
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("匯入年度次月下市預測檔, In報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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
        public void SplitToPerso(int Space_RID, int intRID, long DtDate_Number, string StartDate, FORE_CHANGE_CARD_RR05 Table_Vaule)
        {
            PERSO_FORE_CHANGE_CARD_RR05 pfccrModel = new PERSO_FORE_CHANGE_CARD_RR05();
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
                        pfccrModel.Delist_Date = StartDate;
                        pfccrModel.Type = dsCARDTYPE_PERSO.Tables[0].Rows[0]["Type"].ToString();
                        pfccrModel.Affinity = dsCARDTYPE_PERSO.Tables[0].Rows[0]["Affinity"].ToString();
                        pfccrModel.Photo = dsCARDTYPE_PERSO.Tables[0].Rows[0]["Photo"].ToString();
                        pfccrModel.Perso_Factory_RID = Convert.ToInt32(dsCARDTYPE_PERSO.Tables[0].Rows[0]["Factory_RID"].ToString());
                        pfccrModel.Number = DtDate_Number.ToString();
                        pfccrModel.Fore_RID = intRID;
                        pfccrModel.Old_Affinity = Table_Vaule.Old_Affinity;
                        pfccrModel.Name = Table_Vaule.Name;
                        pfccrModel.Disney_Code = Table_Vaule.Disney_Code;
                        pfccrModel.Card_SEQ = Table_Vaule.Card_SEQ;
                        pfccrModel.Unit_Code = Table_Vaule.Unit_Code;
                        dao.Add<PERSO_FORE_CHANGE_CARD_RR05>(pfccrModel, "RID");
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
                                    pfccrModel.Delist_Date = StartDate;
                                    pfccrModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString();
                                    pfccrModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString();
                                    pfccrModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString();
                                    pfccrModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                    pfccrModel.Number = Convert.ToInt64(Math.Floor(DtDate_Number * (Convert.ToDouble(drCARDTYPE_PERSO[int1]["Value"]) / 100))).ToString();
                                    pfccrModel.Fore_RID = intRID;
                                    pfccrModel.Old_Affinity = Table_Vaule.Old_Affinity;
                                    pfccrModel.Name = Table_Vaule.Name;
                                    pfccrModel.Disney_Code = Table_Vaule.Disney_Code;
                                    pfccrModel.Card_SEQ = Table_Vaule.Card_SEQ;
                                    pfccrModel.Unit_Code = Table_Vaule.Unit_Code;
                                    dao.Add<PERSO_FORE_CHANGE_CARD_RR05>(pfccrModel, "RID");
                                }
                                else
                                {
                                    pfccrModel.Delist_Date = StartDate;
                                    pfccrModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString();
                                    pfccrModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString();
                                    pfccrModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString();
                                    pfccrModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                    pfccrModel.Number = (DtDate_Number - intNumber).ToString();
                                    pfccrModel.Fore_RID = intRID;
                                    pfccrModel.Old_Affinity = Table_Vaule.Old_Affinity;
                                    pfccrModel.Name = Table_Vaule.Name;
                                    pfccrModel.Disney_Code = Table_Vaule.Disney_Code;
                                    pfccrModel.Card_SEQ = Table_Vaule.Card_SEQ;
                                    pfccrModel.Unit_Code = Table_Vaule.Unit_Code;
                                    dao.Add<PERSO_FORE_CHANGE_CARD_RR05>(pfccrModel, "RID");
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
                                        pfccrModel.Delist_Date = StartDate;
                                        pfccrModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString(); ;
                                        pfccrModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString(); ;
                                        pfccrModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString(); ;
                                        pfccrModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                        pfccrModel.Number = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Value"]).ToString();
                                        pfccrModel.Fore_RID = intRID;
                                        pfccrModel.Old_Affinity = Table_Vaule.Old_Affinity;
                                        pfccrModel.Name = Table_Vaule.Name;
                                        pfccrModel.Disney_Code = Table_Vaule.Disney_Code;
                                        pfccrModel.Card_SEQ = Table_Vaule.Card_SEQ;
                                        pfccrModel.Unit_Code = Table_Vaule.Unit_Code;
                                        dao.Add<PERSO_FORE_CHANGE_CARD_RR05>(pfccrModel, "RID");
                                    }
                                    else
                                    {
                                        pfccrModel.Delist_Date = StartDate;
                                        pfccrModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString(); ;
                                        pfccrModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString(); ;
                                        pfccrModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString(); ;
                                        pfccrModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                        pfccrModel.Number = (DtDate_Number - intNumber).ToString();
                                        pfccrModel.Fore_RID = intRID;
                                        pfccrModel.Old_Affinity = Table_Vaule.Old_Affinity;
                                        pfccrModel.Name = Table_Vaule.Name;
                                        pfccrModel.Disney_Code = Table_Vaule.Disney_Code;
                                        pfccrModel.Card_SEQ = Table_Vaule.Card_SEQ;
                                        pfccrModel.Unit_Code = Table_Vaule.Unit_Code;
                                        dao.Add<PERSO_FORE_CHANGE_CARD_RR05>(pfccrModel, "RID");
                                        break;
                                    }
                                }
                                else
                                {
                                    pfccrModel.Delist_Date = StartDate;
                                    pfccrModel.Type = drCARDTYPE_PERSO[int1]["Type"].ToString(); ;
                                    pfccrModel.Affinity = drCARDTYPE_PERSO[int1]["Affinity"].ToString(); ;
                                    pfccrModel.Photo = drCARDTYPE_PERSO[int1]["Photo"].ToString(); ;
                                    pfccrModel.Perso_Factory_RID = Convert.ToInt32(drCARDTYPE_PERSO[int1]["Factory_RID"].ToString());
                                    pfccrModel.Number = (DtDate_Number - intNumber).ToString();
                                    pfccrModel.Fore_RID = intRID;
                                    pfccrModel.Old_Affinity = Table_Vaule.Old_Affinity;
                                    pfccrModel.Name = Table_Vaule.Name;
                                    pfccrModel.Disney_Code = Table_Vaule.Disney_Code;
                                    pfccrModel.Card_SEQ = Table_Vaule.Card_SEQ;
                                    pfccrModel.Unit_Code = Table_Vaule.Unit_Code;
                                    dao.Add<PERSO_FORE_CHANGE_CARD_RR05>(pfccrModel, "RID");
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
        /// 刪除PERSO_FORE_CHANGE_CARD_RR05檔中對應的記錄
        /// </summary>
        private void Del_PERSO_FORE_CHANGE_CARD_RR05(string type, string affinity, string photo, string delistdate, string disneycode)
        {
            dirValues.Clear();
            dirValues.Add("type", type);
            dirValues.Add("affinity", affinity);
            dirValues.Add("photo", photo);
            dirValues.Add("delist_date", delistdate);
            dirValues.Add("disney_code", disneycode);
            dao.ExecuteNonQuery(DEL_PERSO_FORE_CHANGE_CARD_RR05, dirValues);
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
                case 4:
                case 7:
                    Pattern = @"^\d{2}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為2位數字;\n";
                    }
                    break;
                case 5:
                    Pattern = @"^\d{3}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為3位數字;\n";
                    }

                    break;
                case 1:
                case 6:
                    Pattern = @"^\d{4}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為4位數字;\n";
                    }
                    break;
                case 8:
                    Pattern = @"^\d{5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為5位數字;\n";
                    }
                    break;
                case 9:
                    Pattern = @"^\d{7}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為7位數字;\n";
                    }
                    break;
                case 3:
                    Pattern = @"^[a-zA-Z0-9]*$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0 || strColumn.Length != 2)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為2位英數;\n";
                    }
                    break;
            }

            return strErr;
        }
    }
}
