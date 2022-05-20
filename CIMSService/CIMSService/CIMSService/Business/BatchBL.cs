//*****************************************
//*  作    者：
//*  功能說明：
//*  創建日期：
//*  修改日期：2021-03-12
//*  修改記錄：新增次月下市預測表匯入 陳永銘
//*  修改日期：2021-05-18
//*  修改記錄：字串修正 陳永銘
//*****************************************
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using CIMSBatch;
using CIMSBatch.Business;
using CIMSBatch.Mail;
using CIMSBatch.Model;
using CIMSBatch.Public;
using CIMSClass.Business;

namespace CIMSBatch.Business
{
    class BatchBL : BaseLogic
    {
        #region SQL定義
        const string SEL_BATCH_MANAGE = "select Status from BATCH_MANAGE where RID = @RID";
        string Update_BATCH_MANAGE = "update BATCH_MANAGE set Status = @Status,RUU='BatchBL',RUT='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' where RID=@RID";
        const string SEL_WORK_DATE = "select count (Date_Time) from WORK_DATE where Date_Time = @Date_Time and Is_WorkDay='Y'";
        const string SEL_IMPORT_HISTORY = "select count (rid) from import_history where file_name = @file_name";
        const string SEL_SUBTOTAL_ALERT = "SELECT WC.System_Show, WC.Mail_Show,WU.Warning_RID,WU.UserID,WC.Warning_Content,U.Email FROM WARNING_CONFIGURATION WC INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID LEFT JOIN USERS U ON WU.UserID = U.UserID WHERE WC.RST = 'A' AND WC.RID = 2";
        const string SEL_FACTORY_CHANGE_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID"
                         + " LEFT JOIN USERS U ON  WU.UserID = U.UserID"
                         + "WHERE WC.RST = 'A' AND WC.RID =";
        const string SEL_YEAR_REPLACE_CARD_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID"
                         + "LEFT JOIN USERS U ON U.RST = 'A' AND WU.UserID = U.RID"
                         + "WHERE WC.RST = 'A' AND WC.RID =";
        const string SEL_NEXT_MONTH_REPLACE_CARD_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC"
                         + "INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID"
                         + " LEFT JOIN USERS U ON U.RST = 'A' AND WU.UserID = U.RID"
                         + "WHERE WC.RST = 'A' AND WC.RID =";
        const string SEL_DAY_SURPLUS_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC"
                         + " INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID"
                         + " LEFT JOIN USERS U ON U.RST = 'A' AND WU.UserID = U.RID"
                         + " WHERE WC.RST = 'A' AND WC.RID =";
        const string SEL_SURPLUS_CHECK = "SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS"
                         + "WHERE AST= 'A' ORDER BY Stock_Date DESC";
        #endregion


        #region 批次狀態表操作
        /// <summary>
        /// 檢查Web程式是否在做批次動作
        /// </summary>
        /// <returns>Web程式進行中返回true</returns>
        private bool CheckBatchStatus(int BatchIndex)
        {
            Dictionary<string, object> dirValues = new Dictionary<string, object>();
            dirValues.Add("RID", BatchIndex);
            try
            {
                DataSet ds = dao.GetList(SEL_BATCH_MANAGE, dirValues);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0]["Status"].ToString() == GlobalString.BatchStatus.Run)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/12 添加方法名字
                LogFactory.Write("檢查批次狀態CheckBatchStatus錯誤:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        /// <summary>
        /// 更新批次表狀態
        /// </summary>
        /// <param name="BatchIndex">批次序號</param>
        /// <param name="Status">狀態</param>
        /// <returns></returns>
        private bool UpdateBatchStatus(int BatchIndex, string Status)
        {
            Dictionary<string, object> dirValues = new Dictionary<string, object>();
            dirValues.Add("RID", BatchIndex);
            dirValues.Add("Status", Status);
            try
            {
                int returnValue = dao.ExecuteNonQuery(Update_BATCH_MANAGE, dirValues);
                return true;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/12 添加方法名字
                LogFactory.Write("更新批次狀態UpdateBatchStatus錯誤:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        #endregion

        /// <summary>
        /// 測試用方法，ADD BY 郭佳！
        /// </summary>
        /// <param name="testid"></param>
        public void TestBatch(int testid)
        {
            ArrayList al;
            switch (testid)
            {
                case 61:
                    InOut001BL bl61 = new InOut001BL();
                    al = bl61.DownloadSubtotal();
                    Batch61(al);
                    break;
                case 62:
                    InOut002BL bl62 = new InOut002BL();
                    al = bl62.DownLoadModify("TNP");
                    Batch62(al);
                    break;

                case 63:
                    Batch63();
                    break;
                case 64:
                    InOut004BL bl64 = new InOut004BL();
                    al = bl64.YearReplaceCard();
                    Batch64(al);
                    break;
                case 65:
                    InOut005BL bl65 = new InOut005BL();
                    al = bl65.MonthReplaceCard();
                    Batch65(al);
                    break;
                // 2021-03-12 新增次月下市預測表匯入 陳永銘
                case 66:
                    InOut006BL bl66 = new InOut006BL();
                    al = bl66.MonthDelistCard();
                    Batch66(al);
                    break;
                case 410:
                    Depository010BL bl410 = new Depository010BL();
                    al = bl410.Download("TNP");
                    Batch410(al);
                    break;
                case 524:
                    Finance0024BL bl524 = new Finance0024BL();
                    al = bl524.DownLoadModify("Wits");
                    Batch524(al);
                    break;
                case 13:
                    Depository013BL bl013 = new Depository013BL();
                    bl013.ComputeForeData();
                    //
                    bl013 = null;
                    GC.Collect();

                    break;

                case 14:
                    Depository014BL bl014 = new Depository014BL();
                    bl014.ComputeForeData();
                    break;
                case 10:
                    CardSplitToPersoBL blSplit = new CardSplitToPersoBL();
                    blSplit.SplitToPerso();
                    break;
                case 214:
                    this.ImportFactoryData("Test");
                    break;
                case 42:
                    LADPCheckManager.GetLDAPAuth();
                    break;
                case 67:
                    CIMSClass.Business.InOut007BL bl67 = new CIMSClass.Business.InOut007BL();
                    al = bl67.DownLoadModify("tnp");
                    Batch67(al);
                    break;
                case 68:
                    WarningBatch WB = new WarningBatch();
                    WB.BatchStart();
                    break;
                default:
                    break;
            }
        }


        /// <summary>
        /// 執行批次程式
        /// </summary>
        /// <param name="NowTime"></param>
        public void RunBatch(DateTime NowTime, int BatchRid)
        {
            try
            {
                if (!(CheckWorkDate(DateTime.Now))) //非工作日直接返回，不執行批次
                    return;

                LogFactory.Write("Enter RunBatch:" + BatchRid.ToString(), GlobalString.LogType.OpLogCategory);

                switch (BatchRid)
                {
                    //8:30___6.1/6.2/5.2.4/4.10/6.7
                    case GlobalString.BatchRID.BatchOne:
                        if (GlobalString.BatchRunFlag.BatchOne == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchOne)))
                        {
                            UpdateBatchStatus(GlobalString.BatchRID.BatchOne, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchOne = GlobalString.BatchStatus.Run;

                            LogFactory.Write("第一次匯入小計檔開始", GlobalString.LogType.OpLogCategory);
                            InOut001BL bl61 = new InOut001BL();
                            ArrayList al = bl61.DownloadSubtotal();
                            Batch61(al);
                            LogFactory.Write("第一次匯入小計檔結束", GlobalString.LogType.OpLogCategory);


                            ImportFactoryData("一");
                            //200908CR 匯入替換前版面廠商異動檔 add by 楊昆 2009/09/03 start
                            //ImportFactoryReplaceData("一");
                            //200908CR 匯入替換前版面廠商異動檔 add by 楊昆 2009/09/03 end
                        }
                        break;
                    //12:00___6.4/6.5
                    case GlobalString.BatchRID.BatchTwo:
                        if (GlobalString.BatchRunFlag.BatchTwo == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchTwo)))
                        {
                            UpdateBatchStatus(GlobalString.BatchRID.BatchTwo, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchTwo = GlobalString.BatchStatus.Run;

                            LogFactory.Write("匯入年度換卡預測檔開始", GlobalString.LogType.OpLogCategory);
                            InOut004BL bl64 = new InOut004BL();
                            ArrayList al = bl64.YearReplaceCard();
                            Batch64(al);
                            LogFactory.Write("匯入年度換卡預測檔結束", GlobalString.LogType.OpLogCategory);

                            bl64 = null;
                            GC.Collect();

                            //修改每月換卡預測檔的運行時間，但為了不修改網頁程序，不修改批次的批號。
                            UpdateBatchStatus(GlobalString.BatchRID.BatchThree, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchThree = GlobalString.BatchStatus.Run;

                            LogFactory.Write("匯入次月換卡預測檔開始", GlobalString.LogType.OpLogCategory);
                            InOut005BL bl65 = new InOut005BL();
                            al = bl65.MonthReplaceCard();
                            Batch65(al);
                            LogFactory.Write("匯入次月換卡預測檔結束", GlobalString.LogType.OpLogCategory);


                        }
                        break;
                    //13:00___6.1
                    case GlobalString.BatchRID.BatchThree:
                        if (GlobalString.BatchRunFlag.BatchThree == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchThree)))
                        {
                            //為了不修改WEB程序，利用一下批次一的批號！
                            UpdateBatchStatus(GlobalString.BatchRID.BatchOne, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchOne = GlobalString.BatchStatus.Run;

                            LogFactory.Write("第二次匯入小計檔開始", GlobalString.LogType.OpLogCategory);
                            InOut001BL bl61 = new InOut001BL();
                            ArrayList al = bl61.DownloadSubtotal();
                            Batch61(al);
                            LogFactory.Write("第二次匯入小計檔結束", GlobalString.LogType.OpLogCategory);

                        }
                        break;
                    //14:00___6.2/5.2.4/4.10/6.7
                    case GlobalString.BatchRID.BatchFour:
                        if (GlobalString.BatchRunFlag.BatchFour == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchFour)))
                        {
                            UpdateBatchStatus(GlobalString.BatchRID.BatchFour, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchFour = GlobalString.BatchStatus.Run;

                            ImportFactoryData("二");

                            //200908CR 匯入替換前版面廠商異動檔 add by 楊昆 2009/09/03 start
                            //ImportFactoryReplaceData("二");
                            //200908CR 匯入替換前版面廠商異動檔 add by 楊昆 2009/09/03 end

                        }
                        break;
                    //20:00___6.1/6.2/5.2.4/4.10/6.7
                    case GlobalString.BatchRID.BatchFive:
                        if (GlobalString.BatchRunFlag.BatchFive == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchFive)))
                        {
                            UpdateBatchStatus(GlobalString.BatchRID.BatchFive, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchFive = GlobalString.BatchStatus.Run;

                            LogFactory.Write("第三次匯入小計檔開始", GlobalString.LogType.OpLogCategory);
                            InOut001BL bl61 = new InOut001BL();
                            ArrayList al = bl61.DownloadSubtotal();
                            Batch61(al);
                            LogFactory.Write("第三次匯入小計檔結束", GlobalString.LogType.OpLogCategory);


                            ImportFactoryData("三");
                            //200908CR 匯入替換前版面廠商異動檔 add by 楊昆 2009/09/03 start
                            //ImportFactoryReplaceData("三");
                            //200908CR 匯入替換前版面廠商異動檔 add by 楊昆 2009/09/03 end

                        }
                        break;
                    //18:00___6.3/6.spilit/4.13/4.14
                    case GlobalString.BatchRID.BatchSix:
                        if (GlobalString.BatchRunFlag.BatchSix == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchSix)))
                        {

                            //TimeSpan tsFix = new TimeSpan(0, 15, 0);

                            LogFactory.Write("日結處理開始", GlobalString.LogType.OpLogCategory);
                            Batch63();
                            LogFactory.Write("日結處理結束", GlobalString.LogType.OpLogCategory);
                            GC.Collect();
                            //System.Threading.Thread.Sleep(tsFix);

                            LogFactory.Write("拆分預測檔開始", GlobalString.LogType.OpLogCategory);
                            CardSplitToPersoBL blSplit = new CardSplitToPersoBL();
                            blSplit.SplitToPerso();
                            LogFactory.Write("拆分預測檔結束", GlobalString.LogType.OpLogCategory);

                            blSplit = null;
                            GC.Collect();

                            //System.Threading.Thread.Sleep(tsFix);


                            //晚上的批次，4.13在第二天的凌晨12點開始跑！
                            //LogFactory.Write("開始時間：" + NowTime + "  4.13時間：" + NowTime.AddDays(1).Date , GlobalString.LogType.OpLogCategory);
                            if (NowTime.ToString("yyyyMMdd").Equals(DateTime.Now.ToString("yyyyMMdd")))
                            {
                                TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks);
                                TimeSpan ts2 = new TimeSpan(NowTime.AddDays(1).Date.Ticks);

                                TimeSpan ts = ts2.Subtract(ts1).Duration();

                                //LogFactory.Write("等待多久："+ ts.Minutes , GlobalString.LogType.OpLogCategory);
                                System.Threading.Thread.Sleep(ts);
                            }

                            LogFactory.Write("每月監控作業開始", GlobalString.LogType.OpLogCategory);
                            Depository013BL bl013 = new Depository013BL();
                            bl013.ComputeForeData();
                            LogFactory.Write("每月監控作業結束", GlobalString.LogType.OpLogCategory);

                            bl013 = null;
                            GC.Collect();

                            //System.Threading.Thread.Sleep(tsFix);


                            //4.14的批次，在每天早上凌晨兩點開始！
                            //if (DateTime.Now < DateTime.Now.Date.AddHours(2))
                            //{
                            //    TimeSpan ts1 = new TimeSpan(DateTime.Now.Ticks);
                            //    TimeSpan ts2 = new TimeSpan(DateTime.Now.Date.AddHours(2).Ticks);

                            //    TimeSpan ts = ts2.Subtract(ts1).Duration();

                            //    //TimeSpan ts = dt - DateTime.Now;
                            //    LogFactory.Write("等待多少分鐘：" + ts.Minutes, GlobalString.LogType.OpLogCategory);
                            //    System.Threading.Thread.Sleep(ts);
                            //}
                            //GC.Collect();
                            LogFactory.Write("每日監控作業開始", GlobalString.LogType.OpLogCategory);
                            Depository014BL bl014 = new Depository014BL();
                            bl014.ComputeForeData();
                            LogFactory.Write("每日監控作業結束", GlobalString.LogType.OpLogCategory);

                            bl014 = null;
                            GC.Collect();

                            if (System.Configuration.ConfigurationManager.AppSettings["TestType"] == "1")
                            {
                                LogFactory.Write("同步LDAP信息開始", GlobalString.LogType.OpLogCategory);
                                LADPCheckManager.GetLDAPAuth();
                                LogFactory.Write("同步LDAP信息結束", GlobalString.LogType.OpLogCategory);
                            }
                        }
                        break;
                    //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 start 
                    case GlobalString.BatchRID.BatchSeven:
                        if (GlobalString.BatchRunFlag.BatchSeven == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchSeven)))
                        {
                            //201005CR 增加物理寄卡單、DM警訊 add by Ian Huang 2010/06/11 start
                            LogFactory.Write("物理寄卡單、DM警訊開始", GlobalString.LogType.OpLogCategory);
                            WarningBatch WB = new WarningBatch();
                            WB.BatchStart();
                            LogFactory.Write("物理寄卡單、DM警訊結束", GlobalString.LogType.OpLogCategory);

                            WB = null;
                            GC.Collect();
                            //201005CR 增加物理寄卡單、DM警訊 add by Ian Huang 2010/06/11 end
                        }
                        break;
                    case GlobalString.BatchRID.BatchEight:
                        if (GlobalString.BatchRunFlag.BatchEight == GlobalString.BatchStatus.Stop && !(CheckBatchStatus(GlobalString.BatchRID.BatchEight)))
                        {
                            // 2021-05-18 新增次月下市預測表匯入 陳永銘
                            UpdateBatchStatus(GlobalString.BatchRID.BatchEight, GlobalString.BatchStatus.Run);
                            GlobalString.BatchRunFlag.BatchEight = GlobalString.BatchStatus.Run;

                            LogFactory.Write("匯入次月下市預測表開始", GlobalString.LogType.OpLogCategory);
                            InOut006BL bl66 = new InOut006BL();
                            ArrayList al = bl66.MonthDelistCard();
                            Batch66(al);
                            LogFactory.Write("匯入次月下市預測表結束", GlobalString.LogType.OpLogCategory);

                            bl66 = null;
                            GC.Collect();

                        }
                        break;
                    //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 end
                    default:
                        break;
                }


            }
            catch (Exception ex)
            {
                // Legend 2018/4/12 添加方法名字
                LogFactory.Write("批次程式執行RunBatch錯誤:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            finally
            {
                switch (BatchRid)
                {
                    case GlobalString.BatchRID.BatchOne:
                        GlobalString.BatchRunFlag.BatchOne = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchOne, GlobalString.BatchStatus.Stop);
                        break;
                    case GlobalString.BatchRID.BatchTwo:
                        GlobalString.BatchRunFlag.BatchTwo = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchTwo, GlobalString.BatchStatus.Stop);
                        GlobalString.BatchRunFlag.BatchThree = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchThree, GlobalString.BatchStatus.Stop);
                        break;
                    case GlobalString.BatchRID.BatchThree:
                        GlobalString.BatchRunFlag.BatchOne = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchOne, GlobalString.BatchStatus.Stop);
                        break;
                    case GlobalString.BatchRID.BatchFour:
                        GlobalString.BatchRunFlag.BatchFour = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchFour, GlobalString.BatchStatus.Stop);
                        break;
                    case GlobalString.BatchRID.BatchFive:
                        GlobalString.BatchRunFlag.BatchFive = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchFive, GlobalString.BatchStatus.Stop);
                        break;
                    case GlobalString.BatchRID.BatchSix:
                        GlobalString.BatchRunFlag.BatchSix = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchSix, GlobalString.BatchStatus.Stop);
                        break;
                    //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 start 
                    case GlobalString.BatchRID.BatchSeven:
                        GlobalString.BatchRunFlag.BatchSeven = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchSeven, GlobalString.BatchStatus.Stop);
                        break;
                    case GlobalString.BatchRID.BatchEight:
                        GlobalString.BatchRunFlag.BatchEight = GlobalString.BatchStatus.Stop; ;
                        UpdateBatchStatus(GlobalString.BatchRID.BatchEight, GlobalString.BatchStatus.Stop);
                        break;
                    //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 end
                    default:
                        break;
                }


            }
            //Batch61(al);

        }

        /// <summary>
        /// 匯入廠商資料
        /// </summary>
        /// <returns></returns>
        private void ImportFactoryData(String seq)
        {
            string sFactoryList = ConfigurationManager.AppSettings["FactoryList"];
            string[] saFactoryList = sFactoryList.Split(',');

            for (int i = 0; i < saFactoryList.Length; i++)
            {
                LogFactory.Write("第" + seq + "次匯入廠商" + saFactoryList[i] + "異動開始", GlobalString.LogType.OpLogCategory);
                InOut002BL bl62 = new InOut002BL();
                ArrayList al = bl62.DownLoadModify(saFactoryList[i]);
                Batch62(al);
                LogFactory.Write("第" + seq + "次匯入廠商" + saFactoryList[i] + "異動結束", GlobalString.LogType.OpLogCategory);

                LogFactory.Write("第" + seq + "次匯入廠商" + saFactoryList[i] + "特殊代制費用開始", GlobalString.LogType.OpLogCategory);
                Finance0024BL bl524 = new Finance0024BL();
                al = bl524.DownLoadModify(saFactoryList[i]);
                Batch524(al);
                LogFactory.Write("第" + seq + "次匯入廠商" + saFactoryList[i] + "特殊代制費用結束", GlobalString.LogType.OpLogCategory);

                LogFactory.Write("第" + seq + "次匯入廠商" + saFactoryList[i] + "物料異動開始", GlobalString.LogType.OpLogCategory);
                Depository010BL bl410 = new Depository010BL();
                al = bl410.Download(saFactoryList[i]);
                Batch410(al);
                LogFactory.Write("第" + seq + "次匯入廠商" + saFactoryList[i] + "物料異動結束", GlobalString.LogType.OpLogCategory);
            }
        }
        /// <summary>
        /// (替換前版面)匯入廠商資料
        /// </summary>
        /// <returns></returns>
        private void ImportFactoryReplaceData(String seq)
        {
            string sFactoryList = ConfigurationManager.AppSettings["FactoryList"];
            string[] saFactoryList = sFactoryList.Split(',');

            for (int i = 0; i < saFactoryList.Length; i++)
            {
                LogFactory.Write("第" + seq + "次匯入(替換前版面)廠商" + saFactoryList[i] + "異動開始", GlobalString.LogType.OpLogCategory);
                CIMSClass.Business.InOut007BL bl67 = new CIMSClass.Business.InOut007BL();
                ArrayList al = bl67.DownLoadModify(saFactoryList[i]);
                Batch67(al);
                LogFactory.Write("第" + seq + "次匯入(替換前版面)廠商" + saFactoryList[i] + "異動結束", GlobalString.LogType.OpLogCategory);
            }
        }
        /// <summary>
        /// 檢查是否有可執行的批次
        /// </summary>
        /// <param name="NowTime">當前時間</param>
        public void CheckBatch(DateTime NowTime)
        {
            string NowTimeYear = NowTime.ToString("MMdd HH:mm:ss");
            string NowTimeMonth = NowTime.ToString("dd HH:mm:ss");
            string NowTimeDay = NowTime.ToString("HH:mm:ss");
            GlobalString.BatchRunFlag.BatchOne = CheckBatchOne(NowTimeYear, NowTimeMonth, NowTimeDay);
            GlobalString.BatchRunFlag.BatchTwo = CheckBatchTwo(NowTimeYear, NowTimeMonth, NowTimeDay);
            GlobalString.BatchRunFlag.BatchThree = CheckBatchThree(NowTimeYear, NowTimeMonth, NowTimeDay);
            GlobalString.BatchRunFlag.BatchFour = CheckBatchFour(NowTimeYear, NowTimeMonth, NowTimeDay);
            GlobalString.BatchRunFlag.BatchFive = CheckBatchFive(NowTimeYear, NowTimeMonth, NowTimeDay);
            GlobalString.BatchRunFlag.BatchSix = CheckBatchSix(NowTimeYear, NowTimeMonth, NowTimeDay);
            //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 start 
            GlobalString.BatchRunFlag.BatchSeven = CheckBatchSeven(NowTimeYear, NowTimeMonth, NowTimeDay);
            GlobalString.BatchRunFlag.BatchEight = CheckBatchEight(NowTimeYear, NowTimeMonth, NowTimeDay);
            //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 end

        }
        #region 檢查批次是否可執行
        private string CheckBatchOne(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchOne, GlobalString.TriggerTime.BatchOneTimeSpan);
        }
        private string CheckBatchTwo(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchTwo, GlobalString.TriggerTime.BatchTwoTimeSpan);
        }
        private string CheckBatchThree(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchThree, GlobalString.TriggerTime.BatchThreeTimeSpan);
        }
        private string CheckBatchFour(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchFour, GlobalString.TriggerTime.BatchFourTimeSpan);
        }
        private string CheckBatchFive(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchFive, GlobalString.TriggerTime.BatchFiveTimeSpan);
        }
        private string CheckBatchSix(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchSix, GlobalString.TriggerTime.BatchSixTimeSpan);
        }
        //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 start 
        private string CheckBatchSeven(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchSeven, GlobalString.TriggerTime.BatchSevenTimeSpan);
        }
        private string CheckBatchEight(string NowTimeYear, string NowTimeMonth, string NowTimeDay)
        {
            return CheckBatchStatus(NowTimeYear, NowTimeMonth, NowTimeDay, GlobalString.TriggerTime.BatchEight, GlobalString.TriggerTime.BatchEightTimeSpan);
        }
        //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 end
        private string CheckBatchStatus(string NowTimeYear, string NowTimeMonth, string NowTimeDay, string BatchTime, string BatchTimeSpan)
        {
            switch (BatchTimeSpan)
            {
                case GlobalString.TriggerTimeType.Year:
                    if (StringUtil.IsEmpty(BatchTime) == false && BatchTime.IndexOf(NowTimeYear) >= 0)
                    {
                        return GlobalString.BatchStatus.Run;
                    }
                    break;
                case GlobalString.TriggerTimeType.Month:
                    if (StringUtil.IsEmpty(BatchTime) == false && BatchTime.IndexOf(NowTimeMonth) >= 0)
                    {
                        return GlobalString.BatchStatus.Run;
                    }
                    break;
                case GlobalString.TriggerTimeType.Day:
                    if (StringUtil.IsEmpty(BatchTime) == false && BatchTime.IndexOf(NowTimeDay) >= 0)
                    {
                        return GlobalString.BatchStatus.Run;
                    }
                    break;
            }
            return GlobalString.BatchStatus.Stop;
        }
        #endregion

        /// <summary>
        /// 獲取批次執行時間
        /// </summary>
        public void GetTriggerTime()
        {
            GlobalString.TriggerTime.BatchOne = ConfigurationManager.AppSettings["BatchOne"];
            GlobalString.TriggerTime.BatchTwo = ConfigurationManager.AppSettings["BatchTwo"];
            GlobalString.TriggerTime.BatchThree = ConfigurationManager.AppSettings["BatchThree"];
            GlobalString.TriggerTime.BatchFour = ConfigurationManager.AppSettings["BatchFour"];
            GlobalString.TriggerTime.BatchFive = ConfigurationManager.AppSettings["BatchFive"];
            GlobalString.TriggerTime.BatchSix = ConfigurationManager.AppSettings["BatchSix"];
            //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 start 
            GlobalString.TriggerTime.BatchSeven = ConfigurationManager.AppSettings["BatchSeven"];
            GlobalString.TriggerTime.BatchEight = ConfigurationManager.AppSettings["BatchEight"];
            //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 end
            GlobalString.TriggerTime.BatchOneTimeSpan = ConfigurationManager.AppSettings["BatchOneTimeSpan"];
            GlobalString.TriggerTime.BatchTwoTimeSpan = ConfigurationManager.AppSettings["BatchTwoTimeSpan"];
            GlobalString.TriggerTime.BatchThreeTimeSpan = ConfigurationManager.AppSettings["BatchThreeTimeSpan"];
            GlobalString.TriggerTime.BatchFourTimeSpan = ConfigurationManager.AppSettings["BatchFourTimeSpan"];
            GlobalString.TriggerTime.BatchFiveTimeSpan = ConfigurationManager.AppSettings["BatchFiveTimeSpan"];
            GlobalString.TriggerTime.BatchSixTimeSpan = ConfigurationManager.AppSettings["BatchSixTimeSpan"];
            //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 start
            GlobalString.TriggerTime.BatchSevenTimeSpan = ConfigurationManager.AppSettings["BatchSevenTimeSpan"];
            GlobalString.TriggerTime.BatchEightTimeSpan = ConfigurationManager.AppSettings["BatchEightTimeSpan"];
            //200908CR 增加批次為以后預留 add by 楊昆 2009/09/03 end


        }

        /// <summary>
        /// 檢查文檔是否已被匯入過！
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        private bool CheckFileImport(string filename)
        {
            try
            {
                Dictionary<string, object> dirValues = new Dictionary<string, object>();
                dirValues.Add("File_Name", filename);
                object returnValue = dao.ExecuteScalar(SEL_IMPORT_HISTORY, dirValues);
                if (Convert.ToInt32(returnValue) > 0)
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
                // Legend 2018/04/13 添加方法名
                LogFactory.Write("檢查文檔是否已被匯入過CheckFileImport,檔名:" + filename + " 失敗：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }

        /// <summary>
        /// 判斷日期是否為工作日
        /// </summary>
        /// <param name="WorkDate"></param>
        /// <returns>為工作日返回true</returns>
        public bool CheckWorkDate(DateTime WorkDate)
        {
            try
            {
                Dictionary<string, object> dirValues = new Dictionary<string, object>();
                //2021-05-18 字串修正 陳永銘
                dirValues.Add("Date_Time", WorkDate.ToString("yyyy-MM-dd 00:00:00.000"));
                object returnValue = dao.ExecuteScalar(SEL_WORK_DATE, dirValues);
                if (Convert.ToInt32(returnValue) > 0)
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
                // Legend 2018/04/12 添加方法名
                LogFactory.Write("獲取工作日信息CheckWorkDate錯誤：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }

        /// <summary>
        /// 批次6.1
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch61(ArrayList FileList)
        {
            try
            {
                InOut001BL Batch61 = new InOut001BL();
                DataSet dst = new DataSet();
                string result = "";
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {
                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    dst = Batch61.ImportCheck(str[1], str[0]);
                    if (StringUtil.IsEmpty(Batch61.strErr))
                    {
                        if (dst.Tables.Count > 0)
                        {
                            result = Batch61.ImportSubTotal(dst, str[0], str[1]);
                            if (result == "")
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "1";
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入小計檔存入DB, Batch61方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

        }
        /// <summary>
        /// 批次6.2
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch62(ArrayList FileList)
        {
            try
            {
                InOut002BL Batch62 = new InOut002BL();
                //200908CR將替換后版面的廠商異動檔匯入修改成替換前版面與替換后版面同時匯入 ADD BY 楊昆 2009/09/21 start
                CIMSClass.Business.InOut007BL Batch67 = new CIMSClass.Business.InOut007BL();
                //200908CR將替換后版面的廠商異動檔匯入修改成替換前版面與替換后版面同時匯入 ADD BY 楊昆 2009/09/21 end
                DataSet dst = new DataSet();
                string result = "";
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {

                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    string[] str2 = str[1].Split('-');
                    //200908CR將替換后版面的廠商異動檔匯入修改成替換前版面與替換后版面同時匯入 ADD BY 楊昆 2009/09/21 start
                    //dst = Batch62.ImportCheck(str[0]+str[1],str2[1],str2[2]);
                    dst = Batch67.ImportCheck(str[0] + str[1], str2[1], str2[2]);
                    //200908CR將替換后版面的廠商異動檔匯入修改成替換前版面與替換后版面同時匯入 ADD BY 楊昆 2009/09/21 end
                    if (StringUtil.IsEmpty(Batch62.strErr))
                    {
                        if (dst.Tables.Count > 0)
                        {
                            //200908CR將替換后版面的廠商異動檔匯入修改成替換前版面與替換后版面同時匯入 ADD BY 楊昆 2009/09/21 start
                            //result = Batch62.ImportCardTypeChange(dst, str2[1]);
                            result = Batch67.ImportCardTypeChange(dst.Tables[0], str2[1], "1");
                            //200908CR將替換后版面的廠商異動檔匯入修改成替換前版面與替換后版面同時匯入 ADD BY 楊昆 2009/09/21 end
                            if (result == "")
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "2";
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入廠商異動存入DB, Batch62方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
        }
        /// <summary>
        /// 批次6.3
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        private void Batch63()
        {
            try
            {
                InOut003BL Batch63 = new InOut003BL();
                string date = Batch63.GetLastStock_Date();

                DataTable dt = new DataTable();
                if (date != "" && date != DateTime.Now.ToString("yyyy/MM/dd"))
                {
                    dt = Batch63.GetNeedStock_Date(date);
                }
                if (dt.Rows.Count != 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (Batch63.Compare(Convert.ToDateTime(dt.Rows[i][0].ToString())))
                        {
                            try
                            {
                                Batch63.DaySurplus(Convert.ToDateTime(dt.Rows[i][0].ToString()));
                            }
                            catch
                            {
                                return;
                            }
                        }
                        else
                            return;
                        //else
                        //{
                        //    this.EmailAlert63("系統日結失敗！");
                        //}
                    }
                }

                Batch63 = null;
            }
            catch (Exception ex)
            {
                LogFactory.Write("日結處理存入DB, Batch63方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
        }
        /// <summary>
        /// 批次6.4
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch64(ArrayList FileList)
        {
            try
            {
                InOut004BL Batch64 = new InOut004BL();
                DataSet dst = new DataSet();
                string result = "";
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {

                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    dst = Batch64.DetailCheck(str[0] + str[1]);
                    if (StringUtil.IsEmpty(Batch64.strErr))
                    {
                        if (dst.Tables.Count > 0)
                        {
                            result = Batch64.In(dst);
                            if (result == "")
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "6";
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入年度換卡預測檔存入DB, Batch64方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

        }
        /// <summary>
        /// 批次6.5
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch65(ArrayList FileList)
        {
            try
            {
                InOut005BL Batch65 = new InOut005BL();
                DataSet dst = new DataSet();
                string result = "";
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {

                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    dst = Batch65.DetailCheck(str[0] + str[1]);
                    if (StringUtil.IsEmpty(Batch65.strErr))
                    {
                        if (dst.Tables.Count > 0)
                        {
                            result = Batch65.In(dst);
                            if (result == "")
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "5";
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入次月換卡預測檔存入DB, Batch65方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

        }
        /// <summary>
        /// 批次6.6
        /// 2021-03-12 新增次月下市預測表匯入 陳永銘
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch66(ArrayList FileList)
        {
            try
            {
                InOut006BL Batch66 = new InOut006BL();
                DataSet dst = new DataSet();
                string result = "";
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {

                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    dst = Batch66.DetailCheck(str[0] + str[1]);
                    if (StringUtil.IsEmpty(Batch66.strErr))
                    {
                        if (dst.Tables.Count > 0)
                        {
                            result = Batch66.In(dst);
                            if (result == "")
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "6";
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入次月下市預測表存入DB, Batch66方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

        }
        /// <summary>
        /// 批次6.7
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch67(ArrayList FileList)
        {

            CIMSClass.Business.InOut007BL Batch67 = new CIMSClass.Business.InOut007BL();

            DataSet dst = new DataSet();
            string result = "";
            IMPORT_HISTORY IM = new IMPORT_HISTORY();
            for (int i = 0; i < FileList.Count; i++)
            {

                string[] str = (string[])FileList[i];

                if (this.CheckFileImport(str[1]))
                    continue;

                string[] str2 = str[1].Split('-');
                dst = Batch67.ImportCheck(str[0] + str[1], str2[1], str2[2]);
                if (StringUtil.IsEmpty(Batch67.strErr))
                {
                    if (dst.Tables.Count > 0)
                    {
                        result = Batch67.ImportCardTypeChange(dst.Tables[0], str2[1], "1");
                        if (result == "")
                        {
                            IM.File_Name = str[1];
                            IM.File_Type = "7";
                            IM.Import_Date = DateTime.Now;
                            dao.Add<IMPORT_HISTORY>(IM, "RID");
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 批次4.10
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch410(ArrayList FileList)
        {
            try
            {
                Depository010BL Batch410 = new Depository010BL();
                DataSet dst = new DataSet();
                string result = "";
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {

                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    string[] strname = str[1].Split('-');
                    Dictionary<string, object> dirValues = new Dictionary<string, object>();
                    dirValues.Add("FileUpd", str[0] + str[1]);
                    dirValues.Add("Factory_ShortName_EN", strname[2].Split('.')[0]);
                    dst = Batch410.CheckIn(dirValues);
                    if (StringUtil.IsEmpty(Batch410.strErr))
                    {
                        if (dst.Tables.Count > 0)
                        {
                            result = Batch410.In(dst);
                            if (result == "")
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "4";
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                                return;
                            }
                        }
                    }

                    string[] arg = new string[1];
                    arg[0] = str[0] + str[1];
                    Warning.SetWarning(GlobalString.WarningType.MaterialDataInLost, arg);

                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入廠商物料異動存入DB, Batch410方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

        }
        /// <summary>
        /// 批次5.2.4
        /// 加try catch add by judy 2018/03/28
        /// </summary>
        /// <param name="FileList"></param>
        private void Batch524(ArrayList FileList)
        {
            try
            {
                Finance0024BL Batch524 = new Finance0024BL();
                DataTable dt = new DataTable();
                bool result = false;
                IMPORT_HISTORY IM = new IMPORT_HISTORY();
                for (int i = 0; i < FileList.Count; i++)
                {

                    string[] str = (string[])FileList[i];

                    if (this.CheckFileImport(str[1]))
                        continue;

                    string[] strname = str[1].Split('-');//取得Perso廠商英文簡稱

                    dt = Batch524.ImportCheck(str[0] + str[1], strname[2].Split('.')[0]);
                    if (StringUtil.IsEmpty(Batch524.strErr))
                    {
                        if (dt.Rows.Count > 0)
                        {
                            result = Batch524.SaveSpecialIn(dt);
                            if (result)
                            {
                                IM.File_Name = str[1];
                                IM.File_Type = "3"; //代製費用的類型是3
                                IM.Import_Date = DateTime.Now;
                                dao.Add<IMPORT_HISTORY>(IM, "RID");
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("匯入廠商特殊代制費用存入DB, Batch524方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
        }
        #region 郵件警示
        /// <summary>
        /// 紀錄6.1警示信息
        /// </summary>
        /// <param name="ErrList">錯誤的清單ArrayList</param>
        public void EmailAlert(string Errs)
        {


            string System_Show;
            string Mail_Show;
            int Warning_RID;
            string UserID;
            string Warning_Content;
            string Email;
            DataSet ds = dao.GetList(SEL_SUBTOTAL_ALERT);
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                System_Show = dr["System_Show"].ToString();
                Mail_Show = dr["Mail_Show"].ToString();
                Warning_RID = Convert.ToInt32(dr["Warning_RID"]);
                UserID = dr["UserID"].ToString();
                Warning_Content = dr["Warning_Content"].ToString();
                Email = dr["Email"].ToString();

                if (System_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    Insert_WARNING_INFO(Warning_RID, "Y", Errs, UserID);
                }
                if (Mail_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    SendMail(Email, Warning_Content, Errs);
                }
            }

        }
        /// <summary>
        /// 紀錄6.2警示信息
        /// </summary>
        /// <param name="ErrList">錯誤的清單ArrayList</param>
        public void EmailAlert62(string Errs)
        {


            string System_Show;
            string Mail_Show;
            int Warning_RID;
            string UserID;
            string Warning_Content;
            string Email;
            DataSet ds = dao.GetList(SEL_FACTORY_CHANGE_ALERT + GlobalString.MailErroFlag.BatchTwo.ToString());
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                System_Show = dr["System_Show"].ToString();
                Mail_Show = dr["Mail_Show"].ToString();
                Warning_RID = Convert.ToInt32(dr["Warning_RID"]);
                UserID = dr["UserID"].ToString();
                Warning_Content = dr["Warning_Content"].ToString();
                Email = dr["Email"].ToString();

                if (System_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    Insert_WARNING_INFO(Warning_RID, "Y", Errs, UserID);
                }
                if (Mail_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    SendMail(Email, Warning_Content, Errs);
                }
            }

        }
        /// <summary>
        /// 紀錄6.3警示信息
        /// </summary>
        /// <param name="ErrList">錯誤的清單ArrayList</param>
        public void EmailAlert63(string Errs)
        {


            string System_Show;
            string Mail_Show;
            int Warning_RID;
            string UserID;
            string Warning_Content;
            string Email;
            DataSet ds = dao.GetList(SEL_DAY_SURPLUS_ALERT + GlobalString.BatchRID.BatchThree.ToString());
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                System_Show = dr["System_Show"].ToString();
                Mail_Show = dr["Mail_Show"].ToString();
                Warning_RID = Convert.ToInt32(dr["Warning_RID"]);
                UserID = dr["UserID"].ToString();
                Warning_Content = dr["Warning_Content"].ToString();
                Email = dr["Email"].ToString();

                if (System_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    Insert_WARNING_INFO(Warning_RID, "Y", Errs, UserID);
                }
                if (Mail_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    SendMail(Email, Warning_Content, Errs);
                }
            }

        }
        /// <summary>
        /// 紀錄6.4警示信息
        /// </summary>
        /// <param name="ErrList">錯誤的清單ArrayList</param>
        public void EmailAlert64(string Errs)
        {


            string System_Show;
            string Mail_Show;
            int Warning_RID;
            string UserID;
            string Warning_Content;
            string Email;
            DataSet ds = dao.GetList(SEL_YEAR_REPLACE_CARD_ALERT + GlobalString.MailErroFlag.BatchFour.ToString());
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                System_Show = dr["System_Show"].ToString();
                Mail_Show = dr["Mail_Show"].ToString();
                Warning_RID = Convert.ToInt32(dr["Warning_RID"]);
                UserID = dr["UserID"].ToString();
                Warning_Content = dr["Warning_Content"].ToString();
                Email = dr["Email"].ToString();

                if (System_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    Insert_WARNING_INFO(Warning_RID, "Y", Errs, UserID);
                }
                if (Mail_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    SendMail(Email, Warning_Content, Errs);
                }
            }

        }
        /// <summary>
        /// 紀錄6.5警示信息
        /// </summary>
        /// <param name="ErrList">錯誤的清單ArrayList</param>
        public void EmailAlert65(string Errs)
        {


            string System_Show;
            string Mail_Show;
            int Warning_RID;
            string UserID;
            string Warning_Content;
            string Email;
            DataSet ds = dao.GetList(SEL_NEXT_MONTH_REPLACE_CARD_ALERT + GlobalString.MailErroFlag.BatchFive.ToString());
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                System_Show = dr["System_Show"].ToString();
                Mail_Show = dr["Mail_Show"].ToString();
                Warning_RID = Convert.ToInt32(dr["Warning_RID"]);
                UserID = dr["UserID"].ToString();
                Warning_Content = dr["Warning_Content"].ToString();
                Email = dr["Email"].ToString();

                if (System_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    Insert_WARNING_INFO(Warning_RID, "Y", Errs, UserID);
                }
                if (Mail_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    SendMail(Email, Warning_Content, Errs);
                }
            }

        }
        /// <summary>
        /// 紀錄6.7警示信息
        /// </summary>
        /// <param name="ErrList">錯誤的清單ArrayList</param>
        public void EmailAlert67(string Errs)
        {


            string System_Show;
            string Mail_Show;
            int Warning_RID;
            string UserID;
            string Warning_Content;
            string Email;
            DataSet ds = dao.GetList(SEL_FACTORY_CHANGE_ALERT + GlobalString.MailErroFlag.BatchSeven.ToString());
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                System_Show = dr["System_Show"].ToString();
                Mail_Show = dr["Mail_Show"].ToString();
                Warning_RID = Convert.ToInt32(dr["Warning_RID"]);
                UserID = dr["UserID"].ToString();
                Warning_Content = dr["Warning_Content"].ToString();
                Email = dr["Email"].ToString();

                if (System_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    Insert_WARNING_INFO(Warning_RID, "Y", Errs, UserID);
                }
                if (Mail_Show.ToUpper() == GlobalString.ActFlag.Do)
                {
                    SendMail(Email, Warning_Content, Errs);
                }
            }

        }
        /// <summary>
        /// 寫入警示信息到資料庫
        /// </summary>
        /// <param name="RID"></param>
        /// <param name="IsShow"></param>
        /// <param name="Content"></param>
        /// <param name="UserID"></param>
        /// <returns></returns>
        private bool Insert_WARNING_INFO(int RID, string IsShow, string Content, string UserID)
        {
            try
            {
                WARNING_INFO wInfo = new WARNING_INFO();
                wInfo.RID = RID;
                wInfo.Is_Show = IsShow;
                wInfo.Warning_Content = Content;
                wInfo.UserID = UserID;
                dao.AddAndGetID<WARNING_INFO>(wInfo, "RID");
                return true;
            }
            catch (Exception ex)
            {
                // Legend 2018/04/13 調整記錄的錯誤訊息
                LogFactory.Write("寫入警示信息到資料庫WARNING_INFO錯誤：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        /// <summary>
        /// 發送郵件
        /// </summary>
        /// <param name="MailAddress"></param>
        /// <param name="Subject"></param>
        /// <param name="Body"></param>
        /// <returns></returns>
        public bool SendMail(string MailAddress, string Subject, string Body)
        {
            return MailBase.SendMail(MailAddress, Subject, Body);
        }
        #endregion
    }
}
