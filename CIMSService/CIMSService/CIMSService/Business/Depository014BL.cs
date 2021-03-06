using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using CIMSBatch.Model;
using System.Data.Common;
using CIMSBatch.FTP;
using CIMSBatch.Mail;
using CIMSBatch;
using CIMSBatch.Public;
namespace CIMSBatch.Business
{
    class Depository014BL : BaseLogic
    {
        #region SQL定義
        //查詢前日結庫存量
        public const string SEL_STOCKS = "SELECT top 1 Stock_Date,Stocks_Number"
                                    + " FROM CARDTYPE_STOCKS"
                                    + " WHERE RST='A' AND Perso_Factory_RID = @persoRid AND CardType_RID = @cardRid"
                                    + " ORDER BY Stock_Date desc";
        //查詢前日結庫存量
        public const string SEL_STOCKS_NUM = "SELECT Stocks_Number"
                                    + " FROM CARDTYPE_STOCKS"
                                    + " WHERE RST='A' AND Perso_Factory_RID = @persoRid AND CardType_RID = @cardRid and Stock_Date = @stockDate";
        //查詢最近日結后入庫量
        public const string SEL_DEPOSITORY_STOCK = "SELECT ISNULL(Sum(Income_Number),0) as Number"
                                    + " FROM DEPOSITORY_STOCK"
                                    + " WHERE RST = 'A' AND Income_Date > @Stock_Date and Income_Date < @today AND Perso_Factory_RID = @persoRid AND Space_Short_RID = @cardRid";
        //查詢最近日結后退貨量
        public const string SEL_DEPOSITORY_CANCEl_NUMBER = "SELECT ISNULL(Sum(Cancel_Number),0) as Number"
                                    + " FROM DEPOSITORY_CANCEL"
                                    + " WHERE RST = 'A' AND Cancel_Date > @Stock_Date and Cancel_Date < @today  AND Perso_Factory_RID = @persoRid AND Space_Short_RID = @cardRid";
        //查詢最近日結后再入庫量
        public const string SEL_DEPOSITORY_RESTOCK = "SELECT ISNULL(Sum(Reincome_Number),0) as Number"
                                    + " FROM DEPOSITORY_RESTOCK"
                                    + " WHERE RST = 'A' AND Reincome_Date > @Stock_Date and Reincome_Date < @today AND Perso_Factory_RID = @persoRid AND Space_Short_RID = @cardRid";
        //查詢最近日結后卡片轉移量

        public const string CARDTYPE_STOCKS_MOVE_IN = "SELECT ISNULL(sum(Move_Number),0) as Number"
                                    + " FROM CARDTYPE_STOCKS_MOVE"
                                    + " WHERE RST = 'A' AND Move_Date > @Stock_Date and Move_Date < @today AND To_Factory_RID = @persoRid AND CardType_RID = @cardRid";

        public const string CARDTYPE_STOCKS_MOVE_OUT = "SELECT ISNULL(sum(Move_Number),0) as Number"
                                    + " FROM CARDTYPE_STOCKS_MOVE"
                                    + " WHERE RST = 'A' AND Move_Date > @Stock_Date and Move_Date < @today AND From_Factory_RID = @persoRid AND CardType_RID = @cardRid";

        public const string SEL_FACTORY_CHANGE_NUM = "select Status_RID,ISNULL(sum(Number),0) as number,Perso_Factory_RID, PHOTO,AFFINITY,TYPE "
                                     + " from FACTORY_CHANGE_IMPORT "
                                     + " where RST = 'A' AND TYPE = @TYPE "
                                     + " AND PHOTO = @PHOTO AND AFFINITY = @AFFINITY AND Perso_Factory_RID = @Perso_Factory_RID "
                                     + " AND Date_Time > @Stock_Date and Date_Time < @today and Status_RID in (5,6,7,8,9,10,11,12,13) group by Status_RID,Perso_Factory_RID,PHOTO,AFFINITY,TYPE";


        public const string SEL_EXPRESSION_INFO = "select * from EXPRESSIONS_DEFINE "
                                    + "where RST = 'A' AND Expressions_RID = @Expressions_RID and type_rid >4";
        
  
        public const string CON_CHECK_DATE = "SELECT count(*) from cardtype_stocks where rst = 'A' and Stock_Date = @CheckDate";

        public const string SEL_PRECARD = "SELECT stocks.Stocks_Number, factory.Factory_ShortName_CN, type.Name,"
                                    + " type.TYPE, type.AFFINITY, type.PHOTO, stocks.Stock_Date,"
                                    + " stocks.Perso_Factory_RID, stocks.CardType_RID"
                                    + " FROM CARDTYPE_STOCKS  AS stocks"
                                    + " LEFT OUTER JOIN CARD_TYPE AS type ON stocks.CardType_RID = type.RID"
                                    + " LEFT OUTER JOIN FACTORY AS factory ON stocks.Perso_Factory_RID = factory.RID"
                                    + " where stocks.rst = 'A' and stock_date = (select Max(stock_date) from cardtype_stocks)";

        public const string SEL_WORKDAY = "Select Max(Date_Time) FROM WORK_DATE where RST = 'A' AND Is_WorkDay = 'Y' AND Date_Time < @Month";
      
        //取開始時間到結束時間的Action= @action 的小計數量
        public const string SEL_PREMONTHS_NUMBER = "SELECT ISNULL(sum(SI.Number),0) FROM SUBTOTAL_IMPORT AS SI "
                                    + " WHERE SI.RST = 'A' AND SI.Perso_Factory_RID = @PersoRid and SI.Date_Time >= @begin_date and SI.Date_Time < @end_date and SI.action = @action"
                                    + " and SI.PHOTO = @photo and SI.TYPE = @type and SI.AFFINITY = @affinity ";
              
        public const string SEL_PARAM = "SELECT  Param_Name FROM PARAM WHERE ParamType_Code = @Param ";

        public const string CON_DAYLY_DATA = "SELECT count(*) from DAYLY_MONITOR where Perso_factory_RID = @perso_rid "
                                    + " and cDay = @cDay and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType ";

        public const string SEL_DAYLY_DATA = "SELECT Top 1 RID from DAYLY_MONITOR where Perso_factory_RID = @perso_rid "
                                    + " and cDay = @cDay and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType ";
        public const string SEL_ALL_DAYLY_DATA = "SELECT RID,CDay from DAYLY_MONITOR where Perso_factory_RID = @perso_rid "
                                  + " and cDay >= @StartDay  and cDay <= @EndDay and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType "
                                  +" order by RID desc ";

        public const string SEL_DAYLY_DATA_CNumber = "SELECT Top 1 CNumber from DAYLY_MONITOR where Perso_factory_RID = @perso_rid "
                                    + " and cDay = @cDay and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType ";


        public const string DEL_DAY_DATA = "delete FROM  DAYLY_MONITOR "
                                           + " where rid not in (select DAYLY_MONITOR.rid from DAYLY_MONITOR"
                                           + " left join card_type on card_type.[TYPE]=DAYLY_MONITOR.[TYPE]"
                                           + " AND card_type.[AFFINITY]=DAYLY_MONITOR.[AFFINITY]"
                                           + " AND card_type.[PHOTO]=DAYLY_MONITOR.[PHOTO]"
                                           + " inner join CARDTYPE_STOCKS"
                                           + " on CARDTYPE_STOCKS.perso_factory_rid =DAYLY_MONITOR.perso_factory_rid"
                                           + " and card_type.rid= CARDTYPE_STOCKS.cardtype_rid"
                                          // + " WHERE  CARDTYPE_STOCKS.stock_date = (select Max(stock_date) from cardtype_stocks))";
                                           + " WHERE  CARDTYPE_STOCKS.stock_date in (select Max(stock_date) from cardtype_stocks))";

        public const string DEL_NOT_WORKDAY = "DELETE FROM dayly_monitor WHERE CDAY IN (select date_time from WORK_DATE where is_workDay='N')";
        
        public const string SEL_BEFORE_DATE = "select min(date_time) from( select top 5 * from work_date as table1 where is_workDay='Y' and date_time<'2008/11/09' order by date_time desc)a";

        public const string SEL_WORKDAYS = "select top 60 Date_Time FROM WORK_DATE where RST = 'A' AND Is_WorkDay = 'Y' AND Date_Time >= @today order by date_time"
            + " select top (@count) date_time from work_date where rst='A' and is_workDay='Y' and date_time <=@today order by date_time desc";

        public const string SEL_SUBTOTAL = "select a.date_time,isnull(si.A1,0) AS A1,isnull(si.A2,0) AS A2 ,isnull(si.A3,0) AS A3"
                            + " from (select top (@N) work_date.date_time from work_date "
                            + " where rst='A' and is_workDay='Y' "
                            + "and work_date.date_time <@Now"
                            + " order by work_date.date_time desc) a"
                            + " left outer join "
                            + " (SELECT date_time, type,"
                            + " SUM(CASE action WHEN 1 THEN Number ELSE 0 END) AS A1,"
                            + " SUM(CASE action WHEN 2 THEN Number ELSE 0 END) AS A2,"
                            + " SUM(CASE action WHEN 3 THEN Number ELSE 0 END) AS A3,"                          
                            + " affinity,photo"
                            //+",perso_factory_rid"
                            + " FROM subtotal_import"
                            + " where affinity=@affinity and photo=@photo and type=@type "
                            //+ " and perso_factory_rid=@perso_factory_rid"
                            + " GROUP BY date_time,affinity,photo,"
                            //+ " perso_factory_rid,"
                            + " type"
                            + " ) si on a.date_time=si.date_time ";

        public const string SEL_DEPOSITORY_SUBTOTAL_CHANGE_DATA = "SELECT Stocks_Number, Perso_Factory_RID, CardType_RID,TYPE,AFFINITY,photo,Name,stock_date"
                            + " FROM  CARDTYPE_STOCKS "
                            + " left join card_type on card_type.rid=CARDTYPE_STOCKS.CardType_RID"
                            + " where Stock_Date=@stock_date"                   
                            + " select  Change_Date, Type, Photo, Affinity, Perso_Factory_RID, Number from FORE_CHANGE_CARD_DETAIL"
                            + " where Change_Date>=@beginMonth and change_Date<=@endMonth"
                            + " select sum(order_form_detail.number) as number,cardtype_rid,"
                            + " Delivery_Address_RID,fore_delivery_date,sum(stock_Number) as stock_Number"
                            + " from order_form_detail"
                            + " left join ORDER_FORM ON ORDER_FORM.OrderForm_RID = order_form_detail.OrderForm_RID"                
                            + " left join depository_stock on "
                            + " depository_stock.orderform_detail_rid = order_form_detail.orderform_detail_rid"
                            + " and depository_stock.income_date < fore_delivery_date"
                            + " WHERE fore_delivery_date>=@today and fore_delivery_date<@EndOrderDate and ORDER_FORM.Pass_Status = 4"
                            + " and order_form_detail.case_status<>'Y'"
                            + " group by cardtype_rid,"
                            + " Delivery_Address_RID,fore_delivery_date";
        public const string SEL_STOCK_DATE = "select max(stock_date) from CARDTYPE_STOCKS";
        public const string SEL_THIS_MONTH_CHANGE = "select type,affinity,photo,perso_factory_rid,sum(number) as change_number from subtotal_import "
                            + " where action='5' "
                            + " and date_time>@MonthBegin and date_time<=@today "
                            +" group by type,affinity,photo,perso_factory_rid";

        //public const string SEL_PERSO_CARDTYPE = "select pc.* from  PERSO_CARDTYPE pc inner join cardtype_stocks cs on pc.cardtype_rid=cs.cardtype_rid and stock_date=@stockDate";
        public const string SEL_PERSO_CARDTYPE = "select pc.* from  PERSO_CARDTYPE pc ";
                     
#endregion
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        decimal changePercent = 0M;  
        //查詢各廠商各卡種的資料
        public DataSet getStockCardType()
        {
            //執行SQL語句
            DataSet dstSafeStockInfo = null;
            dstSafeStockInfo = dao.GetList(SEL_PRECARD);
            return dstSafeStockInfo;
        }

        /// <summary>
        /// 主程序入口
        /// </summary>
        public void ComputeForeData()
        {           
            //取換卡百分比
            string strPercent = paramList(GlobalString.cardparamType.Percent);
            if (strPercent == "erro")
                return;
            changePercent = Convert.ToDecimal(strPercent.Substring(0, strPercent.Length - 1)) / 100;     
            string strN = paramList(GlobalString.cardparamType.YType);
            string strY1 = paramList(GlobalString.cardparamType.Y1);
            string strY2 = paramList(GlobalString.cardparamType.Y2);
            string strY3 = paramList(GlobalString.cardparamType.Y3);
            if (strN != "erro" && strY1 != "erro" && strY2 != "erro" && strY3 != "erro")
            {
                try
                {
                    int N = int.Parse(strN.Substring(0, strN.Length - 1));
                    int Y1 = int.Parse(strY1.Substring(0, strY1.Length - 1));
                    int Y2 = int.Parse(strY2.Substring(0, strY2.Length - 1));
                    int Y3 = int.Parse(strY3.Substring(0, strY3.Length - 1));
                    dirValues.Clear();
                    object oStockDate = dao.ExecuteScalar(SEL_STOCK_DATE, dirValues);
                    DateTime LastCheckDate = new DateTime();
                    if (oStockDate != null)
                    {
                        LastCheckDate = Convert.ToDateTime(oStockDate);
                    }
                    else
                    {
                        return;
                    }
                    dirValues.Clear();
                    dirValues.Add("count", Math.Max(Math.Max(Y1, Y2), Y3));
                    dirValues.Add("today", Convert.ToDateTime( DateTime.Now.ToString("yyyy/MM/dd")));
                    dirValues.Add("stock_date", LastCheckDate);
                    DataSet dstWorkDay = dao.GetList(SEL_WORKDAYS, dirValues);
                    DataTable next60Day = dstWorkDay.Tables[0];
                    DataTable oldYDay = dstWorkDay.Tables[1];
                    dirValues.Add("begin_date", oldYDay.Rows[oldYDay.Rows.Count - 1]["Date_Time"]);
                    dirValues.Add("beginMonth", DateTime.Now.ToString("yyyyMM"));
                    dirValues.Add("endMonth", DateTime.Now.AddMonths(3).ToString("yyyyMM"));
                    dirValues.Add("EndOrderDate", next60Day.Rows[next60Day.Rows.Count - 1]["Date_Time"]);
                    DataSet dstDepositoryAndSubAndChange = dao.GetList(SEL_DEPOSITORY_SUBTOTAL_CHANGE_DATA, dirValues);
                    dirValues.Clear();
                    dirValues.Add("MonthBegin", DateTime.Now.ToString("yyyy/MM/01"));
                    dirValues.Add("today", DateTime.Now.ToString("yyyy/MM/dd"));
                    DataTable dbChangedNum = dao.GetList(SEL_THIS_MONTH_CHANGE, dirValues).Tables[0];
                    dao.ExecuteNonQuery(DEL_NOT_WORKDAY);
                    dao.ExecuteNonQuery(DEL_DAY_DATA);
                    DataTable PersoCardtype = getPersoCardtype(LastCheckDate);
                   
                    importForeData(Y1, N, GlobalString.cardparamType.Y1, dstWorkDay, dstDepositoryAndSubAndChange, dbChangedNum, PersoCardtype);
                    importForeData(Y2, N, GlobalString.cardparamType.Y2, dstWorkDay, dstDepositoryAndSubAndChange, dbChangedNum, PersoCardtype);
                    importForeData(Y3, N, GlobalString.cardparamType.Y3, dstWorkDay, dstDepositoryAndSubAndChange, dbChangedNum, PersoCardtype);
                    
                }catch(Exception ex){
                    LogFactory.Write("主程序入口ComputeForeData方法報錯：" + ex.Message, GlobalString.LogType.ErrorCategory); // 主程序入口ComputeForeData方法報錯：add by judy 2018/03/28
                }
            }     
        }
        /// <summary>
        /// 查詢最近一次日結卡種的所有廠商分配比例
        /// </summary>
        /// <param name="LastCheckDate"></param>
        /// <returns></returns>
        public DataTable getPersoCardtype(DateTime LastCheckDate)
        {
            DataTable dtb = new DataTable();

            try
            {
                dirValues.Clear();
                dirValues.Add("stockDate", LastCheckDate);
                dtb = dao.GetList(SEL_PERSO_CARDTYPE, dirValues).Tables[0];
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("查詢最近一次日結卡種的所有廠商分配比例, getPersoCardtype報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dtb;
        }
        
        //存預計結果的數據
        public void importForeData(int Y, int N, string paramCode,
                DataSet dstWorkDay,
                DataSet dstDepositoryAndSubAndChange,
                DataTable changedOverNum,
                DataTable PersoCardtype)
        {
            long Mem = GC.GetTotalMemory(true);//已用內存
            LogFactory.Write("每日監控作業：importForeData(" + Y.ToString() + "," + N.ToString() + ",已用內存：" + Mem.ToString() + ")開始", GlobalString.LogType.OpLogCategory);
                    
            int YCode = int.Parse(paramCode);
            //取庫存記錄中所有的廠商和卡種的組合
            DataTable dtbPersoCardType = dstDepositoryAndSubAndChange.Tables[0];       
            //取換卡数据
            DataTable dtbChange = dstDepositoryAndSubAndChange.Tables[1];
            //取已下单未到货量
            DataTable dtbOrder = dstDepositoryAndSubAndChange.Tables[2];
            //取預估開始日
            DateTime beginDay =Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd"));
            dirValues.Add("beginDay", beginDay);
            //取預估開始日之後的60個工作日    
            DataTable next60Day = dstWorkDay.Tables[0];
            //創建包含所有需要數據的表格
            try
            {
                //遍歷所有廠商與卡種的組合
                foreach (DataRow drow in dtbPersoCardType.Rows)
                {
                    CARD_TYPE cardModel = new CARD_TYPE();
                    cardModel.AFFINITY = drow["AFFINITY"].ToString();
                    cardModel.PHOTO = drow["Photo"].ToString(); //Photo
                    cardModel.TYPE = drow["TYPE"].ToString();//card Type
                    cardModel.RID = int.Parse(drow["CardType_RID"].ToString());
                    cardModel.Name = drow["Name"].ToString();                   
                    //取廠商信息
                    int persoRid = int.Parse(drow["Perso_Factory_RID"].ToString());
                    LogFactory.Write("卡种:" + cardModel.Name + " 厂商: " + persoRid + "开始", GlobalString.LogType.OpLogCategory);
                    #region 計算過去Y天的纍計小計檔的結果
                    dirValues.Clear();
                    dirValues.Add("N",Y);
                    dirValues.Add("Now", beginDay);
                    dirValues.Add("affinity", cardModel.AFFINITY);
                    dirValues.Add("type", cardModel.TYPE);
                    dirValues.Add("photo", cardModel.PHOTO);
                    //dirValues.Add("perso_factory_rid", persoRid);
                    //根據廠商和卡種纍加過去Y天的結果。
                    DataTable dtblAction = dao.GetList(SEL_SUBTOTAL, dirValues).Tables[0];
                    DataTable dtbCountNum = new DataTable();
                    dtbCountNum.Columns.Add("A1");
                    dtbCountNum.Columns.Add("A2");
                    dtbCountNum.Columns.Add("A3");                                    
                    decimal Action1 = 0M;
                    decimal Action2 = 0M;
                    decimal Action3 = 0M;
                    foreach (DataRow actionRow in dtblAction.Rows)
                    {
                        Action1 += Convert.ToDecimal(actionRow["A1"]);
                        Action2 += Convert.ToDecimal(actionRow["A2"]);
                        Action3 += Convert.ToDecimal(actionRow["A3"]);
                        DataRow CountRow = dtbCountNum.NewRow();
                        CountRow["A1"] = Action1;
                        CountRow["A2"] = Action2;
                        CountRow["A3"] = Action3;
                        dtbCountNum.Rows.Add(CountRow);
                    }
                    #endregion
                    //創建結果集
                    bool changeFlag = false;
                    string month = beginDay.ToString("yyyy/MM");
                    DataTable result = creatNewDataTable();
                    #region 計算數據
                    for (int i = 0; i < next60Day.Rows.Count; i++)
                    {
                        DataRow drowResult = result.NewRow();
                        DAYLY_MONITOR dayModel = new DAYLY_MONITOR();
                        dayModel.CDay = Convert.ToDateTime(next60Day.Rows[i]["Date_Time"]);
                        if (i == 0)
                        {
                            dayModel.ANumber = getCurrentStock(persoRid, cardModel.RID, Convert.ToDateTime(next60Day.Rows[0]["Date_Time"]));
                            dayModel.A1Number = dayModel.ANumber;
                        }
                        else
                        {
                            dayModel.ANumber = int.Parse(result.Rows[i - 1]["G"].ToString());
                        }
                        //取前Y個工作日的平均值
                        if (i >= Y)
                        {
                            double numberB = 0D;
                            double numberD1 = 0D;
                            double numberD2 = 0D;
                            for (int j = 1; j <= Y; j++)
                            {
                                numberB += Convert.ToDouble(result.Rows[i - j]["BTotal"].ToString());
                                numberD1 += Convert.ToDouble(result.Rows[i - j]["D1Total"].ToString());
                                numberD2 += Convert.ToDouble(result.Rows[i - j]["D2Total"].ToString());
                            }
                            drowResult["BTotal"] = Convert.ToInt32(Math.Ceiling(numberB / Y));
                            drowResult["D1Total"] = Convert.ToInt32(Math.Ceiling(numberD1 / Y));
                            drowResult["D2Total"] = Convert.ToInt32(Math.Ceiling(numberD2 / Y));                   

                            dayModel.BNumber = splitNumber(cardModel, persoRid,
                                Convert.ToInt32(drowResult["BTotal"]), PersoCardtype);
                            dayModel.D1Number = splitNumber(cardModel, persoRid,
                                Convert.ToInt32(drowResult["D1Total"]), PersoCardtype);
                            dayModel.D2Number = splitNumber(cardModel, persoRid,
                                Convert.ToInt32(drowResult["D2Total"]), PersoCardtype);
                        }
                        else
                        {
                            int index = Y - i-1;
                            if (Y - i-1 > dtbCountNum.Rows.Count - 1) {
                                index = dtbCountNum.Rows.Count - 1;
                            }                         
                            //B.過去Y-i個工作日的新卡數                            
                            decimal numberB = Convert.ToDecimal(dtbCountNum.Rows[index]["A1"]);
                            //D1.過去Y-i個工作日的掛補量                            
                            decimal numberD1 = Convert.ToDecimal(dtbCountNum.Rows[index]["A2"]);
                            //D2.過去Y-i個工作日的毀補量                            
                            decimal numberD2 = Convert.ToDecimal(dtbCountNum.Rows[index]["A3"]);
                            //加上前i個預算值的匯總
                            for (int j = 0; j < i; j++)
                            {
                                numberB += Convert.ToDecimal(result.Rows[j]["BTotal"]);
                                numberD1 += Convert.ToDecimal(result.Rows[j]["D1Total"]);
                                numberD2 += Convert.ToDecimal(result.Rows[j]["D2Total"]);
                            }
                            drowResult["BTotal"] = Convert.ToInt32(Math.Ceiling(numberB / Y));
                            drowResult["D1Total"] = Convert.ToInt32(Math.Ceiling(numberD1 / Y));
                            drowResult["D2Total"] = Convert.ToInt32(Math.Ceiling(numberD2 / Y));
                            dayModel.BNumber = splitNumber(cardModel, persoRid,
                                Convert.ToInt32(drowResult["BTotal"]), PersoCardtype);
                            dayModel.D1Number = splitNumber(cardModel, persoRid,
                                Convert.ToInt32(drowResult["D1Total"]), PersoCardtype);
                            dayModel.D2Number = splitNumber(cardModel, persoRid,
                                Convert.ToInt32(drowResult["D2Total"]), PersoCardtype);
                        }
                        if (changeFlag && dayModel.CDay.ToString("yyyy/MM") == month)
                        {
                            dayModel.ENumber = 0;
                        }
                        else
                        {
                            dayModel.ENumber = getChangeNumber(dtbChange, changedOverNum, cardModel, persoRid, dayModel.CDay, PersoCardtype);
                            changeFlag = true;
                            month = dayModel.CDay.ToString("yyyy/MM");
                        }
                        dayModel.CNumber = getCNumber(cardModel, persoRid, dayModel.CDay, YCode);
                        //F.已下單未貨數
                        dayModel.FNumber = getOderedNotStorckNum(dtbOrder, cardModel.RID, persoRid, dayModel.CDay);
                        int TotalXH = dayModel.BNumber + dayModel.CNumber + dayModel.D1Number + dayModel.D2Number + dayModel.ENumber;
                        //G.預估本月底庫存數 
                        dayModel.GNumber = dayModel.ANumber + dayModel.FNumber - TotalXH;
                        dayModel.G1Number = dayModel.GNumber;
                        setModelToDataRow(dayModel, TotalXH, drowResult);
                        result.Rows.Add(drowResult);
                        dayModel = null;
                        drowResult = null;                       
                        
                    }
                    #endregion
                    //根據第一步完成的數據計算后一步的數據                 
                    DataTable totalTable = math(result,N);
                    //result.Dispose();                   
                    inputDate(cardModel, persoRid, YCode, N, totalTable);
                    LogFactory.Write("卡种:" + cardModel.Name + " 厂商: " + persoRid + "结束", GlobalString.LogType.OpLogCategory);
                   
                    //處理每日監控作業占用系統資源太多影響其它批次作業的運行的問題 YangKun 2009/12/03 start
                    totalTable.Dispose();//釋放臨時數據
                    dtblAction.Dispose();
                    dtbCountNum.Dispose();
                    result.Dispose();
                    //處理每日監控作業占用系統資源太多影響其它批次作業的運行的問題 YangKun 2009/12/03 end
                  
                    
                }
            }
            catch (Exception ex) {
                LogFactory.Write("importForeData(" + Y.ToString() + "," + N.ToString() + ")" + ex.Message, GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            //dao.CloseConnection();
            //處理每日監控作業占用系統資源太多影響其它批次作業的運行的問題 YangKun 2009/12/03 start
            dtbPersoCardType.Dispose();
            dtbChange.Dispose();
            dtbOrder.Dispose();
            next60Day.Dispose();
            Mem = GC.GetTotalMemory(true);
            LogFactory.Write("每日監控作業：importForeData(" + Y.ToString() + "," + N.ToString() + ",已用內存：" + Mem.ToString() + ")結束", GlobalString.LogType.OpLogCategory);
            //處理每日監控作業占用系統資源太多影響其它批次作業的運行的問題 YangKun 2009/12/03 end
        }

        /// <summary>
        /// 無條件將數據進位到整數
        /// </summary>
        /// <param name="dec"></param>
        /// <returns></returns>
        private int decimal2Int(decimal dec)
        {
            if (dec > (int)dec)
            {
                return (int)dec + 1;
            }
            else
            {
                return (int)dec;
            }
        }

        /// <summary>
        /// 將數量按照卡种與廠商關係檔拆分
        /// </summary>
        /// <param name="cardModel"></param>
        /// <param name="persoRid"></param>
        /// <param name="totalNum"></param>
        /// <param name="PersoCardtype"></param>
        /// <returns></returns>
        private int splitNumber(CARD_TYPE cardModel, int persoRid, int totalNum, DataTable PersoCardtype)
        {
            int result = 0;
            //如果卡種與perso廠不存在任何關係,返回0
            DataRow[] drows = PersoCardtype.Select("CardType_RID='" + cardModel.RID + "'");
            try
            {
                if (drows.Length > 0)
                {
                    DataTable CardTable = PersoCardtype.Clone();
                    //DataRow drowNew = CardTable.NewRow();
                    foreach (DataRow drow in drows)
                    {
                        CardTable.Rows.Add(drow.ItemArray);
                    }
                    DataRow[] Type1Rows = CardTable.Select("Percentage_Number='1'", "Priority desc");
                    //如果卡種與Perso廠商存在比率分配,計算比率分配情況
                    if (Type1Rows.Length > 0)
                    {
                        int leftNumber = totalNum;
                        foreach (DataRow drow1 in Type1Rows)
                        {
                            decimal percent = Convert.ToDecimal(drow1["value"]) / 100M;
                            if (drow1["Priority"].ToString() == "1")
                                result = leftNumber;
                            else
                                result = Convert.ToInt32(Math.Floor(totalNum * percent));
                            if (Convert.ToInt32(drow1["Factory_RID"]) == persoRid)
                            {
                              return result;
                            }
                            leftNumber -= result;
                        }
                    }
                    DataRow[] Type2Rows = CardTable.Select("Percentage_Number='2'", "Priority");
                    //如果卡種與Perso廠商存在數量,按數量分配
                    if (Type2Rows.Length > 0)
                    {
                        int leftNumber = totalNum;
                        foreach (DataRow drow2 in Type2Rows)
                        {
                            if (leftNumber != 0)
                            {
                                if (drow2["value"].ToString() == "0")
                                    result = leftNumber;
                                else
                                    result = Math.Min(Convert.ToInt32(drow2["value"]), leftNumber);
                                if (Convert.ToInt32(drow2["Factory_RID"]) == persoRid)
                                {
                                   return result;
                                }
                                leftNumber -= result;
                            }

                        }
                    }
                    DataRow[] Type3Rows = CardTable.Select("Base_Special='1'");
                    //如果卡種只存在基本分配
                    if (Type3Rows.Length > 0)
                    {
                        if (Convert.ToInt32(Type3Rows[0]["Factory_RID"]) == persoRid)
                        {
                            return totalNum;
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("將數量按照卡种與廠商關係檔拆分, splitNumber報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return result;
        }



        private int getCNumber(CARD_TYPE cardModel, int persoRid, DateTime CDay, int YCode)
        {
            int cNumber = 0;
            try
            {
                dirValues.Clear();
                dirValues.Add("affinity", cardModel.AFFINITY);
                dirValues.Add("Photo", cardModel.PHOTO);
                dirValues.Add("type", cardModel.TYPE);
                dirValues.Add("perso_rid", persoRid);
                dirValues.Add("xType", YCode);
                dirValues.Add("cDay", CDay);
                Object oResult = dao.ExecuteScalar(SEL_DAYLY_DATA_CNumber, dirValues);
                if (oResult != null)
                {
                    cNumber = Convert.ToInt32(oResult);
                }               
            }
            catch (Exception ex)
            {
                LogFactory.Write("getCNumber報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            return cNumber;
        }

        private void inputDate(CARD_TYPE card, int persoRid, int Y, int N, DataTable totalTable)
        {
            try
            {
                dao.OpenConnection();
                //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                CardTypeManager ctm = new CardTypeManager();
                bool IsNetworkCard = ctm.isNetworkCard(card.Name);

                DateTime StartDay = Convert.ToDateTime(totalTable.Rows[0]["day"].ToString());
                DateTime EndDay = Convert.ToDateTime(totalTable.Rows[totalTable.Rows.Count - 1]["day"].ToString());
                dirValues.Clear();
                dirValues.Add("affinity", card.AFFINITY);
                dirValues.Add("Photo", card.PHOTO);
                dirValues.Add("type", card.TYPE);
                dirValues.Add("perso_rid", persoRid);
                dirValues.Add("xType", Y);
                dirValues.Add("StartDay", StartDay);
                dirValues.Add("EndDay", EndDay);
                DataSet ds = dao.GetList(SEL_ALL_DAYLY_DATA, dirValues);
                DataTable dtDaylyData = new DataTable();
                int DayCount = 0;
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    dtDaylyData = ds.Tables[0];
                    DayCount = dtDaylyData.Rows.Count;
                }

                //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                DAYLY_MONITOR dayModel = new DAYLY_MONITOR();
                dayModel.AFFINITY = card.AFFINITY;
                dayModel.PHOTO = card.PHOTO;
                dayModel.TYPE = card.TYPE;
                dayModel.Name = card.Name;
                dayModel.NType = N;
                dayModel.XType = Y;
                dayModel.Perso_Factory_Rid = persoRid;
                foreach (DataRow drow in totalTable.Rows)
                {
                    //dayModel.AFFINITY = card.AFFINITY;
                    //dayModel.PHOTO = card.PHOTO;
                    //dayModel.TYPE = card.TYPE;
                    //dayModel.Name = card.Name;
                    //dayModel.NType = N;
                    //dayModel.XType = Y;
                    dayModel.CDay = Convert.ToDateTime(drow["Day"].ToString());
                    dayModel.ANumber = toInt(drow["A"]);
                    dayModel.A1Number = toInt(drow["A1"]);
                    dayModel.BNumber = toInt(drow["B"]);
                    dayModel.CNumber = toInt(drow["C"]);
                    dayModel.D1Number = toInt(drow["D1"]);
                    dayModel.D2Number = toInt(drow["D2"]);
                    dayModel.ENumber = toInt(drow["E"]);
                    dayModel.FNumber = toInt(drow["F"]);
                    dayModel.GNumber = toInt(drow["G"]);
                    dayModel.G1Number = toInt(drow["G1"]);
                    dayModel.HNumber = Math.Round(Convert.ToDecimal(drow["H"]), 1);
                    dayModel.JNumber = toInt(drow["J"]);
                    dayModel.KNumber = toInt(drow["K"]);
                    dayModel.LNumber = Math.Round(Convert.ToDecimal(drow["L"]), 1);
                    //dayModel.Perso_Factory_Rid = persoRid;
                    //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                    //dirValues.Clear();
                    //dirValues.Add("affinity", dayModel.AFFINITY);
                    //dirValues.Add("Photo", dayModel.PHOTO);
                    //dirValues.Add("type", dayModel.TYPE);
                    //dirValues.Add("perso_rid", dayModel.Perso_Factory_Rid);
                    //dirValues.Add("xType", dayModel.XType);
                    //dirValues.Add("cDay", dayModel.CDay);

                    //Object oResult = dao.ExecuteScalar(SEL_DAYLY_DATA, dirValues);
                    //if (oResult == null)
                    // {
                    //    dao.Add<DAYLY_MONITOR>(dayModel, "RID");
                    //}
                    //else
                    //{
                    //    dayModel.RID = Convert.ToInt32(oResult);
                    //    dao.Update<DAYLY_MONITOR>(dayModel, "RID");
                    //}
                    //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                    if (DayCount > 0)
                    {
                        if (dtDaylyData.Select("CDay='" + dayModel.CDay.ToString() + "'").Length > 0)
                        {
                            DataRow drtemp = dtDaylyData.Select("CDay='" + dayModel.CDay.ToString() + "'")[0];
                            dayModel.RID = Convert.ToInt32(drtemp["RID"].ToString());
                            dao.Update<DAYLY_MONITOR>(dayModel, "RID");
                            drtemp = null;
                        }
                        else
                        {

                            dao.Add<DAYLY_MONITOR>(dayModel, "RID");
                        }
                    }
                    else
                    {

                        dao.Add<DAYLY_MONITOR>(dayModel, "RID");
                    }

                    if (drow["Day"] == totalTable.Rows[0]["Day"])
                    {
                        //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                        //CardTypeManager ctm = new CardTypeManager();
                        //if (!ctm.isNetworkCard(dayModel.Name))
                        //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                        if (!IsNetworkCard)
                        {
                            Warning.SetWarning(GlobalString.WarningType.DayMonitory, new object[3] { dayModel.Name, dayModel.NType, dayModel.HNumber });
                        }
                    }
                }
                dayModel = null;
                
                //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                dtDaylyData.Clear();
                dtDaylyData.Dispose();
                ds.Clear();
                ds.Dispose();
                //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                //事務提交
                dao.Commit();
              
                
            }
            catch (Exception ex)
            {
                dao.Rollback();
                LogFactory.Write("inputDate:" + ex.Message, GlobalString.LogType.ErrorCategory);
            }
            finally
            {
                dao.CloseConnection();
            }

        }
        //月換卡耗用量
        public const string SEL_NEXTMONTH_CHANGE_NUMBER = "SELECT ISNULL(SUM(PFCC.Number),0) FROM  FORE_CHANGE_CARD_DETAIL AS PFCC"
                                    + " WHERE PFCC.RST = 'A' "
                                    + " AND PFCC.Perso_Factory_RID = @PersoRid AND PFCC.Change_Date = @Change_Date "
                                    + " and PFCC.TYPE = @type and PFCC.PHOTO = @photo and PFCC.AFFINITY = @affinity ";
        /// <summary>
        /// E:換卡耗用量
        /// </summary>
        /// <param name="today"></param>
        /// <param name="dir"></param>
        /// <returns></returns>
        private int getChangeNumber(DataTable dtbChange, DataTable dtbSubtotal, CARD_TYPE cardModel, int persoRid, DateTime today, DataTable PersoCardtype)
        {
            try
            {
                int nextMonthChangeNum = 0;
                DataRow[] drows = dtbChange.Select("type='" + cardModel.TYPE + "' and affinity='" + cardModel.AFFINITY
                     + "' and photo='" + cardModel.PHOTO + "' and change_date ='" + today.AddMonths(1).ToString("yyyyMM")
                     //+ "' and Perso_factory_rid='" + persoRid.ToString() 
                     + "'");
                foreach (DataRow drow in drows)
                {
                    nextMonthChangeNum += Convert.ToInt32(drow["number"]);
                }

                int changedNum = 0;
                if (today >= Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/01")) && today <= Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/01")).AddMonths(1).AddDays(-1))
                {
                    DataRow[] drows1 = dtbSubtotal.Select("type='" + cardModel.TYPE + "' and affinity='" + cardModel.AFFINITY
                  + "' and photo='" + cardModel.PHOTO
                  //+ "' and Perso_factory_rid='" + persoRid.ToString() 
                  + "'");
                    foreach (DataRow drow1 in drows1)
                    {
                        changedNum += Convert.ToInt32(drow1["change_number"]);
                    }
                    drows1 = null;

                }
                
                if (nextMonthChangeNum != 0 && Convert.ToDecimal(changedNum) /Convert.ToDecimal(nextMonthChangeNum) < changePercent)
                {
                    int leftNumTotal = nextMonthChangeNum - changedNum;
                    return splitNumber(cardModel, persoRid, leftNumTotal, PersoCardtype); 
                }
                drows = null;
            }
            catch(Exception ex)
            {
                LogFactory.Write("getChangeNumber報錯:"+ex.Message, GlobalString.LogType.ErrorCategory);
                throw ex;
            }

            return 0;
        }

        public string paramList(string Paramcode)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("Param", GlobalString.ParameterType.CardParam);
                dirValues.Add("Param_code", Paramcode);
                StringBuilder sql = new StringBuilder(SEL_PARAM);
                sql.Append(" and Param_Code  = @Param_code");
                DataSet dstParam = dao.GetList(sql.ToString(), dirValues);
                return dstParam.Tables[0].Rows[0]["Param_Name"].ToString();
            }
            catch(Exception ex) {
                LogFactory.Write("paramList報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
                
                return "erro"; 
            }
        }

        /// <summary>
        /// 根據第一階段獲取的數據計算第二階段獲取的數據(A1,G1,H,J,K,L)
        /// </summary>
        /// <param name="drowResult"></param>
        /// <returns></returns>
        private DataTable math(DataTable dtbStep1, int NType)
        {
            DataTable dataTable2 = creatNewDataTable2();       
            //遍歷第一階段取得的數據表
            try
            {
                for (int j = 0; j < dtbStep1.Rows.Count; j++)
                {
                    DataRow drow2 = dataTable2.NewRow();
                    //複製第一階段的數據
                    foreach (DataColumn col in dtbStep1.Columns)
                    {
                        string colName = col.ColumnName;//取列名
                        if (colName != "BTotal" && colName != "D1Total" && colName != "D2Total")
                        drow2[colName] = dtbStep1.Rows[j][colName];
                    }
                    //計算A'的值
                    if (j > 0)
                    {
                        drow2["A1"] = dataTable2.Rows[j - 1]["K"];
                    }
                    //計算G'的值
                    if (j > 0)
                    {
                        drow2["G1"] = getG1Value(dtbStep1, j, drow2["A1"]);
                    }
                    //如果G值小於0,H值為0
                    if (toInt(dtbStep1.Rows[j]["G"]) < 0)
                    {
                        drow2["H"] = 0;
                    }
                    else
                    {
                        //計算H的值
                        drow2["H"] = getHValue(dtbStep1, j);
                    }
                    //2008/12/10: J值改爲不保存不計算
                    ////計算J值
                    //drow2["J"] = getJValue(dtbStep1, j, NType, drow2["G1"]);
                    drow2["J"] = 0;
                    //計算K值             
                    drow2["K"] = toInt(drow2["J"]) + toInt(drow2["G1"]);
                    //計算L檢核欄位的值
                    drow2["L"] = getLValue(drow2["G1"], drow2["J"], dtbStep1, j);
                    dataTable2.Rows.Add(drow2);
                }               
            }catch(Exception ex){
                LogFactory.Write("math報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
            }
            return dataTable2;
        }

        /// <summary>
        /// 創建第二階段的數據表
        /// </summary>
        /// <param name="dtblSafeStockInfo"></param>
        public DataTable creatNewDataTable2()
        {
            DataTable dtblSafeStockInfo = new DataTable();
            //日期
            dtblSafeStockInfo.Columns.Add("Day");
            //可用庫存量
            dtblSafeStockInfo.Columns.Add("A");
            //修正庫存量
            dtblSafeStockInfo.Columns.Add("A1");
            //過去Y月平均新卡數
            dtblSafeStockInfo.Columns.Add("B");
            //新進件調整
            dtblSafeStockInfo.Columns.Add("C");
            //過去Y月平均掛補數
            dtblSafeStockInfo.Columns.Add("D1");
            //過去Y月平均毀補數
            dtblSafeStockInfo.Columns.Add("D2");
            //該月換卡耗用量
            dtblSafeStockInfo.Columns.Add("E");
            //B+C+D1+D2+E
            dtblSafeStockInfo.Columns.Add("TotalXH");
            //已下單未到貨數
            dtblSafeStockInfo.Columns.Add("F");
            //預估月底庫存
            dtblSafeStockInfo.Columns.Add("G");
            //預估月底庫存
            dtblSafeStockInfo.Columns.Add("G1");
            dtblSafeStockInfo.Columns.Add("H");
            dtblSafeStockInfo.Columns.Add("J");
            dtblSafeStockInfo.Columns.Add("K");
            dtblSafeStockInfo.Columns.Add("L");
            return dtblSafeStockInfo;
        }

        /// <summary>
        /// 取L欄位的值
        /// </summary>
        /// <param name="G1"></param>
        /// <param name="J"></param>
        /// <param name="dtbStep1"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private decimal getLValue(object G1, object J, DataTable dtbStep1, int rowIndex)
        {

            decimal lValue = 0.0M;
            try
            {
                decimal G = Convert.ToDecimal(G1.ToString()) + Convert.ToDecimal(J.ToString());
                int checkRowIndex = rowIndex + 1;
                while (G > 0)
                {
                    //如果檢核的列超過最後一列,直接用當前庫存(G)除以最後一列的消耗數量(B+C+D1+D2+E)
                    if (checkRowIndex == dtbStep1.Rows.Count)
                    {
                        lValue = lValue + G / Convert.ToDecimal(dtbStep1.Rows[checkRowIndex - 1]["TotalXH"].ToString());
                        G = 0;
                    }
                    else
                    {
                        decimal totalXH = Convert.ToDecimal(dtbStep1.Rows[checkRowIndex]["TotalXH"].ToString());
                        if (G - totalXH > 0)
                            lValue += 1;//纍加檢核月數(H)
                        else
                            lValue = lValue + G / totalXH;//黨不滿一個月數的消耗時,加上計算滿足比率
                        G -= totalXH;
                    }
                    checkRowIndex++;
                }
            }
            catch
            {
                return 999999999.9M;
            }
            return lValue;
        }

        /// <summary>
        /// 取G'欄位的值
        /// </summary>
        /// <param name="dtbStep1"></param>
        /// <param name="rowIndex"></param>
        /// <param name="A1"></param>
        /// <returns></returns>
        private int getG1Value(DataTable dtbStep1, int rowIndex, object A1)
        {
            int G1Value = 0;
            try
            {
                int A1Value = toInt(A1);
                int FValue = toInt(dtbStep1.Rows[rowIndex]["F"]);
                int XHValue = toInt(dtbStep1.Rows[rowIndex]["TotalXH"]);
                G1Value = A1Value + FValue - XHValue;
            }
            catch
            {
                return 0;
            }
            return G1Value;
        }

        /// <summary>
        /// 將object 轉化成 int
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        private int toInt(object obj)
        {
            int result = 0;
            if (obj != null)
                result = int.Parse(obj.ToString());
            return result;
        }

        /// <summary>
        /// 取得系統建議or人員調整採購到貨量(J)
        /// </summary>
        /// <param name="dtbStep1">第一階段獲取數據</param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private int getJValue(DataTable dtbStep1, int rowIndex, int manageMonth, object G1)
        {
            int JValue = 0;
            try
            {
                for (int m = 1; m <= manageMonth; m++)
                {
                    JValue += int.Parse(dtbStep1.Rows[rowIndex + m]["TotalXH"].ToString());
                }
                JValue -= Convert.ToInt32(G1);
                if (JValue < 0)
                    JValue = 0;
            }
            catch
            {
                return 0;
            }
            return JValue;
        }

        /// <summary>
        /// 計算H檢核欄位的值
        /// </summary>
        /// <param name="dtbStep1"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        private decimal getHValue(DataTable dtbStep1, int rowIndex)
        {
            decimal hValue = 0.0M;
            decimal G = Convert.ToDecimal(dtbStep1.Rows[rowIndex]["G"].ToString());
            try
            {
                int checkRowIndex = rowIndex + 1;
                while (G > 0)
                {
                    //如果檢核的列超過最後一列,直接用當前庫存(G)除以最後一列的消耗數量(B+C+D1+D2+E)
                    if (checkRowIndex == dtbStep1.Rows.Count)
                    {
                        hValue = hValue + G / Convert.ToDecimal(dtbStep1.Rows[checkRowIndex - 1]["TotalXH"].ToString());
                        G = 0;
                    }
                    else
                    {
                        decimal totalXH = Convert.ToDecimal(dtbStep1.Rows[checkRowIndex]["TotalXH"].ToString());
                        if (G - totalXH > 0)
                            hValue += 1;//纍加檢核月數(H)
                        else
                            hValue = hValue + G / totalXH;//黨不滿一個月數的消耗時,加上計算滿足比率
                        G -= totalXH;
                    }
                    checkRowIndex++;
                }
            }
            catch
            {
                return 999999999.9M;
            }
            return hValue;
        }


        /// <summary>
        /// 將model轉化成dataRow
        /// </summary>
        /// <param name="model"></param>
        /// <param name="drow"></param>
        private void setModelToDataRow(DAYLY_MONITOR model, decimal TotalXH, DataRow drow)
        {
            drow["A"] = model.ANumber;
            drow["A1"] = model.A1Number;
            drow["B"] = model.BNumber;
            drow["C"] = model.CNumber;
            drow["D1"] = model.D1Number;
            drow["D2"] = model.D2Number;
            drow["E"] = model.ENumber;
            drow["TotalXH"] = TotalXH;
            drow["F"] = model.FNumber;
            drow["G"] = model.GNumber;
            drow["G1"] = model.G1Number;
            drow["Day"] = model.CDay;
        }


        //已下單未到貨數
        public const string SEL_NOT_INCOME_NUMBER = "SELECT (SELECT ISNULL(SUM(Number), 0)  FROM ORDER_FORM_DETAIL as ofd"
                                + " LEFT JOIN  ORDER_FORM ON ORDER_FORM.OrderForm_RID = ofd.OrderForm_RID"
                                + " WHERE (ORDER_FORM.Pass_Status = 4)"
                                + " AND (ofd.CardType_RID = @cardRid) AND (ofd.Delivery_Address_RID = @PersoRid)  and ofd.fore_delivery_date =@checkDate ) -"
                                + " (SELECT ISNULL(SUM(Stock_Number), 0) FROM  DEPOSITORY_STOCK"
                                + " where DEPOSITORY_STOCK.orderform_detail_rid in (select ord.orderform_detail_rid "
                                + " from ORDER_FORM_DETAIL  as ord "
                                + " LEFT JOIN  ORDER_FORM ON ORDER_FORM.OrderForm_RID =ord.OrderForm_RID "
                                + " WHERE (ORDER_FORM.Pass_Status = 4)"
                                + " and (ord.CardType_RID = @cardRid) and (ord.Delivery_Address_RID = @PersoRid) and ord.fore_delivery_date=@checkDate)"
                                + " and income_date < @checkDate ) AS notIncomNum";
        /// <summary>
        /// 當天已下單未到貨數
        /// </summary>
        /// <param name="drow"></param>
        /// <returns></returns>
        private int getOderedNotStorckNum(DataTable dtbOrder, int cardRid, int persoRid, DateTime checkDate)
        {
            int OderedNotStorckNum = 0;
            try
            {
                DataRow[] drows = dtbOrder.Select("cardtype_rid='" + cardRid.ToString()
             + "' and Delivery_Address_RID ='" + persoRid.ToString() +
             "' and fore_delivery_date='" + checkDate.ToString ("yyyy/MM/dd") + "'");
                if (drows.Length == 1)
                {
                    int orderNumber = 0;
                    int stockNumber = 0;
                    if (!StringUtil.IsEmpty(drows[0]["number"].ToString()))
                        orderNumber = Convert.ToInt32(drows[0]["number"]);
                    if (!StringUtil.IsEmpty(drows[0]["stock_Number"].ToString()))
                        stockNumber = Convert.ToInt32(drows[0]["stock_Number"]);
                    OderedNotStorckNum = orderNumber - stockNumber;
                }
            }
            catch (Exception ex)
            {
                LogFactory.Write("getOderedNotStorckNum報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
            }
            return OderedNotStorckNum;
        }

        private DataTable  creatNewDataTable()
        {
            DataTable dtblSafeStockInfo = new DataTable();
            //日期
            dtblSafeStockInfo.Columns.Add("Day");
            //可用庫存量
            dtblSafeStockInfo.Columns.Add("A");
            //修正庫存量
            dtblSafeStockInfo.Columns.Add("A1");
            //過去X月平均新卡數
            dtblSafeStockInfo.Columns.Add("B");
            dtblSafeStockInfo.Columns.Add("BTotal");
            //新進件調整
            dtblSafeStockInfo.Columns.Add("C");
            //過去X月平均掛補數
            dtblSafeStockInfo.Columns.Add("D1");
            dtblSafeStockInfo.Columns.Add("D1Total");
            //過去X月平均毀補數
            dtblSafeStockInfo.Columns.Add("D2");
            dtblSafeStockInfo.Columns.Add("D2Total");
            //該月換卡耗用量
            dtblSafeStockInfo.Columns.Add("E");
            //B+C+D1+D2+E
            dtblSafeStockInfo.Columns.Add("TotalXH");
            //已下單未到貨數
            dtblSafeStockInfo.Columns.Add("F");
            //預估月底庫存
            dtblSafeStockInfo.Columns.Add("G");
            //預估月底庫存
            dtblSafeStockInfo.Columns.Add("G1");

            return dtblSafeStockInfo;
        }

        /// <summary>
        /// 根據日結時間，廠商，卡種取庫存數量
        /// </summary>
        /// <param name="cardTypeRid"></param>
        /// <param name="persoRid"></param>
        /// <param name="WorkDay"></param>
        /// <returns></returns>
        public int getStockNumByCheckDay(int persoRid, int cardTypeRid, DateTime CheckDay)
        {
            //DataTable dtblResult = null;
            dirValues.Clear();
            dirValues.Add("persoRid", persoRid);
            dirValues.Add("cardRid", cardTypeRid);
            dirValues.Add("stockDate", CheckDay);

            Object oResult = dao.ExecuteScalar(SEL_STOCKS_NUM, dirValues);

            if ( oResult != null )
            {
                return Convert.ToInt32(oResult);
            }
            return 0;
        }

        /// <summary>
        /// 獲取輸入時間的當前庫存
        /// </summary>
        /// <param name="persoRid"></param>
        /// <param name="cardTypeRid"></param>
        /// <param name="currentDay"></param>
        /// <returns></returns>
        public int getCurrentStock(int persoRid, int cardTypeRid, DateTime currentDay)
        {
            int result = 0;
            try
            {
                if (isCheckDate(currentDay))
                {
                    return getStockNumByCheckDay(persoRid, cardTypeRid, currentDay);
                }
                DataRow  drStock = getCheckStockByPerso(persoRid, cardTypeRid);
                DateTime checkDate = DateTime.Now;

                if (drStock != null)
                {
                    checkDate = Convert.ToDateTime(drStock["Stock_Date"].ToString());
                    result = Convert.ToInt32(drStock["Stocks_Number"].ToString());
                }
                dirValues.Clear();
                dirValues.Add("persoRid", persoRid.ToString());
                dirValues.Add("cardRid", cardTypeRid.ToString());
                dirValues.Add("Stock_Date", checkDate);
                dirValues.Add("today", currentDay);
                //加上所有入庫記錄
                result += Convert.ToInt32(dao.ExecuteScalar(SEL_DEPOSITORY_STOCK, dirValues));
                //減去所有退貨記錄
                result -= Convert.ToInt32(dao.ExecuteScalar(SEL_DEPOSITORY_CANCEl_NUMBER, dirValues));
                //加上所有再入庫記錄
                result += Convert.ToInt32(dao.ExecuteScalar(SEL_DEPOSITORY_RESTOCK, dirValues));
                //加上所有移轉進入的紀錄
                result += Convert.ToInt32(dao.ExecuteScalar(CARDTYPE_STOCKS_MOVE_IN, dirValues));
                //減去所有移出記錄
                result -= Convert.ToInt32(dao.ExecuteScalar(CARDTYPE_STOCKS_MOVE_OUT, dirValues));
                //減去消耗的卡片數量
                result -= getUseCardNum(persoRid.ToString(), cardTypeRid.ToString(), checkDate, currentDay);              
            }catch(Exception ex){
                LogFactory.Write("getCurrentStock報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            return result;
        }

        /// <summary>
        /// 
        /// 根據PERSO厰RID和卡种RID查詢日結庫存量
        /// </summary>
        /// <param name="factoryRid"></param>
        /// <param name="cardTypeRid"></param>
        public DataRow  getCheckStockByPerso(int factoryRid, int cardTypeRid)
        {            
            dirValues.Clear();
            dirValues.Add("persoRid", factoryRid.ToString());
            dirValues.Add("cardRid", cardTypeRid.ToString());
            try
            {
              DataRow     drResult = dao.GetRow(SEL_STOCKS, dirValues,false );
              return drResult;
            }
            catch (Exception ex)
            {
                LogFactory.Write("getCheckStockByPerso報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
                return null;
            }
        }

        /// <summary>
        /// 取消耗卡數量
        /// </summary>
        /// <param name="factoryRID"></param>
        /// <param name="cardRID"></param>
        /// <param name="storkDate"></param>
        /// <returns></returns>
        public int getUseCardNum(String factoryRID, string cardRID, DateTime storkDate, DateTime currentDay)
        {
            int result = 0;
            //查詢卡种基本信息
            CARD_TYPE cardTypeModel = dao.GetModel<CARD_TYPE, int>("RID", int.Parse(cardRID));
            dirValues.Clear();
            dirValues.Add("TYPE", cardTypeModel.TYPE);
            dirValues.Add("PHOTO", cardTypeModel.PHOTO);
            dirValues.Add("AFFINITY", cardTypeModel.AFFINITY);
            dirValues.Add("Perso_Factory_RID", factoryRID);
            dirValues.Add("Stock_Date", storkDate);
            dirValues.Add("today", currentDay);
            //查詢廠商異動信息
            DataSet factorySet = dao.GetList(SEL_FACTORY_CHANGE_NUM, dirValues);
            DataTable factoryTable = null;
            if (factorySet.Tables.Count > 0)
            {
                factoryTable = factorySet.Tables[0];
            }
            else
            {
                return result;
            }
            dirValues.Clear();
            dirValues.Add("Expressions_RID", GlobalString.Expression.Used_RID);
            //查詢消耗量公式信息
            DataTable expTable = dao.GetList(SEL_EXPRESSION_INFO, dirValues).Tables[0];
            //為提高系統性能默認小計檔下來全部為消耗卡，即公式中3D+DA+PM+RN=小計檔總合
            CardTypeManager manager = new CardTypeManager();
            //int subUseNum = manager.getUseCardNumSub(factoryRID, cardRID, storkDate.AddDays(1), currentDay.AddDays(-1));
            int subUseNum = manager.getUseCardNumSub(factoryRID, cardRID, storkDate, currentDay.AddDays(-1));
            result += subUseNum;

            //按照公式計算廠商異動帶來的消耗數量
            foreach (DataRow drowExp in expTable.Rows)
            {
                //查詢該卡种狀態是否應該在公式中出現，如果不出現，則ＣＯＮＴＩＮＵＥ
                int typeRid = Convert.ToInt32(drowExp["Type_RID"]);
                string Operate = drowExp["Operate"].ToString();
                CARDTYPE_STATUS statusModel = dao.GetModel<CARDTYPE_STATUS, int>("RID", typeRid);
                if (statusModel.Is_Display.Equals(GlobalString.YNType.No))
                {
                    continue;
                }
                int SUM = 0;
                DataRow[] rows = factoryTable.Select("Status_RID='" + typeRid + "'");
                foreach (DataRow drowFac in rows)
                {
                    SUM += Convert.ToInt32(drowFac["Number"]);
                }
                //當前卡种狀態為+
                if (Operate == GlobalString.Operation.Add_RID)
                {
                    result += SUM;
                }
                //當前卡种狀態為-
                if (Operate == GlobalString.Operation.Del_RID)
                {
                    result -= SUM;
                }
            }
            expTable.Dispose();//釋放占用資源
            factoryTable.Dispose();
            factorySet.Dispose();
            manager = null;
            return result;
        }
       
        /// <summary>
        /// 判斷是否為日結日
        /// </summary>
        /// <param name="strBudgetID">預算簽呈ID</param>
        /// <returns>true:存在 false:不存在</returns>
        public bool isCheckDate(DateTime CheckDate)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("CheckDate", CheckDate);
                return dao.Contains(CON_CHECK_DATE, dirValues);
            }
            catch (Exception ex)
            {
                LogFactory.Write("isCheckDate報錯:"+ex.Message, GlobalString.LogType.ErrorCategory);
            }
            return false;
        }

    }
}
