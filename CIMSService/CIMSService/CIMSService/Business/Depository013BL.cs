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
    class Depository013BL : BaseLogic
    {
        #region SQL定義

        //查詢最後日結日
        public const string SEL_STOCK_DATE = "select max(stock_date) from CARDTYPE_STOCKS";

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
                                    + " LEFT JOIN CARD_TYPE AS type ON stocks.CardType_RID = type.RID"
                                    + " LEFT JOIN FACTORY AS factory ON stocks.Perso_Factory_RID = factory.RID"
                                    + " where stocks.rst = 'A' and stock_date = @Stock_Date";

        public const string SEL_WORKDAY = "Select Max(Date_Time) FROM WORK_DATE where RST = 'A' AND Is_WorkDay = 'Y' AND RST = 'A' AND Is_WorkDay = 'Y' AND Date_Time < @Month";
        //月換卡耗用量
        public const string SEL_NEXTMONTH_CHANGE_NUMBER = "SELECT ISNULL(SUM(PFCC.Number),0) FROM  FORE_CHANGE_CARD_DETAIL AS PFCC"
                                    + " WHERE PFCC.RST = 'A' "
                                    + " AND PFCC.Perso_Factory_RID = @PersoRid AND PFCC.Change_Date = @Change_Date "
                                    + " and PFCC.TYPE = @type and PFCC.PHOTO = @photo and PFCC.AFFINITY = @affinity ";
                
        //取開始時間到結束時間的Action= @action 的小計數量
        public const string SEL_PREMONTHS_NUMBER = "SELECT ISNULL(sum(SI.Number),0) FROM SUBTOTAL_IMPORT AS SI "
                                    + " WHERE SI.RST = 'A' "
                                   // + " AND SI.Perso_Factory_RID = @PersoRid "
                                    + " and SI.Date_Time >= @begin_date and SI.Date_Time < @end_date and SI.action = @action"
                                    + " and SI.PHOTO = @photo and SI.TYPE = @type and SI.AFFINITY = @affinity ";

        //已下單未到貨數

        public const string SEL_NOT_INCOME_NUMBER = "SELECT (SELECT ISNULL(SUM(Number), 0)  FROM ORDER_FORM_DETAIL as ofd"
                                    + " LEFT JOIN  ORDER_FORM ON ORDER_FORM.OrderForm_RID = ofd.OrderForm_RID"
                                    + " WHERE (ORDER_FORM.Pass_Status = 4)"
                                    + " AND (ofd.CardType_RID = @cardRid) AND (ofd.Delivery_Address_RID = @PersoRid) and ofd.fore_delivery_date >= @beginDate and ofd.fore_delivery_date < @endDate ) -"
                                    + " (SELECT ISNULL(SUM(Stock_Number), 0) FROM  DEPOSITORY_STOCK"
                                    + " where DEPOSITORY_STOCK.orderform_detail_rid in (select ord.orderform_detail_rid   "
                                    + " from ORDER_FORM_DETAIL  as ord "
                                    + " LEFT JOIN  ORDER_FORM ON ORDER_FORM.OrderForm_RID =ord.OrderForm_RID "
                                    + " WHERE (ORDER_FORM.Pass_Status = 4)"
                                    + " and (ord.CardType_RID = @cardRid) AND (ord.Delivery_Address_RID = @PersoRid) and ord.fore_delivery_date >= @beginDate and ord.fore_delivery_date < @endDate )"
                                    + " and income_date < @beginDate ) AS notIncomNum";

        public const string SEL_PARAM = "SELECT  Param_Name FROM PARAM WHERE ParamType_Code = @Param ";

        public const string CON_SAFE_DATA = "SELECT count(*) from MONTHLY_MONITOR where Perso_factory_RID = @perso_rid "
                                    +" and cmonth = @cMonth and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType ";

        public const string SEL_SAFE_DATA = "SELECT * from MONTHLY_MONITOR where Perso_factory_RID = @perso_rid "
                                    + " and cmonth = @cMonth and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType ";
        public const string SEL_ALL_SAFE_DATA = "SELECT RID,Cmonth from MONTHLY_MONITOR where Perso_factory_RID = @perso_rid "
                                   + " and cmonth >= @StartMonth and cmonth <= @EndMonth and TYPE = @type and PHOTO = @photo and AFFINITY = @affinity and XType = @xType ";

        public const string DEL_MONTH_DATA = "delete FROM  MONTHLY_MONITOR "
                                    + " where rid not in (select MONTHLY_MONITOR.rid from MONTHLY_MONITOR"
                                    + " left join card_type on card_type.[TYPE]=MONTHLY_MONITOR.[TYPE]"
                                    + " AND card_type.[AFFINITY]=MONTHLY_MONITOR.[AFFINITY]"
                                    + " AND card_type.[PHOTO]=MONTHLY_MONITOR.[PHOTO]"
                                    + " inner join CARDTYPE_STOCKS"
                                    + " on CARDTYPE_STOCKS.perso_factory_rid =MONTHLY_MONITOR.perso_factory_rid"
                                    + " and card_type.rid= CARDTYPE_STOCKS.cardtype_rid"
                                           + " WHERE  CARDTYPE_STOCKS.stock_date = (select Max(stock_date) from cardtype_stocks))";


        //public const string SEL_PERSO_CARDTYPE = "select pc.* from  PERSO_CARDTYPE pc inner join cardtype_stocks cs on pc.cardtype_rid=cs.cardtype_rid and stock_date=@stockDate";
        public const string SEL_PERSO_CARDTYPE = "select pc.* from  PERSO_CARDTYPE pc ";
       
        #endregion

        Dictionary<string, object> dirValues = new Dictionary<string, object>();

        //查詢各廠商各卡種的資料
        public DataSet getStockCardType(DateTime LastCheckDate)
        {
            //先取出系統的最後日結日
            dirValues.Clear();
            //執行SQL語句
            dirValues.Add("Stock_Date",LastCheckDate );
            DataSet dstSafeStockInfo = dao.GetList(SEL_PRECARD,dirValues );          
            return dstSafeStockInfo;
        }

        /// <summary>
        /// 主程序入口
        /// </summary>
        public void ComputeForeData() {

            try   // 加try catch add by judy 2018/03/28
            {
                string N = paramList(GlobalString.cardparamType.NType);
                string X1 = paramList(GlobalString.cardparamType.X1);
                string X2 = paramList(GlobalString.cardparamType.X2);
                string X3 = paramList(GlobalString.cardparamType.X3);
                if (N != "erro" && X1 != "erro" && X2 != "erro" && X3 != "erro")
                {
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
                    DataTable PersoCardtype = getPersoCardtype(LastCheckDate);
                    importForeData(X1, N, GlobalString.cardparamType.X1, LastCheckDate, PersoCardtype);
                    importForeData(X2, N, GlobalString.cardparamType.X2, LastCheckDate, PersoCardtype);
                    importForeData(X3, N, GlobalString.cardparamType.X3, LastCheckDate, PersoCardtype);
                }
            }
            catch(Exception ex)
            {
                LogFactory.Write("主程序入口ComputeForeData方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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

            try {
                dirValues.Clear();
                dirValues.Add("stockDate", LastCheckDate);
                dtb= dao.GetList(SEL_PERSO_CARDTYPE, dirValues).Tables[0];            
            }
            catch(Exception ex){
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("查詢最近一次日結卡種的所有廠商分配比例, getPersoCardtype報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return dtb;        
        }

        //存預計結果的數據
        public void importForeData(string strX,string strN,string paramCode,DateTime LastCheckDate,DataTable PersoCardtype)
        {
            DataTable dtbPersoCardType = getStockCardType(LastCheckDate).Tables[0];
            int N = int.Parse(strN.Substring(0, strN.Length - 1));
            int X = int.Parse(strX.Substring(0, strX.Length - 1));
            int XCode = int.Parse(paramCode);
            long Mem = GC.GetTotalMemory(true);//已用內存
            LogFactory.Write("每月監控作業：importForeData(" + X.ToString() + "," + N.ToString() + ",已用內存：" + Mem.ToString() + ")開始", GlobalString.LogType.OpLogCategory);
            
            foreach (DataRow drow in dtbPersoCardType.Rows)
            {
                //取卡種信息
                CARD_TYPE cardModel = new CARD_TYPE();
                cardModel.AFFINITY = drow["AFFINITY"].ToString();
                cardModel.PHOTO = drow["Photo"].ToString(); //Photo
                cardModel.TYPE = drow["TYPE"].ToString();//card Type
                cardModel.RID = int.Parse(drow["CardType_RID"].ToString());
                cardModel.Name = drow["Name"].ToString();
                //取廠商信息
                int persoRid = int.Parse(drow["Perso_Factory_RID"].ToString());
                LogFactory.Write("卡种:" + cardModel.Name + " 厂商: " + persoRid + "开始", GlobalString.LogType.OpLogCategory);

                //取預估開始月份
                DateTime thisMonthBeginDay = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-01"));
                //結果集
                DataTable result = new DataTable();
                creatNewDataTable(result);
                #region 計算數據
                for (int i = 0; i < 16; i++)
                {
                    DataRow drowResult = result.NewRow();
                    MONTHLY_MONITOR monthModel = new MONTHLY_MONITOR();
                    if (i == 0)
                    {
                        dirValues.Clear();
                        dirValues.Add("Month", thisMonthBeginDay);
                        //取上個月最後一個工作日
                        DateTime lastMonthWorkDay = Convert.ToDateTime(dao.GetList(SEL_WORKDAY, dirValues).Tables[0].Rows[0][0].ToString());
                        monthModel.ANumber = getCurrentStock(persoRid, cardModel.RID, lastMonthWorkDay);
                        monthModel.A1Number = monthModel.ANumber;
                    }
                    else
                    {
                        monthModel.ANumber = int.Parse(result.Rows[i - 1]["G"].ToString());
                    }
                    //取前X個月的平均值
                    if (i >= X)
                    {
                        double numberB = 0D;
                        double numberD1 = 0D;
                        double numberD2 = 0D;
                        for (int j = 1; j <= X; j++)
                        {
                            numberB += Convert.ToDouble(result.Rows[i - j]["BTotal"].ToString());
                            numberD1 += Convert.ToDouble(result.Rows[i - j]["D1Total"].ToString());
                            numberD2 += Convert.ToDouble(result.Rows[i - j]["D2Total"].ToString());
                        }
                        drowResult["BTotal"] = Convert.ToInt32(Math.Ceiling(numberB / X));
                        drowResult["D1Total"] = Convert.ToInt32(Math.Ceiling(numberD1 / X));
                        drowResult["D2Total"] = Convert.ToInt32(Math.Ceiling(numberD2 / X));
                        monthModel.BNumber = splitNumber(cardModel, persoRid, 
                            Convert.ToInt32(drowResult["BTotal"]), PersoCardtype);
                        monthModel.D1Number = splitNumber(cardModel, persoRid, 
                            Convert.ToInt32(drowResult["D1Total"]), PersoCardtype);
                        monthModel.D2Number = splitNumber(cardModel, persoRid, 
                            Convert.ToInt32(drowResult["D2Total"]), PersoCardtype);
                    }
                    else
                    {
                        DateTime startMonth = thisMonthBeginDay.AddMonths(i - X);
                        //B.過去X月平均新卡數
                        double numberB = getNumberFromSubTotal(cardModel, persoRid, startMonth, thisMonthBeginDay, "1");
                        //D1.預估每月掛補量
                        double numberD1 = getNumberFromSubTotal(cardModel, persoRid, startMonth, thisMonthBeginDay, "2");
                        //D2.預估每月毀補量
                        double numberD2 = getNumberFromSubTotal(cardModel, persoRid, startMonth, thisMonthBeginDay, "3");
                        for (int j = 0; j < i; j++)
                        {
                            numberB += Convert.ToDouble(result.Rows[j]["BTotal"].ToString());
                            numberD1 += Convert.ToDouble(result.Rows[j]["D1Total"].ToString());
                            numberD2 += Convert.ToDouble(result.Rows[j]["D2Total"].ToString());
                        }
                        drowResult["BTotal"] = Convert.ToInt32(Math.Ceiling(numberB / X));
                        drowResult["D1Total"] = Convert.ToInt32(Math.Ceiling(numberD1 / X));
                        drowResult["D2Total"] = Convert.ToInt32(Math.Ceiling(numberD2 / X));
                        monthModel.BNumber = splitNumber(cardModel, persoRid,
                            Convert.ToInt32(drowResult["BTotal"]), PersoCardtype);
                        monthModel.D1Number = splitNumber(cardModel, persoRid,
                            Convert.ToInt32(drowResult["D1Total"]), PersoCardtype);
                        monthModel.D2Number = splitNumber(cardModel, persoRid,
                            Convert.ToInt32(drowResult["D2Total"]), PersoCardtype);
                    }
                    monthModel.CNumber = getCNumber(cardModel, persoRid, thisMonthBeginDay.AddMonths(i), XCode);
                    //E.該月換卡耗用量
                    monthModel.ENumber = getNextMonthChangeNumber(cardModel, persoRid, thisMonthBeginDay.AddMonths(i));
                    //F.已下單預計本月到貨數
                    monthModel.FNumber = getOderedNotStorckNum(cardModel.RID, persoRid, thisMonthBeginDay.AddMonths(i), thisMonthBeginDay.AddMonths(i + 1));
                    int TotalXH = monthModel.BNumber + monthModel.CNumber + monthModel.D1Number + monthModel.D2Number + monthModel.ENumber;
                    //G.預估本月底庫存數 
                    monthModel.GNumber = monthModel.ANumber + monthModel.FNumber - TotalXH;
                    monthModel.G1Number = monthModel.GNumber;
                    monthModel.CMonth = thisMonthBeginDay.AddMonths(i);
                    setModelToDataRow(monthModel, TotalXH, drowResult);
                    result.Rows.Add(drowResult);
                    //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                    drowResult = null;
                    monthModel = null;
                    //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                }
                #endregion
                //根據第一步完成的數據計算后一步的數據
                DataTable totalTable = math(result, N);
                inputDate(cardModel, persoRid, XCode, N, totalTable);
                //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                result.Dispose();
                totalTable.Dispose();
                //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                LogFactory.Write("卡种:" + cardModel.Name + " 厂商: " + persoRid + "結束", GlobalString.LogType.OpLogCategory);

            }
            //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
            dtbPersoCardType.Dispose();
            Mem = GC.GetTotalMemory(true);
            LogFactory.Write("每月監控作業：importForeData(" + X.ToString() + "," + N.ToString() + ",已用內存：" + Mem.ToString() + ")結束", GlobalString.LogType.OpLogCategory);
            //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
        }

        /// <summary>
        /// 將數量按照卡种與廠商關係檔拆分
        /// </summary>
        /// <param name="cardModel"></param>
        /// <param name="persoRid"></param>
        /// <param name="totalNum"></param>
        /// <param name="PersoCardtype"></param>
        /// <returns></returns>
        private int splitNumber(CARD_TYPE cardModel,int persoRid,int totalNum,DataTable PersoCardtype)
        {
            int result = 0;
            //如果卡種與perso廠不存在任何關係,返回0
            DataRow[] drows = PersoCardtype.Select("CardType_RID='" + cardModel.RID+"'");
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

        private int getCNumber(CARD_TYPE cardModel, int persoRid, DateTime month, int xCode)
        {
            int cNumber = 0;
            dirValues.Clear();
            dirValues.Add("affinity", cardModel.AFFINITY);
            dirValues.Add("Photo", cardModel.PHOTO);
            dirValues.Add("type", cardModel.TYPE);
            dirValues.Add("perso_rid", persoRid);
            dirValues.Add("xType", xCode);
            dirValues.Add("cMonth", month);
            DataSet dstResult = dao.GetList(SEL_SAFE_DATA, dirValues);
            if (dstResult.Tables[0] != null && dstResult.Tables[0].Rows.Count > 0)
            {
                cNumber = Convert.ToInt32(dstResult.Tables[0].Rows[0]["CNumber"]);
            }
            return cNumber;        
        }

        private void inputDate(CARD_TYPE card, int persoRid,int X,int N, DataTable totalTable)
        {

            try
            {
                dao.OpenConnection();
                //處理每月監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 start
                CardTypeManager ctm = new CardTypeManager();
                bool IsNetworkCard = ctm.isNetworkCard(card.Name);

                DateTime StartMonth = Convert.ToDateTime(totalTable.Rows[0]["Month"].ToString());
                DateTime EndMonth = Convert.ToDateTime(totalTable.Rows[totalTable.Rows.Count - 1]["Month"].ToString());
                dirValues.Clear();
                dirValues.Add("affinity", card.AFFINITY);
                dirValues.Add("Photo", card.PHOTO);
                dirValues.Add("type", card.TYPE);
                dirValues.Add("perso_rid", persoRid);
                dirValues.Add("xType", X);
                dirValues.Add("StartMonth", StartMonth);
                dirValues.Add("EndMonth", EndMonth);

                DataSet ds = dao.GetList(SEL_ALL_SAFE_DATA, dirValues);
                DataTable dtMonth = new DataTable();
                int MonthCount = 0;
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    dtMonth = ds.Tables[0];
                    MonthCount = dtMonth.Rows.Count;
                }

                dao.ExecuteNonQuery(DEL_MONTH_DATA);
                MONTHLY_MONITOR monthModel = new MONTHLY_MONITOR();
                monthModel.AFFINITY = card.AFFINITY;
                monthModel.PHOTO = card.PHOTO;
                monthModel.TYPE = card.TYPE;
                monthModel.Name = card.Name;
                monthModel.NType = N;
                monthModel.XType = X;
                monthModel.Perso_Factory_Rid = persoRid;
                foreach (DataRow drow in totalTable.Rows)
                {
                    //MONTHLY_MONITOR monthModel = new MONTHLY_MONITOR();
                    //monthModel.AFFINITY = card.AFFINITY;
                    //monthModel.PHOTO = card.PHOTO;
                    //monthModel.TYPE = card.TYPE;
                    //monthModel.Name = card.Name;                
                    //monthModel.NType = N;
                    //monthModel.XType = X;
                    monthModel.CMonth = Convert.ToDateTime(drow["Month"].ToString());
                    monthModel.ANumber = toInt(drow["A"]);
                    monthModel.A1Number = toInt(drow["A1"]);
                    monthModel.BNumber = toInt(drow["B"]);
                    monthModel.CNumber = toInt(drow["C"]);
                    monthModel.D1Number = toInt(drow["D1"]);
                    monthModel.D2Number = toInt(drow["D2"]);
                    monthModel.ENumber = toInt(drow["E"]);
                    monthModel.FNumber = toInt(drow["F"]);
                    monthModel.GNumber = toInt(drow["G"]);
                    monthModel.G1Number = toInt(drow["G1"]);
                    monthModel.HNumber = Math.Round(Convert.ToDecimal(drow["H"]), 1);
                    monthModel.JNumber = toInt(drow["J"]);
                    monthModel.KNumber = toInt(drow["K"]);
                    monthModel.LNumber = Math.Round(Convert.ToDecimal(drow["L"]), 1);

                    //monthModel.Perso_Factory_Rid = persoRid;
                    //dirValues.Clear();
                    //dirValues.Add("affinity", monthModel.AFFINITY);
                    //dirValues.Add("Photo", monthModel.PHOTO);
                    //dirValues.Add("type", monthModel.TYPE);
                    //dirValues.Add("perso_rid", monthModel.Perso_Factory_Rid);
                    //dirValues.Add("xType", monthModel.XType);
                    //dirValues.Add("cMonth", monthModel.CMonth);
                    //if (dao.Contains(CON_SAFE_DATA, dirValues))
                    //{

                    //    int rid = Convert.ToInt32(dao.GetList(SEL_SAFE_DATA, dirValues).Tables[0].Rows[0]["Rid"]);
                    //    monthModel.RID = rid;
                    //    dao.Update<MONTHLY_MONITOR>(monthModel, "RID");

                    //}
                    if (MonthCount > 0)
                    {
                        if (dtMonth.Select("CMonth='" + monthModel.CMonth.ToString() + "'").Length > 0)
                        {
                            DataRow drtemp = dtMonth.Select("CMonth='" + monthModel.CMonth.ToString() + "'")[0];
                            monthModel.RID = Convert.ToInt32(drtemp["RID"].ToString());
                            dao.Update<MONTHLY_MONITOR>(monthModel, "RID");
                            drtemp = null;
                        }
                        else
                        {
                            dao.Add<MONTHLY_MONITOR>(monthModel, "RID");
                        }
                    }
                    else
                    {
                        dao.Add<MONTHLY_MONITOR>(monthModel, "RID");
                    }

                    if (monthModel.CMonth.ToString("yyyy/MM/01") == DateTime.Now.ToString("yyyy/MM/01"))
                    {
                        //CardTypeManager ctm = new CardTypeManager();
                        //if (!ctm.isNetworkCard(monthModel.Name))
                        if (!IsNetworkCard)
                        {
                            Warning.SetWarning(GlobalString.WarningType.MonthMonitory, new object[3] { monthModel.Name, monthModel.NType, monthModel.HNumber });
                        }
                    }
                }
                monthModel = null;

                dtMonth.Clear();
                dtMonth.Dispose();
                ds.Clear();
                ds.Dispose();
                //處理每日監控作業每條記錄處理時間太長的問題 YangKun 2009/12/04 end
                //事務提交
                dao.Commit();

            }
            catch (Exception ex)
            {
                dao.Rollback();
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("inputDate報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);

                // BatchBL Bbl = new BatchBL();
                // Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);

            }
            finally
            {
                dao.CloseConnection();
            }
        }

        /// <summary>
        /// 月換卡耗用量
        /// </summary>
        /// <param name="today"></param>
        /// <param name="dir"></param>
        /// <returns></returns>
        private int getNextMonthChangeNumber(CARD_TYPE model,int persoRid , DateTime today)
        {
            int result = 0;
            dirValues.Clear();
            dirValues.Add("PersoRid", persoRid);
            dirValues.Add("photo", model.PHOTO);
            dirValues.Add("type", model.TYPE);
            dirValues.Add("affinity", model.AFFINITY);
            dirValues.Add("Change_Date", today.AddMonths(1).ToString("yyyyMM"));
            result = int.Parse(dao.GetList(SEL_NEXTMONTH_CHANGE_NUMBER, dirValues).Tables[0].Rows[0][0].ToString());
            return result;
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
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("paramList報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return "erro";            
            }
        }

        /// <summary>
        /// 根據第一階段獲取的數據計算第二階段獲取的數據(A1,G1,H,J,K,L)
        /// </summary>
        /// <param name="drowResult"></param>
        /// <returns></returns>
        private DataTable math(DataTable dtbStep1,int manageMonth)
        {   
            DataTable dataTable2 = new DataTable();

            try
            {
                creatNewDataTable2(dataTable2);//創建包含所有需要數據的表格
                                               //遍歷第一階段取得的數據表
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
                    //2008/12/10 J值改爲不保存,不計算
                    //計算J值
                    //drow2["J"] = getJValue(dtbStep1, j, manageMonth,drow2["G1"]);
                    drow2["J"] = 0;
                    //計算K值             
                    drow2["K"] = toInt(drow2["J"]) + toInt(drow2["G1"]);
                    //計算L檢核欄位的值
                    drow2["L"] = getLValue(drow2["G1"], drow2["J"], dtbStep1, j);
                    dataTable2.Rows.Add(drow2);
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("根據第一階段獲取的數據計算第二階段獲取的數據(A1,G1,H,J,K,L), math報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

            return dataTable2;
        }

        /// <summary>
        /// 創建第二階段的數據表
        /// </summary>
        /// <param name="dtblSafeStockInfo"></param>
        public void creatNewDataTable2(DataTable dtblSafeStockInfo)
        {

            //月份
            dtblSafeStockInfo.Columns.Add("Month");
            //可用庫存量
            dtblSafeStockInfo.Columns.Add("A");
            //修正庫存量
            dtblSafeStockInfo.Columns.Add("A1");
            //過去X月平均新卡數
            dtblSafeStockInfo.Columns.Add("B");
            //新進件調整
            dtblSafeStockInfo.Columns.Add("C");
            //過去X月平均掛補數
            dtblSafeStockInfo.Columns.Add("D1");
            //過去X月平均毀補數
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
            catch(Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取L欄位的值, getLValue報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
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
        private void setModelToDataRow(MONTHLY_MONITOR model, decimal TotalXH, DataRow drow)
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
            drow["Month"] = model.CMonth;
        }


        /// <summary>
        /// 當天之前的已下單未到貨數
        /// </summary>
        /// <param name="drow"></param>
        /// <returns></returns>
        private int getOderedNotStorckNum(int cardRid, int persoRid, DateTime beginDate, DateTime endDate)
        {
            int OderedNotStorckNum = 0;

            dirValues.Clear();
            dirValues.Add("PersoRid", persoRid);
            dirValues.Add("cardRid", cardRid);
            dirValues.Add("beginDate", beginDate);
            dirValues.Add("endDate", endDate);
            OderedNotStorckNum = int.Parse(dao.GetList(SEL_NOT_INCOME_NUMBER, dirValues).Tables[0].Rows[0][0].ToString());
            return OderedNotStorckNum;
        }

        /// <summary>
        /// 取開始時間到結束時間的Action= @action 的小計數量
        /// </summary>
        /// <param name="dir"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        private double getNumberFromSubTotal(CARD_TYPE cardModel, int persoRid,DateTime beginDate, DateTime endDate, string action)
        {
            double result = 0;
            dirValues.Clear();
            //dirValues.Add("PersoRid", persoRid);
            dirValues.Add("photo", cardModel.PHOTO);
            dirValues.Add("type", cardModel.TYPE);
            dirValues.Add("affinity", cardModel.AFFINITY);
            dirValues.Add("begin_date", beginDate);
            dirValues.Add("end_date", endDate);
            dirValues.Add("action", action);
            result =Convert.ToDouble(dao.GetList(SEL_PREMONTHS_NUMBER, dirValues).Tables[0].Rows[0][0].ToString());
            return result;
        }

        private void creatNewDataTable(DataTable dtblSafeStockInfo)
        {
            //月份
            dtblSafeStockInfo.Columns.Add("Month");
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
        }

        /// <summary>
        /// 
        /// 根據PERSO厰RID和卡种RID查詢日結庫存量
        /// </summary>
        /// <param name="factoryRid"></param>
        /// <param name="cardTypeRid"></param>
        public DataRow getCheckStockByPerso(int factoryRid, int cardTypeRid)
        {
            dirValues.Clear();
            dirValues.Add("persoRid", factoryRid.ToString());
            dirValues.Add("cardRid", cardTypeRid.ToString());
            try
            {
                DataRow drResult = dao.GetRow(SEL_STOCKS, dirValues, false);
                return drResult;
            }
            catch (Exception ex)
            {
                LogFactory.Write("getCheckStockByPerso報錯: " + ex.Message, GlobalString.LogType.ErrorCategory);
                return null;
            }
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
                DataRow drStock = getCheckStockByPerso(persoRid, cardTypeRid);
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
            }
            catch (Exception ex)
            {
                LogFactory.Write("getCurrentStock報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            return result;
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
            DataTable dtblResult = null;
            dirValues.Clear();
            dirValues.Add("persoRid", persoRid);
            dirValues.Add("cardRid", cardTypeRid);
            dirValues.Add("stockDate", CheckDay);

            Object oResult = dao.ExecuteScalar(SEL_STOCKS_NUM, dirValues);

            if (oResult != null)
            {
                return Convert.ToInt32(oResult);
            }
            return 0;
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
            //200909IR
            int subUseNum = manager.getUseCardNumSub(factoryRID, cardRID, storkDate, currentDay);
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
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("判斷是否為日結日, isCheckDate報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            return false;
        }

    }
 
}
