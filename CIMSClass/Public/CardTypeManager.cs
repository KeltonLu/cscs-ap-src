//******************************************************************
//*  作    者：FangBao
//*  功能說明：卡種管理
//*  創建日期：2008-08-08
//*  修改日期：2008-08-08 12:00
//*  修改記錄：
//*            □2008-07-31
//*              1.創建 鮑方
//*******************************************************************

using System;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Collections.Generic;
using System.Text;
using Microsoft.Practices.EnterpriseLibrary.Data;

using CIMSClass;
using CIMSClass.Model;

/// <summary>
/// CardTypeManager 的摘要描述
/// </summary>
/// 
namespace CIMSClass
{
    public class CardTypeManager : BaseLogic
    {
        public CardTypeManager()
        {
            //
            // TODO: 在此加入建構函式的程式碼
            //
        }

        Dictionary<string, object> dirValues = new Dictionary<string, object>();

        #region sql語句
        public const string SEL_CARDPURPOSE = "SELECT Param_Code,PARAM_NAME FROM PARAM WHERE RST='A' AND PARAMTYPE_CODE='" + GlobalString.ParameterType.Use + "'";
        public const string SEL_CARDGROUP_BY_PRID = "SELECT RID,GROUP_NAME FROM CARD_GROUP WHERE RST='A' ";
        public const string SEL_CARDTYPE_BY_RID = "SELECT * FROM CARD_TYPE WHERE RST='A' ";
        public const string SEL_CARDTYPEGROUP_BY_RID = "SELECT * FROM CARD_TYPE WHERE RST='A' AND Is_Using='Y'";
        public const string CON_CHECK_DATE = "SELECT count(*) from cardtype_stocks where rst = 'A' and Stock_Date = @CheckDate";
        //查詢前日結庫存量
        public const string SEL_STOCKS = "SELECT TOP 1 Stock_Date,Stocks_Number"
                                    + " FROM CARDTYPE_STOCKS"
                                    + " WHERE RST='A' AND Perso_Factory_RID = @persoRid AND CardType_RID = @cardRid"
                                    + " ORDER BY Stock_Date desc";
        //查詢前日結庫存量
        public const string SEL_STOCKS_NUM = "SELECT Stocks_Number"
                                   + " FROM CARDTYPE_STOCKS"
                                   + " WHERE RST='A' AND Perso_Factory_RID = @persoRid AND CardType_RID = @cardRid and Stock_Date = @stockDate";
        //查詢最近日結後入庫量
        public const string SEL_DEPOSITORY_STOCK = "SELECT ISNULL(Sum(Income_Number),0) as Number"
                                    + " FROM DEPOSITORY_STOCK"
                                    + " WHERE RST = 'A' AND Income_Date > @Stock_Date and Income_Date <= @today AND Perso_Factory_RID = @persoRid AND Space_Short_RID = @cardRid";
        //查詢最近日結後退貨量
        public const string SEL_DEPOSITORY_CANCEl_NUMBER = "SELECT ISNULL(Sum(Cancel_Number),0) as Number"
                                    + " FROM DEPOSITORY_CANCEL"
                                    + " WHERE RST = 'A' AND Cancel_Date > @Stock_Date and Cancel_Date <= @today  AND Perso_Factory_RID = @persoRid AND Space_Short_RID = @cardRid";
        //查詢最近日結後再入庫量
        public const string SEL_DEPOSITORY_RESTOCK = "SELECT ISNULL(Sum(Reincome_Number),0) as Number"
                                    + " FROM DEPOSITORY_RESTOCK"
                                    + " WHERE RST = 'A' AND Reincome_Date > @Stock_Date and Reincome_Date <= @today AND Perso_Factory_RID = @persoRid AND Space_Short_RID = @cardRid";
        //查詢最近日結後卡片轉移量

        public const string CARDTYPE_STOCKS_MOVE_IN = "SELECT ISNULL(sum(Move_Number),0) as Number"
                                    + " FROM CARDTYPE_STOCKS_MOVE"
                                    + " WHERE RST = 'A' AND Move_Date > @Stock_Date and Move_Date <= @today AND To_Factory_RID = @persoRid AND CardType_RID = @cardRid";

        public const string CARDTYPE_STOCKS_MOVE_OUT = "SELECT ISNULL(sum(Move_Number),0) as Number"
                                    + " FROM CARDTYPE_STOCKS_MOVE"
                                    + " WHERE RST = 'A' AND Move_Date > @Stock_Date and Move_Date <= @today AND From_Factory_RID = @persoRid AND CardType_RID = @cardRid";

        public const string SEL_FACTORY_CHANGE_NUM = "select *"
                                                 + "from FACTORY_CHANGE_IMPORT "
                                                 + "where RST = 'A' AND TYPE = @TYPE "
                                                 + "AND PHOTO = @PHOTO AND AFFINITY = @AFFINITY AND Perso_Factory_RID = @Perso_Factory_RID "
                                                 + "AND Date_Time > @Stock_Date and Date_Time <= @today";
        public const string SEL_FACTORY_CHANGE_REPLACE_NUM = "select *"
                                                + "from FACTORY_CHANGE_REPLACE_IMPORT "
                                                + "where RST = 'A' AND TYPE = @TYPE "
                                                + "AND PHOTO = @PHOTO AND AFFINITY = @AFFINITY AND Perso_Factory_RID = @Perso_Factory_RID "
                                                + "AND Date_Time > @Stock_Date and Date_Time <= @today";

        public const string SEL_EXPRESSION_INFO = "select * from EXPRESSIONS_DEFINE "
                                                     + "where RST = 'A' AND Expressions_RID = @Expressions_RID";

        //查詢日結日到指定日期間的小計檔記錄的指定廠商指定卡種的消耗數量。
        private const string SEL_USED_CARD_NUMBER =
            "select isnull(sum(number),0) from subtotal_import si , card_type ct "
                    + "where si.photo=ct.photo"
                    + " and si.type = ct.type"
                    + " and si.affinity = ct.affinity"
                    + " and si.perso_factory_rid = @Perso_Factory_Rid"
                    + " and ct.rid = @Card_Type_Rid"
                    + " and date_time > @Begin_Date"
                    + " and date_time <= @End_Date";

        //根據廠商回復檔案，計算卡片異動信息。
        private const string SEL_CHANGE_CARD_NUMBER =
            "select SUM(CASE status_rid  WHEN 14 THEN number ELSE 0 END) "
            + "-  SUM(CASE status_rid  WHEN 15 THEN number ELSE 0 END)"
            + "+  SUM(CASE status_rid  WHEN 16 THEN number ELSE 0 END)"
            + "+  SUM(CASE status_rid  WHEN 17 THEN number ELSE 0 END)"
            + "from FACTORY_CHANGE_IMPORT fci , card_type ct "
            + "where fci.type = ct.type "
            + " and fci.affinity = ct.affinity"
            + " and fci.photo = ct.photo"
            + " and ct.rid = @Card_Type_Rid"
            + " and perso_factory_rid = @Perso_Factory_Rid"
            + " and date_time > @Begin_Date"
            + " and date_time < @End_Date";
        //根據廠商回復檔案，計算卡片異動信息。
        private const string SEL_CHANGE_REPLACE_CARD_NUMBER =
            "select SUM(CASE status_rid  WHEN 14 THEN number ELSE 0 END) "
            + "-  SUM(CASE status_rid  WHEN 15 THEN number ELSE 0 END)"
            + "+  SUM(CASE status_rid  WHEN 16 THEN number ELSE 0 END)"
            + "+  SUM(CASE status_rid  WHEN 17 THEN number ELSE 0 END)"
            + "from FACTORY_CHANGE_REPLACE_IMPORT fci , card_type ct "
            + "where fci.type = ct.type "
            + " and fci.affinity = ct.affinity"
            + " and fci.photo = ct.photo"
            + " and ct.rid = @Card_Type_Rid"
            + " and perso_factory_rid = @Perso_Factory_Rid"
            + " and date_time > @Begin_Date"
            + " and date_time < @End_Date";

        //判斷是否虛擬卡群組的卡！
        private const string SEL_IS_NETWORK =
            "select count(*) from card_type ct , card_group cg , GROUP_CARD_TYPE gat "
                + " where ct.rid= gat.cardtype_rid "
                + " and cg.rid = gat.group_rid "
                + " and ct.name = @CardName "
                + " and cg.group_name = '虛擬卡' ";


        #endregion


        /// <summary>
        /// 根據RID生成模型
        /// </summary>
        /// <param name="strRID"></param>
        /// <returns></returns>
        public CARD_TYPE GetModel(string strRID)
        {
            return dao.GetModel<CARD_TYPE, int>("RID", int.Parse(strRID));
        }

        /// <summary>
        /// 獲取用途
        /// </summary>
        /// <returns></returns>
        public DataSet GetPurpose()
        {
            DataSet dstPurpose = null;

            try
            {
                dstPurpose = dao.GetList(SEL_CARDPURPOSE);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            return dstPurpose;
        }

        /// <summary>
        /// 獲取群組
        /// </summary>
        /// <param name="strPurposeId">用途ID</param>
        /// <returns></returns>
        public DataSet GetGroupByPurposeId(string strPurposeId)
        {
            DataSet dstGroup = null;
            try
            {
                dirValues.Clear();

                string strSql = SEL_CARDGROUP_BY_PRID;

                if (!StringUtil.IsEmpty(strPurposeId))
                {
                    strSql += " and Param_Code=@Param_Code";
                    dirValues.Add("Param_Code", strPurposeId);
                }

                dstGroup = dao.GetList(strSql, dirValues);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            return dstGroup;
        }

        /// <summary>
        /// 獲取卡種
        /// </summary>
        /// <param name="strGroupId">群組ID</param>
        /// <returns></returns>
        public DataSet GetCardTypeByGroupId(string strPurposeId, string strGroupId, string strSqlCard)
        {
            DataSet dstCardType = null;
            try
            {
                dirValues.Clear();

                string strSql = SEL_CARDTYPE_BY_RID;

                if (!StringUtil.IsEmpty(strSqlCard))
                {
                    strSql = strSqlCard;
                }

                if (!StringUtil.IsEmpty(strPurposeId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid in (SELECT RID FROM CARD_GROUP WHERE Param_Code=@Param_Code AND CARD_GROUP.RST='A'))";
                    dirValues.Add("Param_Code", strPurposeId);
                }

                if (!StringUtil.IsEmpty(strGroupId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid =@group_rid)";
                    dirValues.Add("group_rid", int.Parse(strGroupId));
                }


                dstCardType = dao.GetList(strSql + " order by display_name", dirValues);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            return dstCardType;
        }

        /// <summary>
        /// 獲取卡種
        /// </summary>
        /// <param name="strPurposeId">用途ID</param>
        /// <param name="strGroupId">群組ID</param>
        /// <param name="strSqlCard">卡種SQL</param>
        /// <param name="Is_Using">卡種是否可用</param>
        /// <returns></returns>
        public DataSet GetCardTypeByGroupId(string strPurposeId, string strGroupId, string strSqlCard, bool Is_Using)
        {
            DataSet dstCardType = null;
            try
            {
                dirValues.Clear();

                string strSql = SEL_CARDTYPE_BY_RID;

                if (!StringUtil.IsEmpty(strSqlCard))
                {
                    strSql = strSqlCard;
                }

                if (!StringUtil.IsEmpty(strPurposeId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid in (SELECT RID FROM CARD_GROUP WHERE Param_Code=@Param_Code AND CARD_GROUP.RST='A'))";
                    dirValues.Add("Param_Code", strPurposeId);
                }

                if (!StringUtil.IsEmpty(strGroupId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid =@group_rid)";
                    dirValues.Add("group_rid", int.Parse(strGroupId));
                }

                if (Is_Using)
                    strSql += " and Is_Using='Y'";


                dstCardType = dao.GetList(strSql + " order by display_name", dirValues);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            return dstCardType;
        }

        /// <summary>
        /// 獲取卡種
        /// </summary>
        /// <param name="strGroupId">群組ID</param>
        /// <returns></returns>
        public DataSet GetCardTypeByGroupId(string strPurposeId, string strGroupId)
        {
            DataSet dstCardType = null;
            try
            {
                dirValues.Clear();

                string strSql = SEL_CARDTYPEGROUP_BY_RID;

                if (!StringUtil.IsEmpty(strPurposeId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid in (SELECT RID FROM CARD_GROUP WHERE Param_Code=@Param_Code AND CARD_GROUP.RST='A'))";
                    dirValues.Add("Param_Code", strPurposeId);
                }

                if (!StringUtil.IsEmpty(strGroupId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid =@group_rid)";
                    dirValues.Add("group_rid", int.Parse(strGroupId));
                }


                dstCardType = dao.GetList(strSql + " order by display_name", dirValues);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            return dstCardType;
        }

        /// <summary>
        /// 獲取卡種
        /// </summary>
        /// <param name="strPurposeId"></param>
        /// <param name="strGroupId"></param>
        /// <param name="Is_Using">是否可用</param>
        /// <returns></returns>
        public DataSet GetCardTypeByGroupId(string strPurposeId, string strGroupId, bool Is_Using)
        {
            DataSet dstCardType = null;
            try
            {
                dirValues.Clear();

                string strSql = SEL_CARDTYPEGROUP_BY_RID;

                if (!StringUtil.IsEmpty(strPurposeId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid in (SELECT RID FROM CARD_GROUP WHERE Param_Code=@Param_Code AND CARD_GROUP.RST='A'))";
                    dirValues.Add("Param_Code", strPurposeId);
                }

                if (!StringUtil.IsEmpty(strGroupId))
                {
                    strSql += " and rid in (select distinct cardType_rid from dbo.GROUP_CARD_TYPE where rst='a' and group_rid =@group_rid)";
                    dirValues.Add("group_rid", int.Parse(strGroupId));
                }

                if (Is_Using)
                    strSql += " and Is_Using='Y'";

                dstCardType = dao.GetList(strSql + " order by display_name", dirValues);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            return dstCardType;
        }

        /// <summary>
        /// 根據輸入的廠商，卡種及日期，根據廠商異動檔案計算當前庫存！
        /// </summary>
        /// <param name="PersoRid"></param>
        /// <param name="cardTypeRid"></param>
        /// <param name="Checkday"></param>
        /// <returns></returns>
        public int getCurrentStockPerso(int PersoRid, int CardTypeRid, DateTime CurrentDate)
        {
            if (isCheckDate(CurrentDate))
            {
                return getStockNumByCheckDay(PersoRid, CardTypeRid, CurrentDate);
            }
            DataSet dstlstock = getCheckStockByPerso(PersoRid, CardTypeRid);
            DateTime checkDate = Convert.ToDateTime("1900-01-01");
            int result = 0;
            if (dstlstock.Tables[0].Rows.Count != 0)
            {
                checkDate = Convert.ToDateTime(dstlstock.Tables[0].Rows[0]["Stock_Date"].ToString());
                result = Convert.ToInt32(dstlstock.Tables[0].Rows[0]["Stocks_Number"].ToString());
            }

            dirValues.Clear();
            dirValues.Add("Perso_Factory_Rid", PersoRid.ToString());
            dirValues.Add("Card_Type_Rid", CardTypeRid.ToString());
            dirValues.Add("Begin_Date", checkDate);
            dirValues.Add("End_Date", CurrentDate);

            //減去小計檔的消耗！
            //object oResult = dao.ExecuteScalar(SEL_USED_CARD_NUMBER, dirValues);
            //if (oResult != null && !Convert .IsDBNull ( oResult)  )
            //{
            //    result -= Convert.ToInt32(oResult);
            //}

            //加上入出庫存的數量！
            object oResult = dao.ExecuteScalar(SEL_CHANGE_CARD_NUMBER, dirValues);
            if (oResult != null && !Convert.IsDBNull(oResult))
            {
                result += Convert.ToInt32(oResult);
            }

            //減去消耗的卡片數量
            result -= getUseCardNum(PersoRid.ToString(), CardTypeRid.ToString(), checkDate, CurrentDate); //按廠商異動檔來計算！

            return result;
        }
        /// <summary>
        /// 根據輸入的廠商，卡種及日期，根據(替換前版面)廠商異動檔案計算當前庫存！
        /// </summary>
        /// <param name="PersoRid"></param>
        /// <param name="cardTypeRid"></param>
        /// <param name="Checkday"></param>
        /// <returns></returns>
        public int getCurrentStockPersoReplace(int PersoRid, int CardTypeRid, DateTime CurrentDate)
        {
            if (isCheckDate(CurrentDate))
            {
                return getStockNumByCheckDay(PersoRid, CardTypeRid, CurrentDate);
            }
            DataSet dstlstock = getCheckStockByPerso(PersoRid, CardTypeRid);
            DateTime checkDate = Convert.ToDateTime("1900-01-01");
            int result = 0;
            if (dstlstock.Tables[0].Rows.Count != 0)
            {
                checkDate = Convert.ToDateTime(dstlstock.Tables[0].Rows[0]["Stock_Date"].ToString());
                result = Convert.ToInt32(dstlstock.Tables[0].Rows[0]["Stocks_Number"].ToString());
            }

            dirValues.Clear();
            dirValues.Add("Perso_Factory_Rid", PersoRid.ToString());
            dirValues.Add("Card_Type_Rid", CardTypeRid.ToString());
            dirValues.Add("Begin_Date", checkDate);
            dirValues.Add("End_Date", CurrentDate);

          
            //加上入出庫存的數量！
            object oResult = dao.ExecuteScalar(SEL_CHANGE_REPLACE_CARD_NUMBER, dirValues);
            if (oResult != null && !Convert.IsDBNull(oResult))
            {
                result += Convert.ToInt32(oResult);
            }

            //減去消耗的卡片數量
            result -= getUseCardNumReplace(PersoRid.ToString(), CardTypeRid.ToString(), checkDate, CurrentDate); //按(替換前版面)廠商異動檔來計算！

            return result;
        }

        /// <summary>
        /// 檢查一個卡種是否為網絡卡
        /// </summary>
        /// <param name="cardname">卡片簡稱</param>
        /// <returns></returns>
        public bool isNetworkCard(string cardname)
        {
            bool result = true;
            dirValues.Clear();
            dirValues.Add("CardName", cardname);

            object oResult = dao.ExecuteScalar(SEL_IS_NETWORK, dirValues);
            if (oResult != null && !Convert.IsDBNull(oResult))
            {
                int iResult = Convert.ToInt32(oResult);
                if (iResult == 0)
                {
                    result = false;
                }
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
            dtblResult = dao.GetList(SEL_STOCKS_NUM, dirValues).Tables[0];
            if (dtblResult != null && dtblResult.Rows.Count > 0)
            {
                return int.Parse(dtblResult.Rows[0][0].ToString());
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
            if (isCheckDate(currentDay))
            {
                return getStockNumByCheckDay(persoRid, cardTypeRid, currentDay);
            }
            DataSet dstlstock = getCheckStockByPerso(persoRid, cardTypeRid);
            DateTime checkDate = Convert.ToDateTime("1900-01-01");
            int result = 0;
            if (dstlstock.Tables[0].Rows.Count != 0)
            {
                checkDate = Convert.ToDateTime(dstlstock.Tables[0].Rows[0]["Stock_Date"].ToString());
                result = Convert.ToInt32(dstlstock.Tables[0].Rows[0]["Stocks_Number"].ToString());
            }

            dirValues.Clear();
            dirValues.Add("persoRid", persoRid.ToString());
            dirValues.Add("cardRid", cardTypeRid.ToString());
            dirValues.Add("Stock_Date", checkDate);
            dirValues.Add("today", currentDay);
            //加上所有入庫記錄
            result += Convert.ToInt32(dao.GetList(SEL_DEPOSITORY_STOCK, dirValues).Tables[0].Rows[0][0].ToString());
            //減去所有退貨記錄
            result -= Convert.ToInt32(dao.GetList(SEL_DEPOSITORY_CANCEl_NUMBER, dirValues).Tables[0].Rows[0][0].ToString());
            //加上所有再入庫記錄
            result += Convert.ToInt32(dao.GetList(SEL_DEPOSITORY_RESTOCK, dirValues).Tables[0].Rows[0][0].ToString());
            //加上所有移轉進入的紀錄
            result += Convert.ToInt32(dao.GetList(CARDTYPE_STOCKS_MOVE_IN, dirValues).Tables[0].Rows[0][0].ToString());
            //減去所有移出記錄
            result -= Convert.ToInt32(dao.GetList(CARDTYPE_STOCKS_MOVE_OUT, dirValues).Tables[0].Rows[0][0].ToString());
            //減去消耗的卡片數量
            //result -= getUseCardNum(persoRid.ToString(), cardTypeRid.ToString(), checkDate, currentDay); //按廠商異動檔來計算！

            result -= this.getUseCardNumSub(persoRid.ToString(), cardTypeRid.ToString(), checkDate, currentDay); //按小計檔中的記錄來計算！

            return result;
        }

        /// <summary>
        /// 根據小訂檔中的記錄來計算卡片消耗！
        /// </summary>
        /// <param name="PersoFactoryRid"></param>
        /// <param name="CardTypeRid"></param>
        /// <param name="BeginDate"></param>
        /// <param name="EndDate"></param>
        /// <returns></returns>
        public int getUseCardNumSub(string PersoFactoryRid, string CardTypeRid, DateTime BeginDate, DateTime EndDate)
        {
            int result = 0;
            dirValues.Clear();
            dirValues.Add("Perso_Factory_Rid", PersoFactoryRid);
            dirValues.Add("Card_Type_Rid", CardTypeRid);
            dirValues.Add("Begin_Date", BeginDate);
            dirValues.Add("End_Date", EndDate);

            object oResult = dao.ExecuteScalar(SEL_USED_CARD_NUMBER, dirValues);

            if (oResult != null && !Convert.IsDBNull(oResult))
            {
                result = Convert.ToInt32(oResult);
            }
            return result;
        }


        /// <summary>
        /// 
        /// 根據PERSO厰RID和卡种RID查詢日結庫存量
        /// </summary>
        /// <param name="factoryRid"></param>
        /// <param name="cardTypeRid"></param>
        public DataSet getCheckStockByPerso(int factoryRid, int cardTypeRid)
        {
            DataSet dstlstock = null;
            dirValues.Clear();
            dirValues.Add("persoRid", factoryRid.ToString());
            dirValues.Add("cardRid", cardTypeRid.ToString());
            try
            {
                dstlstock = dao.GetList(SEL_STOCKS, dirValues);
            }
            catch (Exception ex)
            {
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_SearchFail);
            }
            return dstlstock;
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
            //按照公式計算消耗卡數量
            foreach (DataRow drowExp in expTable.Rows)
            {
                EXPRESSIONS_DEFINE expModel = new EXPRESSIONS_DEFINE();
                expModel = dao.GetModelByDataRow<EXPRESSIONS_DEFINE>(drowExp);
                //查詢該卡种是否應該在公式中出現，如果不出現，則ＣＯＮＴＩＮＵＥ
                CARDTYPE_STATUS statusModel = dao.GetModel<CARDTYPE_STATUS, int>("RID", expModel.Type_RID);
                if (statusModel.Is_Display.Equals(GlobalString.YNType.No))
                {
                    continue;
                }
                int SUM = 0;
                foreach (DataRow drowFac in factoryTable.Rows)
                {
                    FACTORY_CHANGE_IMPORT importModel = new FACTORY_CHANGE_IMPORT();
                    importModel = dao.GetModelByDataRow<FACTORY_CHANGE_IMPORT>(drowFac);
                    //計算相同計算卡种狀態的異動數量
                    if (importModel.Status_RID == expModel.Type_RID)
                    {
                        SUM += importModel.Number;
                    }
                }
                //當前卡种狀態為+
                if (expModel.Operate == GlobalString.Operation.Add_RID)
                {
                    result += SUM;
                }
                //當前卡种狀態為-
                if (expModel.Operate == GlobalString.Operation.Del_RID)
                {
                    result -= SUM;
                }
            }
            return result;
        }
        public int getUseCardNumReplace(String factoryRID, string cardRID, DateTime storkDate, DateTime currentDay)
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
            //查詢(替換前版面)廠商異動信息
            DataSet factorySet = dao.GetList(SEL_FACTORY_CHANGE_REPLACE_NUM, dirValues);
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
            //按照公式計算消耗卡數量
            foreach (DataRow drowExp in expTable.Rows)
            {
                EXPRESSIONS_DEFINE expModel = new EXPRESSIONS_DEFINE();
                expModel = dao.GetModelByDataRow<EXPRESSIONS_DEFINE>(drowExp);
                //查詢該卡种是否應該在公式中出現，如果不出現，則ＣＯＮＴＩＮＵＥ
                CARDTYPE_STATUS statusModel = dao.GetModel<CARDTYPE_STATUS, int>("RID", expModel.Type_RID);
                if (statusModel.Is_Display.Equals(GlobalString.YNType.No))
                {
                    continue;
                }
                int SUM = 0;
                foreach (DataRow drowFac in factoryTable.Rows)
                {
                    FACTORY_CHANGE_REPLACE_IMPORT importModel = new FACTORY_CHANGE_REPLACE_IMPORT();
                    importModel = dao.GetModelByDataRow<FACTORY_CHANGE_REPLACE_IMPORT>(drowFac);
                    //計算相同計算卡种狀態的異動數量
                    if (importModel.Status_RID == expModel.Type_RID)
                    {
                        SUM += importModel.Number;
                    }
                }
                //當前卡种狀態為+
                if (expModel.Operate == GlobalString.Operation.Add_RID)
                {
                    result += SUM;
                }
                //當前卡种狀態為-
                if (expModel.Operate == GlobalString.Operation.Del_RID)
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
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_SearchFail);
            }
        }
    }
}
