using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;
using CIMSBatch.Public;

namespace CIMSService.Business
{
    class DataImport_C001BL : BaseLogic
    {
        #region SQL
        public const string SEL_CARDTYPE_STOCKS1 = "select tcs.Stock_Date, tcs.Perso_Factory_RID, tcs.CardType_RID, tcs.Stocks_Number as Stocks_Number_T, cs.Stocks_Number "
                                                        + " from CARDTYPE_STOCKS cs "
                                                        + " inner join T_CARDTYPE_STOCKS tcs "
                                                        + " on tcs.Stock_Date = cs.Stock_Date "
                                                        + " and tcs.Perso_Factory_RID = cs.Perso_Factory_RID "
                                                        + " and tcs.CardType_RID = cs.CardType_RID "
                                                        + " where cs.Stock_Date = @Stock_Date ";

        public const string SEL_CARDTYPE_STOCKS2 = "select cs.Stock_Date, cs.Perso_Factory_RID, cs.CardType_RID, cs.Stocks_Number "
                                                + " from CARDTYPE_STOCKS cs "
                                                + " left join T_CARDTYPE_STOCKS tcs "
                                                + " on tcs.Stock_Date = cs.Stock_Date "
                                                + " and tcs.Perso_Factory_RID = cs.Perso_Factory_RID "
                                                + " and tcs.CardType_RID = cs.CardType_RID "
                                                + " where cs.Stock_Date = @Stock_Date and tcs.RID IS NULL ";

        public const string SEL_CARDTYPE_STOCKS3 = "select tcs.Stock_Date, tcs.Perso_Factory_RID, tcs.CardType_RID, tcs.Stocks_Number "
                                                + " from CARDTYPE_STOCKS cs "
                                                + " right join T_CARDTYPE_STOCKS tcs "
                                                + " on tcs.Stock_Date = cs.Stock_Date "
                                                + " and tcs.Perso_Factory_RID = cs.Perso_Factory_RID "
                                                + " and tcs.CardType_RID = cs.CardType_RID "
                                                + " where tcs.Stock_Date = @Stock_Date and cs.RID IS NULL ";
        #endregion

        Dictionary<string, object> dirValues = new Dictionary<string, object>();

        public void CheckStocks()
        {
            try
            {
                string strStockDate = ConfigurationManager.AppSettings["HistoryDateTo"];

                string strPersoFactoryRID = "";
                string strCardType_RID = "";
                int iStocksNumber_T = 0;
                int iStocksNumber = 0;

                dirValues.Clear();
                dirValues.Add("Stock_Date", strStockDate);

                //B.10�MB.11�����s�b���ƾ�
                DataTable dtStocks = dao.GetList(SEL_CARDTYPE_STOCKS1, dirValues).Tables[0];
                for (int i = 0; i < dtStocks.Rows.Count; i++ )
                {
                    strPersoFactoryRID = dtStocks.Rows[0]["Perso_Factory_RID"].ToString().Trim();
                    strCardType_RID = dtStocks.Rows[0]["CardType_RID"].ToString().Trim();
                    iStocksNumber_T = Convert.ToInt32(dtStocks.Rows[0]["Stocks_Number_T"].ToString().Trim());
                    iStocksNumber = Convert.ToInt32(dtStocks.Rows[0]["Stocks_Number"].ToString().Trim());

                    LogFactory.Write("Perso�tRID:" + strPersoFactoryRID + "; �d�ؽs��RID:" + strCardType_RID + "; �t���w�s��:" + (iStocksNumber - iStocksNumber_T), "");
                }

                //B.10���s�b�BB.11�����s�b���ƾ�
                dtStocks = dao.GetList(SEL_CARDTYPE_STOCKS2, dirValues).Tables[0];
                for (int i = 0; i < dtStocks.Rows.Count; i++)
                {
                    strPersoFactoryRID = dtStocks.Rows[0]["Perso_Factory_RID"].ToString().Trim();
                    strCardType_RID = dtStocks.Rows[0]["CardType_RID"].ToString().Trim();
                    iStocksNumber = Convert.ToInt32(dtStocks.Rows[0]["Stocks_Number"].ToString().Trim());

                    LogFactory.Write("Perso�tRID:" + strPersoFactoryRID + "; �d�ؽs��RID:" + strCardType_RID + "; �t���w�s��:" + iStocksNumber, "");
                }

                //B.10�����s�b�BB.11���s�b���ƾ�
                dtStocks = dao.GetList(SEL_CARDTYPE_STOCKS3, dirValues).Tables[0];
                for (int i = 0; i < dtStocks.Rows.Count; i++)
                {
                    strPersoFactoryRID = dtStocks.Rows[0]["Perso_Factory_RID"].ToString().Trim();
                    strCardType_RID = dtStocks.Rows[0]["CardType_RID"].ToString().Trim();
                    iStocksNumber = Convert.ToInt32(dtStocks.Rows[0]["Stocks_Number"].ToString().Trim());

                    LogFactory.Write("Perso�tRID:" + strPersoFactoryRID + "; �d�ؽs��RID:" + strCardType_RID + "; �t���w�s��:-" + iStocksNumber, "");
                }

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("CheckStocks����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

        }
    }
}
