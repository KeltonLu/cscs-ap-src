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
    class CardSplitToPersoBL : BaseLogic
    { 
        #region SQL
        public const string SEL_CARDTYPE_PERSO_SPECIAL_1 = "SELECT PC.* FROM PERSO_CARDTYPE PC WHERE PC.rst='A' and PC.CardType_RID = @cardtype_rid and PC.Percentage_Number = '1' ORDER BY PC.Priority desc";
        public const string SEL_CARDTYPE_PERSO_SPECIAL_2 = "SELECT PC.* FROM PERSO_CARDTYPE PC WHERE PC.rst='A' and PC.CardType_RID = @cardtype_rid and PC.Percentage_Number = '2' ORDER BY PC.Priority";
        public const string SEL_CARDTYPE_PERSO = "SELECT PC.* FROM PERSO_CARDTYPE PC WHERE PC.RST = 'A' AND PC.CardType_RID = @cardtype_rid and pc.Base_Special = '1'";
        public const string SEL_CARD_TYPE = "SELECT * FROM CARD_TYPE WHERE RST='A' AND TYPE = @TYPE  AND AFFINITY = @AFFINITY AND PHOTO = @PHOTO";
        public const string SEL_NEXT16_MONTH = "select * from  FORE_CHANGE_CARD WHERE Change_Date >@begainMonth and Change_Date <=@endMonth";        
#endregion
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        /// <summary>
        /// 將卡片年度預測檔進行換卡拆分
        /// </summary>
        public void SplitToPerso()
        {
            try
            {
                dao.OpenConnection();
                dirValues.Clear();
                dirValues.Add("begainMonth", DateTime.Now.ToString("yyyyMM"));
                dirValues.Add("endMonth", DateTime.Now.AddMonths(16).ToString("yyyyMM"));
                dao.ExecuteNonQuery("delete FORE_CHANGE_CARD_DETAIL where change_date >@begainMonth and Change_Date <=@endMonth", dirValues);
                DataTable dtbl = dao.GetList(SEL_NEXT16_MONTH, dirValues).Tables[0];
                foreach (DataRow drow in dtbl.Rows)
                {

                    FORE_CHANGE_CARD foreChangeCard = dao.GetModelByDataRow<FORE_CHANGE_CARD>(drow);    //原始換卡model

                    CARD_TYPE card = getCardtype(foreChangeCard.Affinity, foreChangeCard.Photo, foreChangeCard.Type);//卡種對象

                    if (card == null)
                    {
                        continue;
                    }

                    int cardRid = card.RID;
                    //if(如果該卡種有換卡版面),將卡種替換成換卡版面
                    if (card.Change_Space_RID != 0)
                    {
                        CARD_TYPE changeCard = dao.GetModel<CARD_TYPE, int>("RID", card.Change_Space_RID);
                        drow["Photo"] = changeCard.PHOTO;
                        drow["Affinity"] = changeCard.AFFINITY;
                        drow["Type"] = changeCard.TYPE;
                        cardRid = changeCard.RID;
                    }
                    //else if(如果該卡種有替換卡版面),將卡種替換成替換卡版面
                    else if (card.Replace_Space_RID != 0)
                    {
                        CARD_TYPE replaceCard = dao.GetModel<CARD_TYPE, int>("RID", card.Replace_Space_RID);
                        drow["Photo"] = replaceCard.PHOTO;
                        drow["Affinity"] = replaceCard.AFFINITY;
                        drow["Type"] = replaceCard.TYPE;
                        cardRid = replaceCard.RID;
                    }
                    //else (不做任何替換)     
                    #region 拆分卡種到perso廠
                    dirValues.Clear();
                    dirValues.Add("cardtype_rid", cardRid);
                    DataSet dsCARDTYPE_PERSO = dao.GetList(SEL_CARDTYPE_PERSO_SPECIAL_1, dirValues);
                    DataSet dsCARDTYPE_PERSO_2 = dao.GetList(SEL_CARDTYPE_PERSO_SPECIAL_2, dirValues);
                    long totalNumber = foreChangeCard.Number;
                    //if(有比率分配),將卡種類按照比率進行分配
                    if (dsCARDTYPE_PERSO.Tables[0] != null && dsCARDTYPE_PERSO.Tables[0].Rows.Count > 0)
                    {
                        long leftNumber = totalNumber;
                        foreach (DataRow persoRow in dsCARDTYPE_PERSO.Tables[0].Rows)
                        {
                            FORE_CHANGE_CARD_DETAIL foreDetail = new FORE_CHANGE_CARD_DETAIL();
                            decimal percent = Convert.ToDecimal(persoRow["value"]) / 100M;
                            if (persoRow["Priority"].ToString() == "1")
                                foreDetail.Number = leftNumber;
                            else
                                foreDetail.Number = Convert.ToInt64(Math.Floor(totalNumber * percent));
                            foreDetail.Perso_Factory_RID = Convert.ToInt32(persoRow["Factory_RID"]);
                            foreDetail.Photo = drow["Photo"].ToString();
                            foreDetail.Affinity = drow["Affinity"].ToString();
                            foreDetail.Type = drow["Type"].ToString();
                            foreDetail.Change_Date = drow["Change_Date"].ToString();
                            dao.Add<FORE_CHANGE_CARD_DETAIL>(foreDetail, "RID");
                            leftNumber -= foreDetail.Number;
                        }

                    }
                    else if (dsCARDTYPE_PERSO_2.Tables[0] != null && dsCARDTYPE_PERSO_2.Tables[0].Rows.Count > 0)
                    {
                        long leftNumber = totalNumber;
                        foreach (DataRow persoRow in dsCARDTYPE_PERSO_2.Tables[0].Rows)
                        {
                            if (leftNumber != 0)
                            {
                                FORE_CHANGE_CARD_DETAIL foreDetail = new FORE_CHANGE_CARD_DETAIL();
                                if (persoRow["value"].ToString() == "0")
                                    foreDetail.Number = leftNumber;
                                else
                                    foreDetail.Number = Math.Min(Convert.ToInt32(persoRow["value"]), leftNumber);
                                foreDetail.Perso_Factory_RID = Convert.ToInt32(persoRow["Factory_RID"]);
                                foreDetail.Photo = drow["Photo"].ToString();
                                foreDetail.Affinity = drow["Affinity"].ToString();
                                foreDetail.Type = drow["Type"].ToString();
                                foreDetail.Change_Date = drow["Change_Date"].ToString();
                                dao.Add<FORE_CHANGE_CARD_DETAIL>(foreDetail, "RID");
                                leftNumber -= foreDetail.Number;
                            }
                        }
                    }
                    else
                    {
                        //else 將卡種直接分配到基本廠
                        DataTable dtbBasePerso = dao.GetList(SEL_CARDTYPE_PERSO, dirValues).Tables[0];
                        if (dtbBasePerso != null && dtbBasePerso.Rows.Count > 0)
                        {
                            FORE_CHANGE_CARD_DETAIL foreDetail = new FORE_CHANGE_CARD_DETAIL();
                            foreDetail.Number = Convert.ToInt64(totalNumber);
                            foreDetail.Photo = drow["Photo"].ToString();
                            foreDetail.Affinity = drow["Affinity"].ToString();
                            foreDetail.Type = drow["Type"].ToString();
                            foreDetail.Change_Date = drow["Change_Date"].ToString();
                            foreDetail.Perso_Factory_RID = Convert.ToInt32(dtbBasePerso.Rows[0]["Factory_RID"]);
                            dao.Add<FORE_CHANGE_CARD_DETAIL>(foreDetail, "RID");
                        }
                    }
                    #endregion
                } //事務提交
                dao.Commit();
            }
            catch (Exception ex)
            {
                dao.Rollback();
                LogFactory.Write("將卡片年度預測檔進行換卡拆分SplitToPerso方法報錯:" + ex.Message, GlobalString.LogType.ErrorCategory); // 加 將卡片年度預測檔進行換卡拆分SplitToPerso方法報錯 add judy 2018/03/29
                // BatchBL Bbl = new BatchBL();
                // Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);

            }
            finally
            {
                dao.CloseConnection();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="affinity"></param>
        /// <param name="photo"></param>
        /// <param name="cardtype"></param>
        /// <returns></returns>
        public CARD_TYPE getCardtype(string affinity, string photo, string cardtype)
        {
            dirValues.Clear();
            dirValues.Add("AFFINITY", affinity);
            dirValues.Add("PHOTO", photo);
            dirValues.Add("TYPE", cardtype);
            DataTable cardTable = dao.GetList(SEL_CARD_TYPE, dirValues).Tables[0];

            if (cardTable.Rows.Count == 0)
            {
                return null;
            }
            else
            {
                CARD_TYPE cardType = dao.GetModelByDataRow<CARD_TYPE>(cardTable.Rows[0]);
                return cardType;
            }
        }

    }
}
