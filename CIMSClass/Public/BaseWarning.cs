using System;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Collections.Generic;
using CIMSClass.Model;
using CIMSClass.FTP;
using CIMSClass.Mail;
using System.Text;
using System.Web;
using Microsoft.Practices.EnterpriseLibrary.Data;


/// <summary>
/// BaseWarning 的摘要描述
/// </summary>
namespace CIMSClass
{
    public class BaseWarning : BaseLogic
    {
        public const string SEL_USERS_BY_RID = "select WU.USERID,UR.USERNAME,UR.EMAIL from WARNING_USER WU inner join UseRS UR on WU.USERID=UR.USERID WHERE WU.WARNING_RID=@WARNING_RID";

        //?
        Dictionary<string, object> dirValues = new Dictionary<string, object>();

        public BaseWarning()
        {
            //
            // TODO: 在此加入建構函式的程式碼
            //
        }


        private string GetUser()
        {
            return "admin";
        }


        /// <summary>
        /// 獲取需要警訊的用戶并添加
        /// </summary>
        /// <param name="strWarningRID"></param>
        /// <returns></returns>
        private void GetWarningUser(WARNING_CONFIGURATION wcModel, string strWarningContent)
        {
            try
            {

                dirValues.Clear();
                dirValues.Add("WARNING_RID", wcModel.RID);
                DataTable dtblUser = dao.GetList(SEL_USERS_BY_RID, dirValues).Tables[0];

                string strMailTitle = ConfigurationManager.AppSettings["MailTitle"].ToString() + wcModel.Item_Name;

                if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                    return;

                foreach (DataRow drowUser in dtblUser.Rows)
                {
                    /*2009-06-24 modify by huangping 
                       * 【每日批次作業時】RID=15
                       * 【物料低於安全庫存(自動或人工匯入)】RID=14
                       * 【每月批次作業時】RID=29
                       * add by Ian Huang start
                       * 【物料控管作業】RID=13
                       * add by Ian Huang end
                       * 每日只發送一次警訊
                       * 
                    */
                    if (wcModel.RID == 15 || wcModel.RID == 14 || wcModel.RID == 29 || wcModel.RID == 13)
                    {
                        strWarningContent = strWarningContent.Replace("\\n", "\n");

                        DataSet ds = dao.GetList("select wi.RID from WARNING_INFO wi,WARNING_CONFIGURATION wc where wc.RID="
                        + wcModel.RID + " and wi.Warning_Content='" + strWarningContent + "' and wi.UserID='" + drowUser["userid"].ToString()
                        + "' and convert(char,wi.Warning_Date,111)=convert(char,getdate(),111)");

                        if (ds.Tables[0].Rows.Count > 0)
                        {
                            return;
                        }
                        else
                        {
                            WARNING_INFO wiModel = new WARNING_INFO();
                            wiModel.Warning_RID = wcModel.RID;
                            wiModel.Warning_Content = strWarningContent.Replace("\\n", "\n");
                            wiModel.UserID = drowUser["userid"].ToString();
                            wiModel.Is_Show = wcModel.System_Show;
                            dao.Add<WARNING_INFO>(wiModel, "RID");
                        }
                    }
                    else
                    {
                        if (wcModel.System_Show == "Y")
                        {
                            WARNING_INFO wiModel = new WARNING_INFO();
                            wiModel.Warning_RID = wcModel.RID;
                            wiModel.Warning_Content = strWarningContent.Replace("\\n", "\n");
                            wiModel.UserID = drowUser["userid"].ToString();
                            wiModel.Is_Show = wcModel.System_Show;
                            dao.Add<WARNING_INFO>(wiModel, "RID");
                        }
                    }


                    if (wcModel.Mail_Show == "Y")
                    {
                        USERS uModel = dao.GetModel<USERS, string>("UserID", drowUser["userid"].ToString());
                        if (uModel != null)
                        {
                            MailBase.SendMail(uModel.Email, strMailTitle, strWarningContent);
                        }
                    }
                }
            }
            catch
            {
            }

        }

        /// <summary>
        /// 代製費用異動 自動匯入有誤時
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void PersoProjectChange(string strRID, string strFactoryName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strFactoryName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }


        private PARAM GetParamModel(string strParamCode)
        {
            PARAM pModel = new PARAM();
            try
            {
                dirValues.Clear();
                dirValues.Add("param_code", strParamCode);
                pModel = dao.GetModel<PARAM>("select param_name from param where param_code=@param_code and paramType_code='" + GlobalString.ParameterType.CardParam + "'", dirValues);

            }
            catch { }

            return pModel;
        }

        private WARNING_CONFIGURATION GetModel(string strRID)
        {

            WARNING_CONFIGURATION wcModel = new WARNING_CONFIGURATION();

            try
            {
                wcModel = dao.GetModel<WARNING_CONFIGURATION, int>("RID", int.Parse(strRID));

            }
            catch { }

            return wcModel;
        }



        /// <summary>
        /// 修改預算警訊
        /// </summary>
        /// <param name="strRID"></param>
        public void EditBugdet(string strRID, string strBudgetID, string strType)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[4];
            arg[0] = GetUser();
            arg[1] = DateTime.Now.ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
            arg[2] = strType;
            arg[3] = strBudgetID;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);

        }

        /// <summary>
        /// 預算卡數過小
        /// </summary>
        /// <param name="strRID"></param>
        public void BudgetCardLower(string strRID, string strBudgetID, int Total_Card_Num, int Remain_Total_Num)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            if (Total_Card_Num == 0)
                return;

            PARAM pModel = GetParamModel("2");

            if (pModel == null)
                return;

            decimal decResult = Remain_Total_Num * 100 / Total_Card_Num;
            decimal decStard = Convert.ToDecimal(pModel.Param_Name.Substring(0, pModel.Param_Name.Length - 1));

            if (decResult < decStard)
            {
                object[] arg = new object[2];
                arg[0] = strBudgetID;
                arg[1] = decStard;
                string strWarningContent = string.Format(wcModel.Warning_Content, arg);
                GetWarningUser(wcModel, strWarningContent);
            }
        }

        /// <summary>
        /// 預算金額過小
        /// </summary>
        /// <param name="strRID"></param>
        public void BudgetAmtLower(string strRID, string strBudgetID, decimal Total_Card_AMT, decimal Remain_Total_AMT)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            PARAM pModel = GetParamModel("1");

            if (pModel == null)
                return;

            decimal decResult = Remain_Total_AMT * 100 / Total_Card_AMT;
            decimal decStard = Convert.ToDecimal(pModel.Param_Name.Substring(0, pModel.Param_Name.Length - 1));

            if (decResult < decStard)
            {
                object[] arg = new object[2];
                arg[0] = strBudgetID;
                arg[1] = decStard;
                string strWarningContent = string.Format(wcModel.Warning_Content, arg);
                GetWarningUser(wcModel, strWarningContent);
            }
        }

        /// <summary>
        /// 預算日期過小
        /// </summary>
        /// <param name="strRID"></param>
        public void BudgetDateLower(string strRID, string strBudgetID)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            PARAM pModel = GetParamModel("3");

            if (pModel == null)
                return;

            DataTable dtblBudget = dao.GetList("select max(valid_date_to) from CARD_BUDGET where Budget_Main_RID in (select rid from CARD_BUDGET where budget_id='" + strBudgetID + "')").Tables[0];

            DateTime dtMaxBudget = Convert.ToDateTime(dtblBudget.Rows[0][0]);

            TimeSpan ts = dtMaxBudget - Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));

            if (ts.Days < Convert.ToInt32(pModel.Param_Code))
            {
                object[] arg = new object[2];
                arg[0] = strBudgetID;
                arg[1] = pModel.Param_Code;
                string strWarningContent = string.Format(wcModel.Warning_Content, arg);
                GetWarningUser(wcModel, strWarningContent);
            }
        }


        /// <summary>
        /// 修改合約警訊
        /// </summary>
        /// <param name="strRID"></param>
        public void EditAgreement(string strRID, string strAgreementID, string strType)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[4];
            arg[0] = GetUser();
            arg[1] = DateTime.Now.ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
            arg[2] = strType;
            arg[3] = strAgreementID;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);

        }

        /// <summary>
        ///合約卡數過小
        /// </summary>
        /// <param name="strRID"></param>
        public void AgreementCardLower(string strRID, string strAgreementID, int Total_Card_Num, int Remain_Total_Num)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            if (Total_Card_Num == 0)
                return;

            PARAM pModel = GetParamModel("4");

            if (pModel == null)
                return;

            decimal decResult = Remain_Total_Num * 100 / Total_Card_Num;

            decimal decStard = Convert.ToDecimal(pModel.Param_Name.Substring(0, pModel.Param_Name.Length - 1));

            if (decResult < decStard)
            {
                object[] arg = new object[2];
                arg[0] = strAgreementID;
                arg[1] = decStard;

                string strWarningContent = string.Format(wcModel.Warning_Content, arg);

                GetWarningUser(wcModel, strWarningContent);
            }
        }

        /// <summary>
        /// 合約日期過小
        /// </summary>
        /// <param name="strRID"></param>
        public void AgreementDateLower(string strRID, string strAgreementID)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            PARAM pModel = GetParamModel("5");

            if (pModel == null)
                return;

            //if (intDay < Convert.ToInt32(pModel.Param_Code))
            //{
            //    object[] arg = new object[2];
            //    arg[0] = strAgreementID;
            //    arg[1] = pModel.Param_Code;
            //    string strWarningContent = string.Format(wcModel.Warning_Content, arg);
            //    GetWarningUser(wcModel, strWarningContent);
            //}
        }


        /// <summary>
        /// 卡種修改
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strName"></param>
        public void CardTypeEdit(string strRID, string strName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 卡種新增
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strAgreementID"></param>
        /// <param name="strType"></param>
        public void CardTypeAdd(string strRID, string strDate, string strName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[2];
            arg[0] = strDate;
            arg[1] = strName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }


        /// <summary>
        /// 采購下單待放行時
        /// </summary>
        /// <param name="strRID"></param>
        public void OrderFormCommit(string strRID)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;


            string strWarningContent = wcModel.Warning_Content;

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 新增物料採購
        /// </summary>
        /// <param name="strRID"></param>    
        public void AddPurchase(string strRID)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            string strWarningContent = wcModel.Warning_Content;

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 新增或修改物料採購作業時
        /// </summary>
        /// <param name="strRID"></param>
        public void PlsAskFinance(string strRID, string strMonthDay, string strOperator)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[2];
            arg[0] = strMonthDay;
            arg[1] = strOperator;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 獲得指定perso廠指定物料的耗損率
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        /// <param name="MaterialName"></param>
        /// <param name="wafer"></param>
        public void MaterialDataIn(string strRID, string strFactoryName, string MaterialName, decimal wafer)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[3];
            arg[0] = strFactoryName;
            arg[1] = MaterialName;
            arg[2] = wafer;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);

        }

        /// <summary>
        /// perso廠物料庫存自動匯入格式有誤
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void MaterialAutoDataIn(string strRID, string strFactoryName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strFactoryName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 物料系統結餘數不足
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void MaterialDataInMiss(string strRID, string MaterialName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = MaterialName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 物料系統結餘數低於安全庫存
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void MaterialDataInSafe(string strRID, string MaterialName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = MaterialName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 代製費用請款新增、修改、刪除时
        /// </summary>
        /// <param name="strRID"></param>
        public void PersoProjectSapAskMoney(string strRID, string strMateriel_Type)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strMateriel_Type;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 自動匯入失敗
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void MaterialDataInLost(string strRID, string strFactoryName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strFactoryName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 小計檔匯入時，格式檢查
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void SubTotalDataIn(string strRID, string strFactoryEN, string strErrorMsg)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[2];
            arg[0] = strFactoryEN;
            arg[1] = strErrorMsg;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 廠商庫存異動匯入時，格式不正確
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void FactoryStocksChange(string strRID, string strFactoryEN, string strErrorMsg)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[2];
            arg[0] = strFactoryEN;
            arg[1] = strErrorMsg;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }
        /// <summary>
        /// (替換前版面)廠商庫存異動匯入時，格式不正確
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void FactoryStocksChangeReplace(string strRID, string strFactoryEN, string strErrorMsg)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[2];
            arg[0] = strFactoryEN;
            arg[1] = strErrorMsg;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }
        public void DayMonitory(string strRID, string cardTypeName, string NType, decimal HNumber)
        {
            string strWarningContent = "";
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;
            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;
            PARAM pModel = GetParamModel(GlobalString.cardparamType.YType);
            if (pModel == null)
                return;
            if (HNumber < Convert.ToDecimal(pModel.Param_Name.Substring(0, pModel.Param_Name.Length - 1)))
            {
                object[] arg = new object[2];
                arg[0] = cardTypeName;
                arg[1] = NType;
                strWarningContent = string.Format(wcModel.Warning_Content, arg);
                GetWarningUser(wcModel, strWarningContent);
            }

        }

        public void MonthMonitory(string strRID, string cardTypeName, string NType, decimal HNumber)
        {
            string strWarningContent = "";
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            PARAM pModel = GetParamModel(GlobalString.cardparamType.NType);
            if (pModel == null)
                return;

            if (HNumber < Convert.ToDecimal(pModel.Param_Name.Substring(0, pModel.Param_Name.Length - 1)))
            {
                object[] arg = new object[2];
                arg[0] = cardTypeName;
                arg[1] = NType;
                strWarningContent = string.Format(wcModel.Warning_Content, arg);
                GetWarningUser(wcModel, strWarningContent);
            }
        }

        /// <summary>
        /// 日結時，XXXPerso廠XXX版面簡稱庫存不足
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void CardTypeNotEnough(string strRID, string strFactoryCN, string strCardTypeName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[2];
            arg[0] = strFactoryCN;
            arg[1] = strCardTypeName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 年度換卡預測檔格式不正確
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strErrorMsg"></param>
        public void YearChangeCardForeCast(string strRID, string strErrorMsg)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strErrorMsg;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 月度換卡預測檔格式不正確
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strErrorMsg"></param>
        public void MonthChangeCardForeCast(string strRID, string strErrorMsg)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strErrorMsg;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }

        /// <summary>
        /// 批次自動日結錯誤
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strErrorMsg"></param>
        public void BatchCompareNotPass(string strRID,
                        string strDate,
                        string strFactroyCN,
                        string strName)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[3];
            arg[0] = strDate;
            arg[1] = strFactroyCN;
            arg[2] = strName;

            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }
        /// <summary>
        /// 日結時，(替換前版面)廠商庫存異動核對有誤
        /// </summary>
        /// <param name="strRID"></param>
        /// <param name="strFactoryName"></param>
        public void BatchCompareFactoryReplace(string strRID, string strDate)
        {
            WARNING_CONFIGURATION wcModel = GetModel(strRID);
            if (wcModel == null)
                return;

            if (wcModel.System_Show == "N" && wcModel.Mail_Show == "N")
                return;

            object[] arg = new object[1];
            arg[0] = strDate;


            string strWarningContent = string.Format(wcModel.Warning_Content, arg);

            GetWarningUser(wcModel, strWarningContent);
        }
    }
}