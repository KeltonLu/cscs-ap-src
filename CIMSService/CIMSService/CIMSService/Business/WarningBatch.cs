using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using CIMSBatch.Public;

namespace CIMSBatch.Business
{
    class WarningBatch : BaseLogic
    {
        #region SQL
        // 物料寄卡單警訊
        public const string SQL_SEL_WarningMateriel002 = @"select RID,Name,Maturity_Date from CARD_EXPONENT 
                                                            where datediff(d,getdate(),Maturity_Date) = 15 or datediff(d,getdate(),Maturity_Date) = 30 
                                                            order by RID";

        // 物料DM警訊
        public const string SQL_SEL_WarningMateriel003 = @"select RID,Name,End_Date from DMTYPE_INFO 
                                                            where (datediff(d,getdate(),End_Date) = 15 or datediff(d,getdate(),End_Date) = 30) and Is_Using = 'N' 
                                                            order by RID";
        #endregion SQL

        /// <summary>
        /// 主入口
        /// </summary>
        public void BatchStart()
        {
            try
            {
                WarningMateriel002();
                WarningMateriel003();
            }
            catch(Exception ex)
            {
                LogFactory.Write("主入口BatchStart方法保錯報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
           
        }

        /// <summary>
        /// 物料寄卡單警訊
        /// </summary>
        private void WarningMateriel002()
        {
            DataTable dtb = new DataTable();

            try
            {
                dtb = dao.GetList(SQL_SEL_WarningMateriel002).Tables[0];

                if (dtb.Rows.Count > 0)
                {
                    for (int i = 0; i < dtb.Rows.Count; i++)
                    {
                        Warning.SetWarning(GlobalString.WarningType.BatchWarningMateriel002And003, new object[2] { dtb.Rows[i]["Name"].ToString(), DateTime.Parse(dtb.Rows[i]["Maturity_Date"].ToString()).ToString("yyyy年MM月dd日") });
                    }
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("物料寄卡單警訊, WarningMateriel002報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
            }
        }

        /// <summary>
        /// 物料DM警訊
        /// </summary>
        private void WarningMateriel003()
        {
            DataTable dtb = new DataTable();

            try
            {
                dtb = dao.GetList(SQL_SEL_WarningMateriel003).Tables[0];

                if (dtb.Rows.Count > 0)
                {
                    for (int i = 0; i < dtb.Rows.Count; i++)
                    {
                        Warning.SetWarning(GlobalString.WarningType.BatchWarningMateriel002And003, new object[2] { dtb.Rows[i]["Name"].ToString(), DateTime.Parse(dtb.Rows[i]["End_Date"].ToString()).ToString("yyyy年MM月dd日") });
                    }
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("物料DM警訊, WarningMateriel003報錯:" + ex.Message, GlobalString.LogType.ErrorCategory);
            }
        }

    }
}
