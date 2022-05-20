using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using CIMSBatch.Public;

namespace CIMSBatch.Business
{
    class DataImport_B010BL : BaseLogic
    {
        public bool ImportHistoryStocks()
        {
            bool bRtn = true;
            try
            {
                dao.OpenConnection();

                Dictionary<string, object> dirValues = new Dictionary<string, object>();
                dirValues.Clear();
                dirValues.Add("dateFrom", ConfigurationManager.AppSettings["HistoryDateFrom"]);
                dirValues.Add("dateTo", ConfigurationManager.AppSettings["HistoryDateTo"]);

                dao.ExecuteNonQuery("proc_DATAIMPORT_B010", dirValues, true);

                //�ưȴ���
                dao.Commit();
            }
            catch (Exception ex)
            {
                dao.Rollback();
                bRtn = false;
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("ImportHistoryStocks����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
            finally
            {
                dao.CloseConnection();
            }

            return bRtn;
        }

    }
}
