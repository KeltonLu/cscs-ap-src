//******************************************************************
//*  作    者：KunYang
//*  功能說明：使用者資料維護邏輯
//*  創建日期：2009-07-08
//*  修改日期：2009-07-08 12:00
//*  修改記錄：
//*            □2008-07-08
//*              1.創建 楊昆
//*******************************************************************
using System;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Web;
using System.Collections.Generic;
using System.Text;
using Microsoft.Practices.EnterpriseLibrary.Data;
using CIMSBatch.Model;
using CIMSBatch.FTP;
using CIMSBatch.Mail;
using CIMSBatch.Public;

namespace CIMSBatch.Business
{
    class BaseInfo042BL : BaseLogic
    {
        public void AddLDAP(DataSet dstUser)
        {
            if (dstUser.Tables.Count != 2)
                return;

            try
            {
                //事務開始
                dao.OpenConnection();

                if (dstUser.Tables[0].Rows.Count != 0)
                {

                    dao.ExecuteNonQuery("delete from [users] where UserID!='admin'");
                    dao.ExecuteNonQuery("delete from [USERSROLE] where UserID!='admin'");

                    foreach (DataRow drowUser in dstUser.Tables[0].Rows)
                    {
                        USERS uModel = new USERS();
                        uModel.UserID = drowUser[0].ToString();
                        uModel.UserName = drowUser[1].ToString();
                        uModel.Email = drowUser[2].ToString();
                        dao.Add<USERS>(uModel);                       
                    }

                    foreach (DataRow drowRole in dstUser.Tables[1].Rows)
                    {
                        USERSROLE urModel = new USERSROLE();
                        urModel.UserID = drowRole[0].ToString();
                        urModel.RoleID = drowRole[1].ToString();
                        dao.Add<USERSROLE>(urModel);
                    }
                }
                else
                {
                    throw new Exception("無添加信息");
                }

                //操作日誌
                //SetOprLog();


                //事務提交
                dao.Commit();
            }
            //catch (AlertException ex)
            //{
            //    throw ex;
            //}
            catch (Exception ex)
            {
                //事務回滾
                dao.Rollback();
                //異常處理
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("AddLDAP報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            finally
            {
                //關閉連接
                dao.CloseConnection();
            }

        }

    }
}
