//******************************************************************
//*  �@    �̡GKunYang
//*  �\�໡���G�ϥΪ̸�ƺ��@�޿�
//*  �Ыؤ���G2009-07-08
//*  �ק����G2009-07-08 12:00
//*  �ק�O���G
//*            ��2008-07-08
//*              1.�Ы� ����
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
                //�ưȶ}�l
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
                    throw new Exception("�L�K�[�H��");
                }

                //�ާ@��x
                //SetOprLog();


                //�ưȴ���
                dao.Commit();
            }
            //catch (AlertException ex)
            //{
            //    throw ex;
            //}
            catch (Exception ex)
            {
                //�ưȦ^�u
                dao.Rollback();
                //���`�B�z
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("AddLDAP����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            finally
            {
                //�����s��
                dao.CloseConnection();
            }

        }

    }
}
