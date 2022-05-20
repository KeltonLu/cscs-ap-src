using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using CIMSClass;
using CIMSClass.Model;
using CIMSClass.FTP;
using CIMSClass.Mail;

/// <summary>
/// BaseLogic 的摘要描述
/// </summary>
namespace CIMSClass
{
    public class BaseLogic
    {
        public BaseLogic()
        {
            _dao = new DataBaseDAO();
        }

        /// <summary>
        /// 連接名稱
        /// </summary>
        /// <param name="connectionName"></param>
        public BaseLogic(string connectionName)
        {
            _dao = new DataBaseDAO(connectionName);
        }

        private DataBaseDAO _dao;

        public DataBaseDAO dao
        {
            get
            {
               return _dao;
            }
            set
            {
                _dao = value;
            }
        }

        /// <summary>
        /// 設置操作日誌參數
        /// </summary>
        /// 
        public string strActionID;
        public string strUserID;
        public string strUserName;       
        
        public void SetOprLogNull()
        {
            strActionID = null;
            strUserID = null;
            strUserName = null;
            ExceptionFactory.SetExceptionInfo();
        }
        public void SetOprLogActionID(string ActionID)
        {
            if (!StringUtil.IsEmpty(ActionID))
            {
                strActionID = ActionID;
                ExceptionFactory.action = ActionID;
            }
        }
        public void SetOprLogUser(string UserID, string UserName)
        {
            if (!StringUtil.IsEmpty(UserID))
            {
                strUserID = UserID;
            }
            if (!StringUtil.IsEmpty(UserName))
            {
                strUserName = UserName;
                ExceptionFactory.userinfo=UserName;
            }
        }       
        /// <summary>
        /// 操作日誌
        /// </summary>
        /// 
        public void SetOprLog()
        {
            ACTION_HISTORY ah = new ACTION_HISTORY();           

            if (!StringUtil.IsEmpty(strActionID))
            {
                ah.FunctionID = strActionID.Substring(0, 4);
                ah.Operate_Date = DateTime.Now;
                ah.Param_Code = strActionID.Substring(4);
                ah.UserID = strUserID;
                _dao.Add<ACTION_HISTORY>(ah, "RID");
            }
        }

        /// <summary>
        /// 操作日誌
        /// </summary>
        public void SetOprLog(string Param_Code)
        {
            ACTION_HISTORY ah = new ACTION_HISTORY();          

            if (!StringUtil.IsEmpty(strActionID))
            {
                ah.FunctionID = strActionID.Substring(0, 4);
                ah.Operate_Date = DateTime.Now;
                ah.Param_Code = Param_Code;
                if (!StringUtil.IsEmpty(strUserID)) ah.UserID = strUserID;                
                _dao.Add<ACTION_HISTORY>(ah, "RID");
            }
        }
        public void AddLog(string strType, string strFileName)
        {
            try
            {
                IMPORT_HISTORY ihModel = new IMPORT_HISTORY();
                ihModel.File_Name = strFileName;
                ihModel.File_Type = strType;
                ihModel.Import_Date = DateTime.Now;
                if (!StringUtil.IsEmpty(strUserID))
                {
                    ihModel.RCU = strUserID;
                    ihModel.RUU = strUserID;
                }
                dao.Add<IMPORT_HISTORY>(ihModel, "RID");
            }
            catch
            {
                throw new Exception("記錄歷史失敗");
            }
        }

    }
}
