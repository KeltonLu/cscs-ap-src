//******************************************************************
//*  作    者：QingChen
//*  功能說明：異常生成器
//*  創建日期：2008/05/21
//*  修改日期：2008/07/25  
//*  修改記錄：
//*            □2008/05/21 
//*              1.創建 陳青
//*              2.修改 鮑方   
//*******************************************************************
using System;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Text;
using CIMSClass;
using CIMSClass.Model;
using CIMSClass.FTP;
using CIMSClass.Mail;
using CIMSClass.Business;

/// <summary>
/// ExceptionFactory 的摘要描述
/// </summary>
namespace CIMSClass
{
    public class ExceptionFactory
    {
        /// <summary>
        /// 將Command物件格式化成異常信息
        /// </summary>
        /// <param name="command">要格式化的Command物件</param>
        public static string FormatDbCommand(params DbCommand[] commands)
        {
            if (commands != null && commands.Length > 0)
            {
                StringBuilder strReturn = new StringBuilder();
                string strGo = (commands.Length == 1 ? "" : " go ");
                for (int cindex = 0; cindex < commands.Length; cindex++)
                {
                    string returnString = commands[cindex].CommandText;
                    if (commands[cindex].Parameters.Count == 0)
                    {
                        strReturn.Append(returnString);
                        strReturn.Append(strGo);
                        break;
                    }

                    for (int index = 0; index < commands[cindex].Parameters.Count; index++)//整理參數
                    {
                        if (commands[cindex].Parameters[index].Value != null && commands[cindex].Parameters[index].Value != DBNull.Value)
                        {
                            returnString = returnString.Replace(commands[cindex].Parameters[index].ParameterName, "'" + commands[cindex].Parameters[index].Value.ToString() + "'");
                        }
                        else
                        {
                            returnString = returnString.Replace(commands[cindex].Parameters[index].ParameterName, "null");
                        }
                    }
                    strReturn.Append(returnString);
                    strReturn.Append(strGo);
                }
                return strReturn.ToString();
            }
            else
            {
                return "";
            }
        }
        /// <summary>
        /// 設置異常參數
        /// </summary>
        /// 
        public static string userinfo;
        public static string action;       
        public static void SetExceptionInfo()
        {
            userinfo = "unkonw";
            action = "8888";          
        }
       
        /// <summary>
        /// 生成一個重訂向並記錄的異常
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="innerMessage">系統消息</param>
        /// <param name="command">SQL語句</param>
        public static void CreateCustomSaveException(string message, string innerMessage, DbCommand command)
        {
            LogFactory.Write(" User : " + userinfo + " Page :  Action : " + action + " ErrMessage : " + message + " InnerMessage:" + innerMessage + " DBAction : " + FormatDbCommand(command), GlobalString.LogType.ErrorCategory);
        }

        /// <summary>
        /// 生成一個重訂向並記錄的異常
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="innerMessage">系統消息</param>
        /// <param name="commands">SQL語句</param>
        public static void CreateCustomSaveException(string message, string innerMessage, params DbCommand[] commands)
        {
            LogFactory.Write(" User : " + userinfo + " Page :  Action : " + action + " ErrMessage : " + message + " InnerMessage :" + innerMessage + " DBAction : " + FormatDbCommand(commands), GlobalString.LogType.ErrorCategory);
        }



        /// <summary>
        /// 生成一個重訂向並記錄的異常
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="innerException">系統消息</param>
        /// <param name="commands">SQL語句</param>
        public static void CreateCustomSaveException(string message, Exception innerException, params DbCommand[] commands)
        {
            LogFactory.Write(" User : " + userinfo + " Page :  Action : " + action + " ErrMessage : " + message + " InnerMessage :" + innerException.Message + " DBAction : " + FormatDbCommand(commands), GlobalString.LogType.ErrorCategory);
        }




        /// <summary>
        /// 生成一個重訂向並記錄的異常
        /// </summary>
        /// <param name="message">消息</param>
        /// <param name="innerException">系統消息</param>
        public static void CreateCustomSaveException(string message, Exception innerException)
        {
            LogFactory.Write(" User : " + userinfo + " Page :  Action : " + action + " Message : " + message + "InnerMessage:" + innerException.Message, GlobalString.LogType.ErrorCategory);
        }


        /// <summary>
        /// 記錄操作日誌
        /// </summary>
        /// <param name="command">記錄的SQL語句</param>
        public static void CreateCustomOptSaveException(DbCommand command)
        {
            SetExceptionInfo();
            LogFactory.Write(" User : " + userinfo + " Page :  Action : " + action + " DBAction : " + FormatDbCommand(command), GlobalString.LogType.OpLogCategory);
        }


        /// <summary>
        /// 生成一個頁面警告異常
        /// </summary>
        /// <param name="page">顯示異常的頁面</param>
        /// <param name="userMessage">顯示的友好信息</param>
        public static AlertException CreateAlertException(Page page, string userMessage)
        {
            return new AlertException(page, userMessage);
        }
    }
}