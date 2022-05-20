//******************************************************************
//*  作    者：QingChen
//*  功能說明：自動警告的異常
//*  創建日期：2008/05/21
//*  修改日期：2008/05/21  16:59
//*  修改記錄：
//*            □2008/05/21 
//*              1.創建 陳青
//*******************************************************************
using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// AlertException 的摘要描述
/// </summary>
namespace CIMSClass
{
    public class AlertException : ApplicationException
    {
        /// <summary>
        /// 建構函式
        /// </summary>
        /// <param name="message">異常信息</param>
        /// <param name="exception">導致本異常的異常</param>
        public AlertException(string message)
            : base(message)
        {

        }

        public AlertException(Page alertPage, string userMessage)
        {
            ScriptManager.RegisterStartupScript(alertPage, alertPage.GetType(), Guid.NewGuid().ToString(), "alert('" + userMessage.Replace("'", " ") + "');", true);
        }

        public AlertException(Page alertPage, string message, string userMessage, Exception innerException)
            : base(message, innerException)
        {
            ScriptManager.RegisterStartupScript(alertPage, alertPage.GetType(), Guid.NewGuid().ToString(), "alert('" + userMessage.Replace("'", " ") + "');", true);
        }
    }
}
