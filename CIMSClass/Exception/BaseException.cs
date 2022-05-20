//******************************************************************
//*  作    者：QingChen
//*  功能說明：可記錄的異常
//*  創建日期：2008/05/21
//*  修改日期：2008/05/21  16:59
//*  修改記錄：
//*            □2008/05/21 
//*              1.創建 陳青
//*******************************************************************
using System;
using System.Data;
using System.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Logging;

/// <summary>
///底層程序異常，一般用來保存底層程式中不向上抛出但需要記錄的異常，常常用於非關鍵的靜態方法異常處理
/// </summary>
namespace CIMSClass
{

    public class BaseException : ApplicationException
    {
        /// <summary>
        /// 建構函式
        /// </summary>
        /// <param name="message">異常信息</param>
        /// <param name="exception">導致本異常的異常</param>
        public BaseException(string message)
            : base(message)
        {

        }
        /// <summary>
        /// 建構函式
        /// </summary>
        /// <param name="message">異常信息</param>
        /// <param name="exception">導致本異常的異常</param>
        public BaseException(string message, Exception exception)
            : base(message, exception)
        {

        }

    }
}
