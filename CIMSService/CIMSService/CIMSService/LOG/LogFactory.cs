using System;
using System.Data;
using System.Configuration;
using CIMSBatch;
using CIMSBatch.Public;

/// <summary>
/// LogFactory 的摘要描述



/// </summary>
public class LogFactory
{
    public static LogError ErrorLog;
    public static LogOperator OpLog;

    public static string ErrorLogTime;
    public static string OpLogTime;

    public LogFactory()
    {
        
    }

    private void CreatObject(string Category)
    {
        if (Category == GlobalString.LogType.ErrorCategory)
        {
            if (StringUtil.IsEmpty(ErrorLogTime))
                ErrorLogTime = DateTime.Now.ToString("yyyyMMdd");

            if (ErrorLogTime != DateTime.Now.ToString("yyyyMMdd"))
            {
                ErrorLogTime = DateTime.Now.ToString("yyyyMMdd");
                ErrorLog = new LogError();
            }


            if (ErrorLog == null)
                ErrorLog = new LogError();

        }
        else if (Category == GlobalString.LogType.OpLogCategory)
        {
            if (StringUtil.IsEmpty(OpLogTime))
                OpLogTime = DateTime.Now.ToString("yyyyMMdd");

            if (OpLogTime != DateTime.Now.ToString("yyyyMMdd"))
            {
                OpLogTime = DateTime.Now.ToString("yyyyMMdd");
                OpLog = new LogOperator();
            }

            if (OpLog == null)
                OpLog = new LogOperator();
        }
    }


    public static void Write(string message, string Category)
    {
        if (Category == GlobalString.LogType.ErrorCategory)
        {
            if (StringUtil.IsEmpty(ErrorLogTime))
                ErrorLogTime = DateTime.Now.ToString("yyyyMMdd");

            if (ErrorLogTime != DateTime.Now.ToString("yyyyMMdd"))
            {
                ErrorLogTime = DateTime.Now.ToString("yyyyMMdd");
                ErrorLog = new LogError();
            }


            if (ErrorLog == null)
                ErrorLog = new LogError();

            ErrorLog.Write(message);
        }
        else if (Category == GlobalString.LogType.OpLogCategory)
        {
            if (StringUtil.IsEmpty(OpLogTime))
                OpLogTime = DateTime.Now.ToString("yyyyMMdd");

            if (OpLogTime != DateTime.Now.ToString("yyyyMMdd"))
            {
                OpLogTime = DateTime.Now.ToString("yyyyMMdd");
                OpLog = new LogOperator();
            }

            if (OpLog == null)
                OpLog = new LogOperator();

            OpLog.Write(message);
        }
    }

}
