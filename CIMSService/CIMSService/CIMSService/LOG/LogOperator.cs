using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using Microsoft.Practices.EnterpriseLibrary.Common.Configuration;
using Microsoft.Practices.EnterpriseLibrary.Logging.Filters;
using Microsoft.Practices.EnterpriseLibrary.Logging.Formatters;
using Microsoft.Practices.EnterpriseLibrary.Logging.TraceListeners;
using System.Configuration;
using CIMSBatch;
using CIMSBatch.Public;

public class LogOperator
{
    public static LogWriter _writer;


    // 模版
    const string Template = "Timestamp: {timestamp}{newline}" +
                            "Message: {message}{newline}" +
                            "Category: {category}{newline}" +
                            "Machine: {machine}{newline}";

    public LogOperator()
    {
        
        string LogFilePath = ConfigurationManager.AppSettings["LogPath"] + "Operator" + DateTime.Now.ToString("yyyyMMdd") + ".log";


        // 实例化一个TextFormatter，使用前面定义的模版
        TextFormatter formatter = new TextFormatter(Template);

        // 实例化TraceListener，记录到文本文件用FlatFileTraceListener
        //FlatFileTraceListener logFileListener =
        //    new FlatFileTraceListener(LogFilePath,
        //                               "----------",
        //                               "----------",
        //                               formatter);

        // 这里是TraceListener的集合，可以增加多个
        //LogSource mainLogSource = new LogSource("MainLogSource", SourceLevels.All);
        //mainLogSource.Listeners.Add(logFileListener);
        #region // 因企業庫升級為6.0, 調整寫法  Legend 2016/10/26

        List<TraceListener> logFileListeners = new List<TraceListener>();
        //logFileListeners.Add(logFileListener);

        LogSource mainLogSource = new LogSource("MainLogSource", logFileListeners, SourceLevels.All);

        #endregion

        IDictionary<string, LogSource> traceSources = new Dictionary<string, LogSource>();
        traceSources.Add(GlobalString.LogType.ErrorCategory, mainLogSource);
        traceSources.Add(GlobalString.LogType.OpLogCategory, mainLogSource);
        traceSources.Add(GlobalString.LogType.DebugCategory, mainLogSource);

        // 用来表示不记录日志，这点需要注意一下


        LogSource nonExistantLogSource = new LogSource("Empty");

        // 创建一个类别过滤器
        ICollection<string> categoryfilters = new List<string>();
        categoryfilters.Add(GlobalString.LogType.DebugCategory);

        CategoryFilter categoryFilter
            = new CategoryFilter("CategoryFilter", categoryfilters, CategoryFilterMode.AllowAllExceptDenied);

        // 加入类别过滤器到集合中


        ICollection<ILogFilter> filters = new List<ILogFilter>();
        filters.Add(categoryFilter);

        _writer = new LogWriter(filters,
                        traceSources,
                        nonExistantLogSource,
                        nonExistantLogSource,
                        mainLogSource,
                        GlobalString.LogType.OpLogCategory,
                        false,
                        true);
    }


    /// <summary>
    /// 记录日志信息到特定类别


    /// </summary>
    /// <param name="message">日志信息</param>
    /// <param name="category">类别</param>
    public void Write(string message)
    {
        LogEntry entry = new LogEntry();
        entry.TimeStamp = DateTime.Now;
        entry.Categories.Add(GlobalString.LogType.OpLogCategory);
        entry.Message = message;

        _writer.Write(entry);
    }
}