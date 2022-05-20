using System;
using System.Data;
using System.Configuration;
//using System.Web;
//using System.Web.Security;
//using System.Web.UI;
//using System.Web.UI.WebControls;
//using System.Web.UI.WebControls.WebParts;
//using System.Web.UI.HtmlControls;

/// <summary>
/// BaseLogic 的摘要描述

/// </summary>
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
            //if (_dao == null)
            //    _dao=new DataBaseDAO();
            return _dao;
        }
        set
        {
            _dao= value;
        }
    }
    
}
