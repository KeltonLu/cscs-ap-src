using System;
using System.Data;
using System.Configuration;
using System.Web;
//using System.Web.Security;
//using System.Web.UI;
//using System.Web.UI.WebControls;
//using System.Web.UI.WebControls.WebParts;
//using System.Web.UI.HtmlControls;


/// <summary>
/// FACTORY_CHANGE_IMPORT
/// Date Created: 2008年9月22日
/// Created By:  FangBao
/// </summary>
public class FACTORY_CHANGE_IMPORT
{
    public FACTORY_CHANGE_IMPORT()
    { }

    #region Model
    private int _RID;
    private string _RCU;
    private string _RUU;
    private DateTime _RCT = Convert.ToDateTime("1900-01-01");
    private DateTime _RUT = Convert.ToDateTime("1900-01-01");
    private string _RST;
    private int _Perso_Factory_RID;
    private string _Is_Check;
    private int _Status_RID;
    private string _TYPE;
    private string _AFFINITY;
    private string _PHOTO;
    private int _Number;
    private string _Space_Short_Name;
    private string _Is_Auto_Import;
    private DateTime _Date_Time = Convert.ToDateTime("1900-01-01");
    private DateTime _Check_Date = Convert.ToDateTime("1900-01-01");

    /// <summary>
    /// 
    /// </summary>
    public int RID
    {
        get { return _RID; }
        set { _RID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string RCU
    {
        get { return _RCU; }
        set { _RCU = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string RUU
    {
        get { return _RUU; }
        set { _RUU = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public DateTime RCT
    {
        get { return _RCT; }
        set { _RCT = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public DateTime RUT
    {
        get { return _RUT; }
        set { _RUT = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string RST
    {
        get { return _RST; }
        set { _RST = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int Perso_Factory_RID
    {
        get { return _Perso_Factory_RID; }
        set { _Perso_Factory_RID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Is_Check
    {
        get { return _Is_Check; }
        set { _Is_Check = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int Status_RID
    {
        get { return _Status_RID; }
        set { _Status_RID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string TYPE
    {
        get { return _TYPE; }
        set { _TYPE = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string AFFINITY
    {
        get { return _AFFINITY; }
        set { _AFFINITY = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string PHOTO
    {
        get { return _PHOTO; }
        set { _PHOTO = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int Number
    {
        get { return _Number; }
        set { _Number = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Space_Short_Name
    {
        get { return _Space_Short_Name; }
        set { _Space_Short_Name = value; }
    }
    /// <summary>
    /// 
    /// </summary>
    public string Is_Auto_Import
    {
        get { return _Is_Auto_Import; }
        set { _Is_Auto_Import = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public DateTime Date_Time
    {
        get { return _Date_Time; }
        set { _Date_Time = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public DateTime Check_Date
    {
        get { return _Check_Date; }
        set { _Check_Date = value; }
    }

    #endregion


}

