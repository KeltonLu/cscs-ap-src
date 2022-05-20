using System;
using System.Collections.Generic;
using System.Text;

public class CARDTYPE_STATUS
{
    public CARDTYPE_STATUS()
    { }

    #region Model
    private int _RID;
    private DateTime _RUT = Convert.ToDateTime("1900-01-01");
    private DateTime _RCT = Convert.ToDateTime("1900-01-01");
    private string _RUU;
    private string _RCU;
    private string _RST;
    private string _Status_Code;
    private string _Status_Name;
    private string _Is_UptDepository;
    private string _Operate;
    private string _Is_Delete;
    private string _Is_Display;
    private string _Comment;

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
    public DateTime RUT
    {
        get { return _RUT; }
        set { _RUT = value; }
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
    public string RUU
    {
        get { return _RUU; }
        set { _RUU = value; }
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
    public string RST
    {
        get { return _RST; }
        set { _RST = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Status_Code
    {
        get { return _Status_Code; }
        set { _Status_Code = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Status_Name
    {
        get { return _Status_Name; }
        set { _Status_Name = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Is_UptDepository
    {
        get { return _Is_UptDepository; }
        set { _Is_UptDepository = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Operate
    {
        get { return _Operate; }
        set { _Operate = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Is_Delete
    {
        get { return _Is_Delete; }
        set { _Is_Delete = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Is_Display
    {
        get { return _Is_Display; }
        set { _Is_Display = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Comment
    {
        get { return _Comment; }
        set { _Comment = value; }
    }

    #endregion


}
