using System;
using System.Data;
using System.Configuration;
using System.Web;


/// <summary>
/// MATERIEL_STOCKS_MOVE
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
public class MATERIEL_STOCKS_MOVE
{
    public MATERIEL_STOCKS_MOVE()
    { }

    #region Model
    private int _RID;
    private DateTime _Move_Date = Convert.ToDateTime("1900-01-01");
    private string _RCU;
    private string _RUU;
    private DateTime _RCT = Convert.ToDateTime("1900-01-01");
    private DateTime _RUT = Convert.ToDateTime("1900-01-01");
    private string _RST;
    private int _Move_Number;
    private int _From_Factory_RID;
    private int _To_Factory_RID;
    private string _Move_ID;
    private string _Serial_Number;

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
    public DateTime Move_Date
    {
        get { return _Move_Date; }
        set { _Move_Date = value; }
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
    public int Move_Number
    {
        get { return _Move_Number; }
        set { _Move_Number = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int From_Factory_RID
    {
        get { return _From_Factory_RID; }
        set { _From_Factory_RID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int To_Factory_RID
    {
        get { return _To_Factory_RID; }
        set { _To_Factory_RID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Move_ID
    {
        get { return _Move_ID; }
        set { _Move_ID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Serial_Number
    {
        get { return _Serial_Number; }
        set { _Serial_Number = value; }
    }

    #endregion


}

