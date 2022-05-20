﻿using System;
using System.Data;
using System.Configuration;


/// <summary>
/// FORE_CHANGE_CARD
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
public class FORE_CHANGE_CARD
{
    public FORE_CHANGE_CARD()
    { }

    #region Model
    private int _RID;
    private string _Change_Date;
    private string _Type;
    private string _Photo;
    private string _Affinity;
    private long _Number;
    private string _RCU;
    private string _RUU;
    private DateTime _RCT = Convert.ToDateTime("1900-01-01");
    private DateTime _RUT = Convert.ToDateTime("1900-01-01");
    private string _RST;
    private string _IsMonth;
    private string _IsYear;

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
    public string Change_Date
    {
        get { return _Change_Date; }
        set { _Change_Date = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Type
    {
        get { return _Type; }
        set { _Type = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Photo
    {
        get { return _Photo; }
        set { _Photo = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Affinity
    {
        get { return _Affinity; }
        set { _Affinity = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public long Number
    {
        get { return _Number; }
        set { _Number = value; }
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
    public string IsMonth
    {
        get { return _IsMonth; }
        set { _IsMonth = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string IsYear
    {
        get { return _IsYear; }
        set { _IsYear = value; }
    }

    #endregion


}

