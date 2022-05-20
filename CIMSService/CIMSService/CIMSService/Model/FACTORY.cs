using System;
using System.Data;
using System.Configuration;


/// <summary>
/// FACTORY
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
public class FACTORY
{
    public FACTORY()
    { }

    #region Model
    private int _RID;
    private string _RCU;
    private string _RUU;
    private DateTime _RCT = Convert.ToDateTime("1900-01-01");
    private DateTime _RUT = Convert.ToDateTime("1900-01-01");
    private string _RST;
    private string _Is_Perso;
    private string _Is_Blank;
    private string _Factory_ID;
    private string _Unit_ID;
    private string _Factory_Name;
    private string _Factory_ShortName_CN;
    private string _Factory_ShortName_EN;
    private string _Factory_Principal;
    private DateTime _Creat_Date = Convert.ToDateTime("1900-01-01");
    private string _CoLinkMan_Name;
    private string _CoLinkMan_Mobil;
    private string _CoLinkMan_Phone;
    private string _FacLinkMan_Name;
    private string _FacLinkMan_Mobil;
    private string _FacLinkMan_Phone;
    private string _CoAddress;
    private string _CoPhone;
    private string _CoFax;
    private string _CoLinkMan_Email;
    private string _FacAddress;
    private string _FacPhone;
    private string _FacFax;
    private string _FacLinkMan_Email;
    private string _Country_Name;
    private string _Is_Cooperate;
    private string _Product_Main;
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
    public string Is_Perso
    {
        get { return _Is_Perso; }
        set { _Is_Perso = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Is_Blank
    {
        get { return _Is_Blank; }
        set { _Is_Blank = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Factory_ID
    {
        get { return _Factory_ID; }
        set { _Factory_ID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Unit_ID
    {
        get { return _Unit_ID; }
        set { _Unit_ID = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Factory_Name
    {
        get { return _Factory_Name; }
        set { _Factory_Name = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Factory_ShortName_CN
    {
        get { return _Factory_ShortName_CN; }
        set { _Factory_ShortName_CN = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Factory_ShortName_EN
    {
        get { return _Factory_ShortName_EN; }
        set { _Factory_ShortName_EN = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Factory_Principal
    {
        get { return _Factory_Principal; }
        set { _Factory_Principal = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public DateTime Creat_Date
    {
        get { return _Creat_Date; }
        set { _Creat_Date = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoLinkMan_Name
    {
        get { return _CoLinkMan_Name; }
        set { _CoLinkMan_Name = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoLinkMan_Mobil
    {
        get { return _CoLinkMan_Mobil; }
        set { _CoLinkMan_Mobil = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoLinkMan_Phone
    {
        get { return _CoLinkMan_Phone; }
        set { _CoLinkMan_Phone = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacLinkMan_Name
    {
        get { return _FacLinkMan_Name; }
        set { _FacLinkMan_Name = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacLinkMan_Mobil
    {
        get { return _FacLinkMan_Mobil; }
        set { _FacLinkMan_Mobil = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacLinkMan_Phone
    {
        get { return _FacLinkMan_Phone; }
        set { _FacLinkMan_Phone = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoAddress
    {
        get { return _CoAddress; }
        set { _CoAddress = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoPhone
    {
        get { return _CoPhone; }
        set { _CoPhone = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoFax
    {
        get { return _CoFax; }
        set { _CoFax = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string CoLinkMan_Email
    {
        get { return _CoLinkMan_Email; }
        set { _CoLinkMan_Email = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacAddress
    {
        get { return _FacAddress; }
        set { _FacAddress = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacPhone
    {
        get { return _FacPhone; }
        set { _FacPhone = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacFax
    {
        get { return _FacFax; }
        set { _FacFax = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string FacLinkMan_Email
    {
        get { return _FacLinkMan_Email; }
        set { _FacLinkMan_Email = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Country_Name
    {
        get { return _Country_Name; }
        set { _Country_Name = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Is_Cooperate
    {
        get { return _Is_Cooperate; }
        set { _Is_Cooperate = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public string Product_Main
    {
        get { return _Product_Main; }
        set { _Product_Main = value; }
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

