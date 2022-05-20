using System;
using System.Data;
using System.Configuration;

public class MONTHLY_MONITOR
{
    public MONTHLY_MONITOR() 
    { }
    #region Model
    private int _RID;
    private string _RUU;
    private DateTime _RUT = Convert.ToDateTime("1900-01-01");
    private DateTime _RCT = Convert.ToDateTime("1900-01-01");
    private string _RCU;
    private string _RST;
    private string _TYPE;
    private string _AFFINITY;
    private string _PHOTO;
    private string _Name;
    private int _XType;
    private int _NType;
    private DateTime _CMonth = Convert.ToDateTime("1900-01-01");
    private int _ANumber;
    private int _BNumber;
    private int _D1Number;
    private int _D2Number;
    private int _ENumber;
    private int _FNumber;
    private int _GNumber;
    private int _CNumber;
    private int _JNumber;
    private int _A1Number;
    private int _G1Number;
    private decimal _HNumber;
    private int _KNumber;
    private decimal _LNumber;
    private int _Perso_Factory_Rid;

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
    public string RUU
    {
        get { return _RUU; }
        set { _RUU = value; }
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
    public string Name
    {
        get { return _Name; }
        set { _Name = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int XType
    {
        get { return _XType; }
        set { _XType = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int NType
    {
        get { return _NType; }
        set { _NType = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public DateTime CMonth
    {
        get { return _CMonth; }
        set { _CMonth = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int ANumber
    {
        get { return _ANumber; }
        set { _ANumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int BNumber
    {
        get { return _BNumber; }
        set { _BNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int D1Number
    {
        get { return _D1Number; }
        set { _D1Number = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int D2Number
    {
        get { return _D2Number; }
        set { _D2Number = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int ENumber
    {
        get { return _ENumber; }
        set { _ENumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int FNumber
    {
        get { return _FNumber; }
        set { _FNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int GNumber
    {
        get { return _GNumber; }
        set { _GNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int CNumber
    {
        get { return _CNumber; }
        set { _CNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int JNumber
    {
        get { return _JNumber; }
        set { _JNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int A1Number
    {
        get { return _A1Number; }
        set { _A1Number = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int G1Number
    {
        get { return _G1Number; }
        set { _G1Number = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public decimal HNumber
    {
        get { return _HNumber; }
        set { _HNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int KNumber
    {
        get { return _KNumber; }
        set { _KNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public decimal LNumber
    {
        get { return _LNumber; }
        set { _LNumber = value; }
    }

    /// <summary>
    /// 
    /// </summary>
    public int Perso_Factory_Rid
    {
        get { return _Perso_Factory_Rid; }
        set { _Perso_Factory_Rid = value; }
    }

    #endregion
    
}
