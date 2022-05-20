using System;
using System.Data;
using System.Configuration;


/// <summary>
/// CARDTYPE_STOCKS
/// Date Created: 2008年12月12日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class CARDTYPE_STOCKS
    {
        public CARDTYPE_STOCKS()
        { }

        #region Model
        private int _RID;

        private DateTime _Stock_Date = Convert.ToDateTime("1900-01-01");

        private string _RCU;

        private string _RUU;

        private DateTime _RCT = Convert.ToDateTime("1900-01-01");

        private DateTime _RUT = Convert.ToDateTime("1900-01-01");

        private string _RST;

        private int _Stocks_Number;

        private int _Perso_Factory_RID;

        private int _CardType_RID;

        public int RID
        {
            get
            {
                return this._RID;
            }
            set
            {
                this._RID = value;
            }
        }

        public DateTime Stock_Date
        {
            get
            {
                return this._Stock_Date;
            }
            set
            {
                this._Stock_Date = value;
            }
        }

        public string RCU
        {
            get
            {
                return this._RCU;
            }
            set
            {
                this._RCU = value;
            }
        }

        public string RUU
        {
            get
            {
                return this._RUU;
            }
            set
            {
                this._RUU = value;
            }
        }

        public DateTime RCT
        {
            get
            {
                return this._RCT;
            }
            set
            {
                this._RCT = value;
            }
        }

        public DateTime RUT
        {
            get
            {
                return this._RUT;
            }
            set
            {
                this._RUT = value;
            }
        }

        public string RST
        {
            get
            {
                return this._RST;
            }
            set
            {
                this._RST = value;
            }
        }

        public int Stocks_Number
        {
            get
            {
                return this._Stocks_Number;
            }
            set
            {
                this._Stocks_Number = value;
            }
        }

        public int Perso_Factory_RID
        {
            get
            {
                return this._Perso_Factory_RID;
            }
            set
            {
                this._Perso_Factory_RID = value;
            }
        }

        public int CardType_RID
        {
            get
            {
                return this._CardType_RID;
            }
            set
            {
                this._CardType_RID = value;
            }
        }

        #endregion


    }
}

