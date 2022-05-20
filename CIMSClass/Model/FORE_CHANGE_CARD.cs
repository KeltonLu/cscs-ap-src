using System;
using System.Data;
using System.Configuration;


/// <summary>
/// FORE_CHANGE_CARD
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
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

        public string Change_Date
        {
            get
            {
                return this._Change_Date;
            }
            set
            {
                this._Change_Date = value;
            }
        }

        public string Type
        {
            get
            {
                return this._Type;
            }
            set
            {
                this._Type = value;
            }
        }

        public string Photo
        {
            get
            {
                return this._Photo;
            }
            set
            {
                this._Photo = value;
            }
        }

        public string Affinity
        {
            get
            {
                return this._Affinity;
            }
            set
            {
                this._Affinity = value;
            }
        }

        public long Number
        {
            get
            {
                return this._Number;
            }
            set
            {
                this._Number = value;
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

        public string IsMonth
        {
            get
            {
                return this._IsMonth;
            }
            set
            {
                this._IsMonth = value;
            }
        }

        public string IsYear
        {
            get
            {
                return this._IsYear;
            }
            set
            {
                this._IsYear = value;
            }
        }

        #endregion


    }
}

