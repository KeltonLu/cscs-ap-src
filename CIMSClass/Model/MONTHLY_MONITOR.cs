using System;
using System.Data;
using System.Configuration;

namespace CIMSClass.Model
{
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

        public string TYPE
        {
            get
            {
                return this._TYPE;
            }
            set
            {
                this._TYPE = value;
            }
        }

        public string AFFINITY
        {
            get
            {
                return this._AFFINITY;
            }
            set
            {
                this._AFFINITY = value;
            }
        }

        public string PHOTO
        {
            get
            {
                return this._PHOTO;
            }
            set
            {
                this._PHOTO = value;
            }
        }

        public string Name
        {
            get
            {
                return this._Name;
            }
            set
            {
                this._Name = value;
            }
        }

        public int XType
        {
            get
            {
                return this._XType;
            }
            set
            {
                this._XType = value;
            }
        }

        public int NType
        {
            get
            {
                return this._NType;
            }
            set
            {
                this._NType = value;
            }
        }

        public DateTime CMonth
        {
            get
            {
                return this._CMonth;
            }
            set
            {
                this._CMonth = value;
            }
        }

        public int ANumber
        {
            get
            {
                return this._ANumber;
            }
            set
            {
                this._ANumber = value;
            }
        }

        public int BNumber
        {
            get
            {
                return this._BNumber;
            }
            set
            {
                this._BNumber = value;
            }
        }

        public int D1Number
        {
            get
            {
                return this._D1Number;
            }
            set
            {
                this._D1Number = value;
            }
        }

        public int D2Number
        {
            get
            {
                return this._D2Number;
            }
            set
            {
                this._D2Number = value;
            }
        }

        public int ENumber
        {
            get
            {
                return this._ENumber;
            }
            set
            {
                this._ENumber = value;
            }
        }

        public int FNumber
        {
            get
            {
                return this._FNumber;
            }
            set
            {
                this._FNumber = value;
            }
        }

        public int GNumber
        {
            get
            {
                return this._GNumber;
            }
            set
            {
                this._GNumber = value;
            }
        }

        public int CNumber
        {
            get
            {
                return this._CNumber;
            }
            set
            {
                this._CNumber = value;
            }
        }

        public int JNumber
        {
            get
            {
                return this._JNumber;
            }
            set
            {
                this._JNumber = value;
            }
        }

        public int A1Number
        {
            get
            {
                return this._A1Number;
            }
            set
            {
                this._A1Number = value;
            }
        }

        public int G1Number
        {
            get
            {
                return this._G1Number;
            }
            set
            {
                this._G1Number = value;
            }
        }

        public decimal HNumber
        {
            get
            {
                return this._HNumber;
            }
            set
            {
                this._HNumber = value;
            }
        }

        public int KNumber
        {
            get
            {
                return this._KNumber;
            }
            set
            {
                this._KNumber = value;
            }
        }

        public decimal LNumber
        {
            get
            {
                return this._LNumber;
            }
            set
            {
                this._LNumber = value;
            }
        }

        public int Perso_Factory_Rid
        {
            get
            {
                return this._Perso_Factory_Rid;
            }
            set
            {
                this._Perso_Factory_Rid = value;
            }
        }

        #endregion

    }
}
