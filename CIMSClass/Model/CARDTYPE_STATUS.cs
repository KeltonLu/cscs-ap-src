using System;

namespace CIMSClass.Model
{
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

        public string Status_Code
        {
            get
            {
                return this._Status_Code;
            }
            set
            {
                this._Status_Code = value;
            }
        }

        public string Status_Name
        {
            get
            {
                return this._Status_Name;
            }
            set
            {
                this._Status_Name = value;
            }
        }

        public string Is_UptDepository
        {
            get
            {
                return this._Is_UptDepository;
            }
            set
            {
                this._Is_UptDepository = value;
            }
        }

        public string Operate
        {
            get
            {
                return this._Operate;
            }
            set
            {
                this._Operate = value;
            }
        }

        public string Is_Delete
        {
            get
            {
                return this._Is_Delete;
            }
            set
            {
                this._Is_Delete = value;
            }
        }

        public string Is_Display
        {
            get
            {
                return this._Is_Display;
            }
            set
            {
                this._Is_Display = value;
            }
        }

        public string Comment
        {
            get
            {
                return this._Comment;
            }
            set
            {
                this._Comment = value;
            }
        }

        #endregion


    }
}
