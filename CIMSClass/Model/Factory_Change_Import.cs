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
namespace CIMSClass.Model
{
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

        private string _Mk_No;

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

        public string Is_Check
        {
            get
            {
                return this._Is_Check;
            }
            set
            {
                this._Is_Check = value;
            }
        }

        public int Status_RID
        {
            get
            {
                return this._Status_RID;
            }
            set
            {
                this._Status_RID = value;
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

        public int Number
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

        public string Space_Short_Name
        {
            get
            {
                return this._Space_Short_Name;
            }
            set
            {
                this._Space_Short_Name = value;
            }
        }

        public string Is_Auto_Import
        {
            get
            {
                return this._Is_Auto_Import;
            }
            set
            {
                this._Is_Auto_Import = value;
            }
        }

        public DateTime Date_Time
        {
            get
            {
                return this._Date_Time;
            }
            set
            {
                this._Date_Time = value;
            }
        }

        public DateTime Check_Date
        {
            get
            {
                return this._Check_Date;
            }
            set
            {
                this._Check_Date = value;
            }
        }

        public string Mk_No
        {
            get
            {
                return this._Mk_No;
            }
            set
            {
                this._Mk_No = value;
            }
        }
        #endregion


    }
}

