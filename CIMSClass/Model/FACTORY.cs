using System;
using System.Data;
using System.Configuration;


/// <summary>
/// FACTORY
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
/// 
namespace CIMSClass.Model
{
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

        public string Is_Perso
        {
            get
            {
                return this._Is_Perso;
            }
            set
            {
                this._Is_Perso = value;
            }
        }

        public string Is_Blank
        {
            get
            {
                return this._Is_Blank;
            }
            set
            {
                this._Is_Blank = value;
            }
        }

        public string Factory_ID
        {
            get
            {
                return this._Factory_ID;
            }
            set
            {
                this._Factory_ID = value;
            }
        }

        public string Unit_ID
        {
            get
            {
                return this._Unit_ID;
            }
            set
            {
                this._Unit_ID = value;
            }
        }

        public string Factory_Name
        {
            get
            {
                return this._Factory_Name;
            }
            set
            {
                this._Factory_Name = value;
            }
        }

        public string Factory_ShortName_CN
        {
            get
            {
                return this._Factory_ShortName_CN;
            }
            set
            {
                this._Factory_ShortName_CN = value;
            }
        }

        public string Factory_ShortName_EN
        {
            get
            {
                return this._Factory_ShortName_EN;
            }
            set
            {
                this._Factory_ShortName_EN = value;
            }
        }

        public string Factory_Principal
        {
            get
            {
                return this._Factory_Principal;
            }
            set
            {
                this._Factory_Principal = value;
            }
        }

        public DateTime Creat_Date
        {
            get
            {
                return this._Creat_Date;
            }
            set
            {
                this._Creat_Date = value;
            }
        }

        public string CoLinkMan_Name
        {
            get
            {
                return this._CoLinkMan_Name;
            }
            set
            {
                this._CoLinkMan_Name = value;
            }
        }

        public string CoLinkMan_Mobil
        {
            get
            {
                return this._CoLinkMan_Mobil;
            }
            set
            {
                this._CoLinkMan_Mobil = value;
            }
        }

        public string CoLinkMan_Phone
        {
            get
            {
                return this._CoLinkMan_Phone;
            }
            set
            {
                this._CoLinkMan_Phone = value;
            }
        }

        public string FacLinkMan_Name
        {
            get
            {
                return this._FacLinkMan_Name;
            }
            set
            {
                this._FacLinkMan_Name = value;
            }
        }

        public string FacLinkMan_Mobil
        {
            get
            {
                return this._FacLinkMan_Mobil;
            }
            set
            {
                this._FacLinkMan_Mobil = value;
            }
        }

        public string FacLinkMan_Phone
        {
            get
            {
                return this._FacLinkMan_Phone;
            }
            set
            {
                this._FacLinkMan_Phone = value;
            }
        }

        public string CoAddress
        {
            get
            {
                return this._CoAddress;
            }
            set
            {
                this._CoAddress = value;
            }
        }

        public string CoPhone
        {
            get
            {
                return this._CoPhone;
            }
            set
            {
                this._CoPhone = value;
            }
        }

        public string CoFax
        {
            get
            {
                return this._CoFax;
            }
            set
            {
                this._CoFax = value;
            }
        }

        public string CoLinkMan_Email
        {
            get
            {
                return this._CoLinkMan_Email;
            }
            set
            {
                this._CoLinkMan_Email = value;
            }
        }

        public string FacAddress
        {
            get
            {
                return this._FacAddress;
            }
            set
            {
                this._FacAddress = value;
            }
        }

        public string FacPhone
        {
            get
            {
                return this._FacPhone;
            }
            set
            {
                this._FacPhone = value;
            }
        }

        public string FacFax
        {
            get
            {
                return this._FacFax;
            }
            set
            {
                this._FacFax = value;
            }
        }

        public string FacLinkMan_Email
        {
            get
            {
                return this._FacLinkMan_Email;
            }
            set
            {
                this._FacLinkMan_Email = value;
            }
        }

        public string Country_Name
        {
            get
            {
                return this._Country_Name;
            }
            set
            {
                this._Country_Name = value;
            }
        }

        public string Is_Cooperate
        {
            get
            {
                return this._Is_Cooperate;
            }
            set
            {
                this._Is_Cooperate = value;
            }
        }

        public string Product_Main
        {
            get
            {
                return this._Product_Main;
            }
            set
            {
                this._Product_Main = value;
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

