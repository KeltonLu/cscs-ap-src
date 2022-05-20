using System;
using System.Data;
using System.Configuration;



/// <summary>
/// SPECIAL_PERSO_PROJECT_IMPORT
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class SPECIAL_PERSO_PROJECT_IMPORT
    {
        public SPECIAL_PERSO_PROJECT_IMPORT()
        { }

        #region Model
        private int _RID;
        private int _Perso_Factory_RID;
        private DateTime _Project_Date = Convert.ToDateTime("1900-01-01");
        private int _PersoProject_RID;
        private long _Number;
        private string _RCU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private string _RUU;
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;

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
        public int Perso_Factory_RID
        {
            get { return _Perso_Factory_RID; }
            set { _Perso_Factory_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime Project_Date
        {
            get { return _Project_Date; }
            set { _Project_Date = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int PersoProject_RID
        {
            get { return _PersoProject_RID; }
            set { _PersoProject_RID = value; }
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
        public DateTime RCT
        {
            get { return _RCT; }
            set { _RCT = value; }
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
        public string RST
        {
            get { return _RST; }
            set { _RST = value; }
        }

        #endregion


    }
}

