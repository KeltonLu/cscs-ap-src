using System;
using System.Data;
using System.Configuration;



/// <summary>
/// SUBTOTAL_IMPORT
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class SUBTOTAL_IMPORT
    {
        public SUBTOTAL_IMPORT()
        { }

        #region Model
        private int _RID;
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;
        private string _Action;
        private string _TYPE;
        private string _AFFINITY;
        private string _PHOTO;
        private long _Number;
        private DateTime _Date_Time = Convert.ToDateTime("1900-01-01");
        private int _Perso_Factory_RID;
        private int _MakeCardType_RID;
        private string _Import_FileName;
        private string _Is_Check;
        private DateTime _Check_Date = Convert.ToDateTime("1900-01-01");
        private int _Old_CardType_RID;

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
        public int Old_CardType_RID
        {
            get { return _Old_CardType_RID; }
            set { _Old_CardType_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Action
        {
            get { return _Action; }
            set { _Action = value; }
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
        public long Number
        {
            get { return _Number; }
            set { _Number = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime Date_Time
        {
            get { return _Date_Time; }
            set { _Date_Time = value; }
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
        public int MakeCardType_RID
        {
            get { return _MakeCardType_RID; }
            set { _MakeCardType_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Import_FileName
        {
            get { return _Import_FileName; }
            set { _Import_FileName = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Is_Check
        {
            get { return _Is_Check; }
            set { _Is_Check = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime Check_Date
        {
            get { return _Check_Date; }
            set { _Check_Date = value; }
        }

        #endregion


    }
}

