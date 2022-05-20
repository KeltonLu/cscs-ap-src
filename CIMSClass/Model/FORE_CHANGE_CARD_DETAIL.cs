using System;
using System.Data;
using System.Configuration;


/// <summary>
/// FORE_CHANGE_CARD_DETAIL
/// Date Created: 2008¦~12¤ë11¤é
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class FORE_CHANGE_CARD_DETAIL
    {
        public FORE_CHANGE_CARD_DETAIL()
        { }

        #region Model
        private int _RID;
        private string _Change_Date;
        private string _Type;
        private string _Photo;
        private string _Affinity;
        private int _Perso_Factory_RID;
        private long _Number;
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
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
        public string Change_Date
        {
            get { return _Change_Date; }
            set { _Change_Date = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Type
        {
            get { return _Type; }
            set { _Type = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Photo
        {
            get { return _Photo; }
            set { _Photo = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Affinity
        {
            get { return _Affinity; }
            set { _Affinity = value; }
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

        #endregion


    }
}

