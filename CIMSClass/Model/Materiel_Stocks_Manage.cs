using System;
using System.Data;
using System.Configuration;



/// <summary>
/// MATERIEL_STOCKS_MANAGE
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class MATERIEL_STOCKS_MANAGE
    {
        public MATERIEL_STOCKS_MANAGE()
        { }

        #region Model
        private int _RID;
        private DateTime _Stock_Date = Convert.ToDateTime("1900-01-01");
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;
        private long _Number;
        private int _Perso_Factory_RID;
        private string _Type;
        private string _Comment;
        private string _Serial_Number;
        private long _Invenroty_Remain;
        private int _Real_Wear_Rate;

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
        public DateTime Stock_Date
        {
            get { return _Stock_Date; }
            set { _Stock_Date = value; }
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
        public long Number
        {
            get { return _Number; }
            set { _Number = value; }
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
        public string Type
        {
            get { return _Type; }
            set { _Type = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Comment
        {
            get { return _Comment; }
            set { _Comment = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Serial_Number
        {
            get { return _Serial_Number; }
            set { _Serial_Number = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public long Invenroty_Remain
        {
            get { return _Invenroty_Remain; }
            set { _Invenroty_Remain = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Real_Wear_Rate
        {
            get { return _Real_Wear_Rate; }
            set { _Real_Wear_Rate = value; }
        }
        #endregion


    }
}

