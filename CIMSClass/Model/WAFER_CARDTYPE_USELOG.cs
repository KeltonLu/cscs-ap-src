using System;
using System.Data;
using System.Configuration;



/// <summary>
/// WAFER_CARDTYPE_USELOG
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class WAFER_CARDTYPE_USELOG
    {
        public WAFER_CARDTYPE_USELOG()
        { }

        #region Model
        private int _RID;
        private DateTime _Income_Date = Convert.ToDateTime("1900-01-01");
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;
        private int _Usable_Number;
        private int _Number;
        private int _Factory_RID;
        private int _CardType_RID;
        private DateTime _Begin_Date = Convert.ToDateTime("1900-01-01");
        private DateTime _End_Date = Convert.ToDateTime("1900-01-01");
        private int _Wafer_RID;
        private int _Operate_RID;
        private int _CardType_Move_RID;
        private string _Operate_Type;
        private decimal _Unit_Price;

        /// <summary>
        /// 
        /// </summary>
        public decimal Unit_Price
        {
            get { return _Unit_Price; }
            set { _Unit_Price = value; }
        }

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
        public DateTime Income_Date
        {
            get { return _Income_Date; }
            set { _Income_Date = value; }
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
        public int Usable_Number
        {
            get { return _Usable_Number; }
            set { _Usable_Number = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Number
        {
            get { return _Number; }
            set { _Number = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Factory_RID
        {
            get { return _Factory_RID; }
            set { _Factory_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int CardType_RID
        {
            get { return _CardType_RID; }
            set { _CardType_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime Begin_Date
        {
            get { return _Begin_Date; }
            set { _Begin_Date = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime End_Date
        {
            get { return _End_Date; }
            set { _End_Date = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Wafer_RID
        {
            get { return _Wafer_RID; }
            set { _Wafer_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Operate_RID
        {
            get { return _Operate_RID; }
            set { _Operate_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int CardType_Move_RID
        {
            get { return _CardType_Move_RID; }
            set { _CardType_Move_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Operate_Type
        {
            get { return _Operate_Type; }
            set { _Operate_Type = value; }
        }

        #endregion


    }
}

