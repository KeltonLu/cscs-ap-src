using System;
using System.Data;
using System.Configuration;



/// <summary>
/// PERSO_PROJECT_DETAIL
/// Date Created: 2008年12月15日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class PERSO_PROJECT_DETAIL
    {
        public PERSO_PROJECT_DETAIL()
        { }

        #region Model
        private int _RID;
        private DateTime _Use_Date = Convert.ToDateTime("1900-01-01");
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;
        private decimal _Sum;
        private int _Perso_Factory_RID;
        private int _Card_Group_RID;
        private int _CardType_RID;
        private int _Number;
        private int _Project_RID;
        private decimal _Unit_Price;

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
        public DateTime Use_Date
        {
            get { return _Use_Date; }
            set { _Use_Date = value; }
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
        public decimal Sum
        {
            get { return _Sum; }
            set { _Sum = value; }
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
        public int Card_Group_RID
        {
            get { return _Card_Group_RID; }
            set { _Card_Group_RID = value; }
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
        public int Number
        {
            get { return _Number; }
            set { _Number = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Project_RID
        {
            get { return _Project_RID; }
            set { _Project_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public decimal Unit_Price
        {
            get { return _Unit_Price; }
            set { _Unit_Price = value; }
        }

        #endregion


    }
}

