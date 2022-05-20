using System;
using System.Data;
using System.Configuration;
using System.Web;

/// <summary>
/// CARD_GROUP
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class CARD_GROUP
    {
        public CARD_GROUP()
        { }

        #region Model
        private int _RID;
        private string _Group_Name;
        private string _Param_Code;
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
        public string Group_Name
        {
            get { return _Group_Name; }
            set { _Group_Name = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Param_Code
        {
            get { return _Param_Code; }
            set { _Param_Code = value; }
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

