using System;
using System.Collections.Generic;
using System.Text;

namespace CIMSClass.Model
{
    public class EXPRESSIONS_DEFINE
    {
        public EXPRESSIONS_DEFINE()
        { }

        #region Model
        private int _RID;
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private string _RUU;
        private string _RCU;
        private string _RST;
        private int _Expressions_RID;
        private int _Type_RID;
        private string _Operate;
        private string _Expressions_Name;

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
        public DateTime RUT
        {
            get { return _RUT; }
            set { _RUT = value; }
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
        public string RCU
        {
            get { return _RCU; }
            set { _RCU = value; }
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
        public int Expressions_RID
        {
            get { return _Expressions_RID; }
            set { _Expressions_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Type_RID
        {
            get { return _Type_RID; }
            set { _Type_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Operate
        {
            get { return _Operate; }
            set { _Operate = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Expressions_Name
        {
            get { return _Expressions_Name; }
            set { _Expressions_Name = value; }
        }

        #endregion


    }
}
