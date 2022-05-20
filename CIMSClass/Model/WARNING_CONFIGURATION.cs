using System;
using System.Collections.Generic;
using System.Text;

namespace CIMSClass.Model
{
    internal class WARNING_CONFIGURATION
    {
        #region WARNING_CONFIGURATION Model
        private int _RID;
        private string _Item_Name;
        private string _Condition;
        private string _Warning_Content;
        private string _System_Show;
        private string _Mail_Show;
        private string _RST;
        private DateTime  _RUT = Convert.ToDateTime("1900-01-01");
        private DateTime  _RCT = Convert.ToDateTime("1900-01-01");
        private string _RUU;
        private string _RCU;

        public int RID
        {
            get { return _RID; }
            set { _RID = value; }
        }
        public string Item_Name
        {
            get { return _Item_Name; }
            set { _Item_Name = value; }
        }
        public string Condition
        {
            get { return _Condition; }
            set { _Condition = value; }
        }
        public string Warning_Content
        {
            get { return _Warning_Content; }
            set { _Warning_Content = value; }
        }
        public string System_Show
        {
            get { return _System_Show; }
            set { _System_Show = value; }
        }
        public string Mail_Show
        {
            get { return _Mail_Show; }
            set { _Mail_Show = value; }
        }
        public string RST
        {
            get { return _RST; }
            set { _RST = value; }
        }
        public DateTime RUT
        {
            get { return _RUT; }
            set { _RUT = value; }
        }
        public DateTime RCT
        {
            get { return _RCT; }
            set { _RCT = value; }
        }
        public string RUU
        {
            get { return _RUU; }
            set { _RUU = value; }
        }
        public string RCU
        {
            get { return _RCU; }
            set { _RCU = value; }
        }
        #endregion
    }
}
