using System;
using System.Collections.Generic;
using System.Text;

namespace CIMSClass.Model
{
    internal class WARNING_INFO
    {
        private int _RID;
        private int _Warning_RID;
        private string _UserID;
        private string _Is_Show;
        private string _Warning_Content;

        public int RID
        {
            get { return _RID; }
            set { _RID = value; }
        }

        public int Warning_RID
        {
            get { return _Warning_RID; }
            set { _Warning_RID = value; }
        }
        public string UserID
        {
            get { return _UserID; }
            set { _UserID = value; }
        }
        public string Is_Show
        {
            get { return _Is_Show; }
            set { _Is_Show = value; }
        }
        public string Warning_Content
        {
            get { return _Warning_Content; }
            set { _Warning_Content = value; }
        }
    }
}
