using System;
using System.Collections.Generic;
using System.Text;

namespace CIMSService.Model
{
    class BATCH_MANAGE
    {
        #region Model
        private int _RID;
        private string _Type;
        private string _Status;

        public int RID
        {
            get{return _RID;}
            set { _RID = value; }
        }
        public string Type
        {
            get { return _Type; }
            set { _Type = value; }
        }
        public string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }

        #endregion

    }
}
