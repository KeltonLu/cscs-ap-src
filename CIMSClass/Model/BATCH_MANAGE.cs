using System;

namespace CIMSService.Model
{
    internal class BATCH_MANAGE
    {
        private int _RID;

        private string _Type;

        private string _Status;

        public int RID
        {
            get
            {
                return this._RID;
            }
            set
            {
                this._RID = value;
            }
        }

        public string Type
        {
            get
            {
                return this._Type;
            }
            set
            {
                this._Type = value;
            }
        }

        public string Status
        {
            get
            {
                return this._Status;
            }
            set
            {
                this._Status = value;
            }
        }
    }
}
