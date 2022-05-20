using System;
using System.Data;
using System.Configuration;



/// <summary>
/// IMPORT_HISTORY
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class IMPORT_HISTORY
    {
        public IMPORT_HISTORY()
        { }

        #region Model
        private int _RID;
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;
        private DateTime _Import_Date = Convert.ToDateTime("1900-01-01");
        private string _File_Type;
        private string _File_Name;

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
        public DateTime Import_Date
        {
            get { return _Import_Date; }
            set { _Import_Date = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string File_Type
        {
            get { return _File_Type; }
            set { _File_Type = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string File_Name
        {
            get { return _File_Name; }
            set { _File_Name = value; }
        }

        #endregion


    }
}

