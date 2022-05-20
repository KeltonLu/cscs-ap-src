using System;

/// <summary>
/// ACTION_HISTORY
/// Date Created: 2008年11月24日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class ACTION_HISTORY
    {
        public ACTION_HISTORY()
        { }

        #region Model
        private int _RID;
        private string _FunctionID;
        private string _UserID;
        private DateTime _Operate_Date = Convert.ToDateTime("1900-01-01");
        private string _Param_Code;

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
        public string FunctionID
        {
            get { return _FunctionID; }
            set { _FunctionID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string UserID
        {
            get { return _UserID; }
            set { _UserID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime Operate_Date
        {
            get { return _Operate_Date; }
            set { _Operate_Date = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Param_Code
        {
            get { return _Param_Code; }
            set { _Param_Code = value; }
        }

        #endregion


    }
}

