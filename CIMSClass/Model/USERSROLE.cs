using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Configuration;

namespace CIMSClass.Model
{
    public class USERSROLE
    {
        public USERSROLE()
        { }

        #region Model
        private string _UserID;
        private string _RoleID;

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
        public string RoleID
        {
            get { return _RoleID; }
            set { _RoleID = value; }
        }

        #endregion


    }
}
