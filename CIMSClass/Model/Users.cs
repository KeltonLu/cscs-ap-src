using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// USERS
/// Date Created: 2008年11月19日
/// Created By:  FangBao
/// </summary>
namespace CIMSClass.Model
{
    public class USERS
    {
        public USERS()
        { }

        #region Model
        private string _UserID;
        private string _UserName;
        private string _Email;
        private string _ProcessRole;
        private DateTime _LastLoginDateTime = Convert.ToDateTime("1900-01-01");

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
        public string UserName
        {
            get { return _UserName; }
            set { _UserName = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Email
        {
            get { return _Email; }
            set { _Email = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string ProcessRole
        {
            get { return _ProcessRole; }
            set { _ProcessRole = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime LastLoginDateTime
        {
            get { return _LastLoginDateTime; }
            set { _LastLoginDateTime = value; }
        }

        #endregion


    }
}

