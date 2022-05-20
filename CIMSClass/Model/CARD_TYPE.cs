using System;
using System.Data;
using System.Configuration;
//using System.Web;
//using System.Web.Security;
//using System.Web.UI;
//using System.Web.UI.WebControls;
//using System.Web.UI.WebControls.WebParts;
//using System.Web.UI.HtmlControls;


/// <summary>
/// CARD_TYPE
/// Date Created: 2008年11月13日


/// Created By:  aigao
/// </summary>
namespace CIMSClass.Model
{
    public class CARD_TYPE
    {
        public CARD_TYPE()
        { }

        #region Model
        private int _RID;
        private string _Name;
        private int _CardType_Group_RID;
        private int _Finance_Group_RID;
        private string _TOTAL_TYPE;
        private string _TYPE;
        private string _AFFINITY;
        private string _PHOTO;
        private int _BIN;
        private string _Card_Whole_Name;
        private int _Change_Space_RID;
        private int _Replace_Space_RID;
        private DateTime _End_Time = Convert.ToDateTime("1900-01-01");
        private int _Exponent_RID;
        private int _Envelope_RID;
        private string _RCU;
        private string _RUU;
        private DateTime _RCT = Convert.ToDateTime("1900-01-01");
        private DateTime _RUT = Convert.ToDateTime("1900-01-01");
        private string _RST;
        private string _Is_Using;
        private string _Print_Cover;
        private string _Print_Back;
        private string _Comment;
        private DateTime _Begin_Time = Convert.ToDateTime("1900-01-01");
        private string _Display_Name;
        //private string _CID;
        private string _FullName;
        //private string _File_Type;

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
        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int CardType_Group_RID
        {
            get { return _CardType_Group_RID; }
            set { _CardType_Group_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Finance_Group_RID
        {
            get { return _Finance_Group_RID; }
            set { _Finance_Group_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string TOTAL_TYPE
        {
            get { return _TOTAL_TYPE; }
            set { _TOTAL_TYPE = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string TYPE
        {
            get { return _TYPE; }
            set { _TYPE = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string AFFINITY
        {
            get { return _AFFINITY; }
            set { _AFFINITY = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string PHOTO
        {
            get { return _PHOTO; }
            set { _PHOTO = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int BIN
        {
            get { return _BIN; }
            set { _BIN = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Card_Whole_Name
        {
            get { return _Card_Whole_Name; }
            set { _Card_Whole_Name = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Change_Space_RID
        {
            get { return _Change_Space_RID; }
            set { _Change_Space_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Replace_Space_RID
        {
            get { return _Replace_Space_RID; }
            set { _Replace_Space_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime End_Time
        {
            get { return _End_Time; }
            set { _End_Time = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Exponent_RID
        {
            get { return _Exponent_RID; }
            set { _Exponent_RID = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Envelope_RID
        {
            get { return _Envelope_RID; }
            set { _Envelope_RID = value; }
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
        public string Is_Using
        {
            get { return _Is_Using; }
            set { _Is_Using = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Print_Cover
        {
            get { return _Print_Cover; }
            set { _Print_Cover = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Print_Back
        {
            get { return _Print_Back; }
            set { _Print_Back = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Comment
        {
            get { return _Comment; }
            set { _Comment = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public DateTime Begin_Time
        {
            get { return _Begin_Time; }
            set { _Begin_Time = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Display_Name
        {
            get { return _Display_Name; }
            set { _Display_Name = value; }
        }


        /// <summary>
        /// 
        /// </summary>
        //public int CID
        //{
        //    get { return _CID; }
        //    set { _CID = value; }
        //}

        /// <summary>
        /// 
        /// </summary>
        public string FullName
        {
            get { return _FullName; }
            set { _FullName = value; }
        }

        /// <summary>
        /// 
        /// </summary>
        //public int File_Type
        //{
        //    get { return _File_Type; }
        //    set { _File_Type = value; }
        //}

        #endregion


    }
}

