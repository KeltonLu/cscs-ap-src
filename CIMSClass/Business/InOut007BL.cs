//*****************************************
//*  作    者：YangKun
//*  功能說明：替換前廠商庫存異動匯入
//*  創建日期：2009-07-14
//*  修改日期：
//*  修改記錄：
//*****************************************

//**************************
using CIMSClass.FTP;
using CIMSClass.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace CIMSClass.Business
{
    public class InOut007BL : BaseLogic
    {
        private InOut000BL BL000 = new InOut000BL();

        public const string SEL_FACTORY = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' order by RID";

        public const string SEL_BATCH_MANAGE = "SELECT COUNT(*) FROM BATCH_MANAGE WHERE (RID = 1 OR RID = 4 OR RID = 5) AND Status = 'Y'";

        public string UPDATE_BATCH_MANAGE_START = "UPDATE BATCH_MANAGE SET Status = 'Y',RUU='InOut007BL.cs',RUT='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' WHERE (RID = 1 OR RID = 4 OR RID = 5)";

        public const string SEL_CARDTYPE_PERSO_FACTORY = "SELECT * FROM FACTORY WHERE RST = 'A' AND Is_Perso='Y' AND Factory_ShortName_EN = @Factory_ShortName_EN";

        public const string CON_CARDTYPE_SURPLUS = "SELECT COUNT(*) FROM CARDTYPE_STOCKS WHERE RST='A' AND Stock_Date>=@Stock_Date";

        public const string CON_IMPORT_CARDTYPE_CHANGE_CHECK = "SELECT FCI.* FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN FACTORY F ON FCI.Perso_Factory_RID = F.RID AND F.RST = 'A' WHERE FCI.RST='A' AND F.Factory_ID = @Factory_ID AND (CONVERT(char(10), FCI.Date_Time, 111) = @Date_Time)";

        public const string CON_CARDTYPE_STATUS = "SELECT COUNT(*) FROM CARDTYPE_STATUS WHERE RST='A'";

        public const string SEL_CARD_TYPE = "SELECT * FROM CARD_TYPE WHERE RST='A' AND TYPE=@TYPE AND AFFINITY=@AFFINITY AND PHOTO=@PHOTO";

        public const string SEL_FACTORY_CHANGE_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID LEFT JOIN USERS U ON  WU.UserID = U.UserID WHERE WC.RST = 'A' AND WC.RID =";

        public const string SEL_CHECK_DATE = "SELECT RID FROM CARDTYPE_STOCKS WHERE RST='A' AND Stock_Date = @check_date";

        public const string SEL_FACTORY_CHANGE_IMPORT_ALL = "SELECT FCI.Space_Short_Name,CS.Status_Code,FCI.Perso_Factory_RID,FCI.Date_Time FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @FactoryRID AND  FCI.Date_Time >= @Import_Date_Start AND FCI.Date_Time <= @Import_Date_End";

        public const string SEL_CARDTYPE_STATUS = "SELECT RID,Status_Code,Status_Name FROM CARDTYPE_STATUS WHERE RST='A' ";

        public const string SEL_CARDTYPE_End_Time = "SELECT TYPE,AFFINITY,PHOTO,Name,End_Time,Is_Using FROM CARD_TYPE WHERE RST='A'";

        public const string SEL_FACTORY_RID = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND Factory_ID = @factory_id order by RID";

        public const string SEL_FACTORY_Name = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND RID = @RID ";

        public const string SEL_MADE_CARD_WARNNING = "SELECT FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID,SUM(FCI.Number) AS Number FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID  INNER JOIN CARD_TYPE CT on CT.TYPE=FCI.TYPE and CT.AFFINITY=FCI.AFFINITY and CT.PHOTO=FCI.PHOTO  WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time GROUP BY FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID";

        public const string SEL_EXPRESSIONS_DEFINE_WARNNING = "SELECT ED.Operate,CS.RID FROM EXPRESSIONS_DEFINE ED INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND ED.Type_RID = CS.RID WHERE ED.RST = 'A' AND ED.Expressions_RID = 1";

        public const string DEL_TEMP_MADE_CARD = "DELETE FROM TEMP_MADE_CARD WHERE Perso_Factory_RID = @perso_factory_rid";

        public const string INSERT_INTO_TEMP_MADE_CARD = "INSERT INTO TEMP_MADE_CARD(Perso_Factory_RID,CardType_RID,Number)values(@Perso_Factory_RID,@CardType_RID,@Number)";

        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD = " SELECT EI.Serial_Number AS EI_Number,CE.Serial_Number as CE_Number,TMC.Perso_Factory_RID,TMC.Number FROM TEMP_MADE_CARD TMC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND TMC.CardType_RID = CT.RID LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE TMC.Perso_Factory_RID = @perso_factory_rid";

        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD_DM = " SELECT DI.Serial_Number DI_Number,A.Perso_Factory_RID,A.Number FROM TEMP_MADE_CARD A LEFT JOIN DM_CARDTYPE DCT ON DCT.RST = 'A' AND A.CardType_RID = DCT.CardType_RID LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND DCT.DM_RID = DI.RID WHERE A.Perso_Factory_RID = @perso_factory_rid";

        public const string SEL_LAST_WORK_DATE = "SELECT TOP 1 Date_Time FROM WORK_DATE WHERE Date_Time < @date_time AND Is_WorkDay='Y' ORDER BY Date_Time DESC";

        public const string SEL_MATERIEL_STOCKS_MANAGER = "SELECT Top 1 MSM.Stock_Date,MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number,CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number WHERE Type = '4' AND MSM.Perso_Factory_RID = @perso_factory_rid AND MSM.Serial_Number = @serial_number ORDER BY Stock_Date Desc";

        public const string SEL_MATERIEL_USED = "SELECT SUM(Number) as Number FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Serial_Number = @serial_number  AND Stock_Date>@from_stock_date AND Stock_Date<=@end_stock_date ";

        public const string SEL_LAST_SURPLUS_DAY = "SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS WHERE RST = 'A' ORDER BY  Stock_Date DESC";

        public const string SEL_ENVELOPE_INFO = "SELECT * FROM ENVELOPE_INFO WHERE RST = 'A' AND Serial_Number = @serial_number";

        public const string SEL_CARD_EXPONENT = "SELECT * FROM CARD_EXPONENT WHERE RST = 'A' AND Serial_Number = @serial_number";

        public const string SEL_DMTYPE_INFO = "SELECT * FROM DMTYPE_INFO WHERE RST = 'A' AND Serial_Number = @serial_number";

        public const string SEL_MATERIEL_STOCKS_USED = "select * from MATERIEL_STOCKS_USED where rst='A' AND Serial_Number=@Serial_Number AND Perso_Factory_RID=@Perso_Factory_RID AND Stock_Date > @lastSurplusDateTime AND Stock_Date <= @thisSurplusDateTime";

        public const string SEL_CARDTYPE_ALL = "SELECT * FROM CARD_TYPE WHERE RST='A'";

        public string UPDATE_BATCH_MANAGE_END = "UPDATE BATCH_MANAGE SET Status = 'N',RUU='InOut007BL.cs',RUT='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' WHERE (RID = 1 OR RID = 5 OR RID = 4)";

        public const string DEL_CHANGE_REPLACE_IMPORT = "DELETE FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE RST='A' AND Is_Check='N' AND Perso_Factory_RID = @perso_factory_rid AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";

        public const string DEL_CHANGE_IMPORT = "DELETE FROM FACTORY_CHANGE_IMPORT WHERE RST='A' AND Is_Check='N' AND Perso_Factory_RID = @perso_factory_rid AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";

        public const string SEL_FACTORY_CHANGE = " SELECT FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,FCI.Status_RID,CS.Status_Name,SUM(Number) as Number FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN  FACTORY FC ON FC.RID=FCI.PerSo_Factory_RID   LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID  WHERE FCI.RST = 'A' AND Date_Time >= @date_time_start AND Date_Time<= @date_time_end   GROUP BY FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,FCI.Status_RID,CS.Status_Name ";

        public const string SEL_FACTORY_CHANGE_REPLACE = " SELECT FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.Change_Space_RID,CT.Replace_SPace_RID,FCI.Status_RID,CS.Status_Name,SUM(Number) as Number FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN  FACTORY FC ON FC.RID=FCI.PerSo_Factory_RID   LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID  WHERE FCI.RST = 'A' AND Date_Time >= @date_time_start AND Date_Time<= @date_time_end   GROUP BY FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.Change_Space_RID,CT.Replace_SPace_RID,FCI.Status_RID,CS.Status_Name ";

        public const string SEL_CHANGE_IMPORT_CARDGROUP = "SELECT FCI.RID,FCI.Perso_Factory_RID,FCI.TYPE,FCI.AFFINITY,FCI.PHOTO,FCI.Date_Time,CT.RID as Space_Short_RID,FCI.Status_RID,FCI.Number,P.Param_Name,CG.Group_Name,CG.RID AS CGRID,Space_Short_Name,P.Param_Code,FCI.Is_Auto_Import,CS.Status_Name FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO INNER JOIN GROUP_CARD_TYPE GCT ON CT.RID = GCT.CardType_RID INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND GCT.Group_RID = CG.RID INNER JOIN PARAM P ON P.RST = 'A' AND CG.Param_Code = P.Param_Code AND P.Param_Code = 'use2' INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @perso_factory_rid AND FCI.Date_Time >= @date_time_start AND FCI.Date_Time <= @date_time_end ";

        public const string DEL_FACTORY_CHANGE_REPLACE_IMPORT = "DELETE FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE RID = @RID AND RST='A' AND Is_Check='N' ";

        public const string SEL_FACTORY_CHANGE_IMPORT_RID = "SELECT *  FROM FACTORY_CHANGE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and  Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND RST='A' AND Is_Check='N' AND  Status_RID=@Status_RID AND TYPE = @TYPE AND AFFINITY = @AFFINITY AND PHOTO = @PHOTO ";

        public const string SEL_FACTORY_CHANGE_REPLACE_IMPORT = "SELECT *  FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND ( Status_RID=@OldStatus_RID OR Status_RID=@Status_RID) ";

        public const string DEL_FACTORY_CHANGE_IMPORT = "DELETE FROM FACTORY_CHANGE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and  Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND RST='A' AND Is_Check='N' AND ( Status_RID=@OldStatus_RID OR Status_RID=@Status_RID) ";

        public const string DEL_FACTORY_CHANGE_IMPORT_CARD = "DELETE FROM FACTORY_CHANGE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and  Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND RST='A' AND Is_Check='N' AND  Status_RID=@Status_RID AND TYPE = @TYPE AND AFFINITY = @AFFINITY AND PHOTO = @PHOTO";

        public const string CON_CHECK_DATE = "SELECT count(*) from cardtype_stocks where rst = 'A' and Stock_Date = @CheckDate";

        public DataTable dtPeso;

        public string strErr;

        private Dictionary<string, object> dirValues = new Dictionary<string, object>();

        public InOut007BL()
        {
            string arg_79_0 = ConfigurationManager.AppSettings["FTPCardModifyReplace"];
            string arg_89_0 = ConfigurationManager.AppSettings["FTPCardModifyReplacePath"];
        }

        public ArrayList DownLoadModify(string FactoryPath)
        {
            DateTime.Now.ToString("yyyy");
            DateTime.Now.ToString("MMdd");
            string remotePath = ConfigurationManager.AppSettings["FTPCardModifyReplace"] + "/" + FactoryPath;
            string str = ConfigurationManager.AppSettings["FTPCardModifyReplacePath"];
            FTPFactory fTPFactory = new FTPFactory("CardModifyReplace");
            ArrayList arrayList = new ArrayList();
            string[] fileList = fTPFactory.GetFileList(remotePath);
            if (fileList != null)
            {
                string[] array = fileList;
                for (int i = 0; i < array.Length; i++)
                {
                    string text = array[i];
                    if (this.CheckFile(text))
                    {
                        string[] array2 = text.Split(new char[]
                        {
                            '-'
                        });
                        if (array2 != null)
                        {
                            string text2 = str + "\\" + array2[1] + "\\";
                            if (fTPFactory.Download(remotePath, text, text2, text))
                            {
                                arrayList.Add(new string[]
                                {
                                    text2,
                                    text
                                });
                                fTPFactory.Delete(remotePath, text);
                            }
                        }
                    }
                }
            }
            return arrayList;
        }

        public DataSet GetMateriel(string Serial_Number)
        {
            DataSet dataSet = null;
            DataSet result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("serial_number", Serial_Number);
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                {
                    dataSet = base.dao.GetList("SELECT * FROM ENVELOPE_INFO WHERE RST = 'A' AND Serial_Number = @serial_number", this.dirValues);
                }
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                {
                    dataSet = base.dao.GetList("SELECT * FROM CARD_EXPONENT WHERE RST = 'A' AND Serial_Number = @serial_number", this.dirValues);
                }
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                {
                    dataSet = base.dao.GetList("SELECT * FROM DMTYPE_INFO WHERE RST = 'A' AND Serial_Number = @serial_number", this.dirValues);
                }
                result = dataSet;
            }
            catch (Exception arg_B6_0)
            {
                throw arg_B6_0;
            }
            return result;
        }

        public DateTime getLastSurplusDate()
        {
            DateTime dateTime = Convert.ToDateTime("1900-01-01");
            DateTime result;
            try
            {
                DataSet list = base.dao.GetList("SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS WHERE RST = 'A' ORDER BY  Stock_Date DESC");
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dateTime = Convert.ToDateTime(list.Tables[0].Rows[0]["Stock_Date"].ToString());
                }
                result = dateTime;
            }
            catch (Exception arg_76_0)
            {
                throw arg_76_0;
            }
            return result;
        }

        public DataSet ImportCheck(string strPath, string Date, string filename)
        {
            StreamReader streamReader = null;
            DataSet dataSet = new DataSet();
            new FACTORY_CHANGE_REPLACE_IMPORT();
            ArrayList arrayList = new ArrayList();
            DataTable dataTable = this.CreatTable();
            DataSet dataSet2 = this.CheckFileStatus();
            DataSet dataSet3 = this.CheckCARD_TYPE_EndTime();
            string[] array = filename.Split(new char[]
            {
                '.'
            });
            DataTable perso = this.GetPerso(array[0].Trim());
            DataSet dataSet4 = this.CheckFACTORY_CHANGE_IMPORT(Date, perso.Rows[0]["RID"].ToString().Trim());
            streamReader = new StreamReader(strPath, Encoding.Default);
            int num = 1;
            this.strErr = "";
            try
            {
                string text;
                while ((text = streamReader.ReadLine()) != null)
                {
                    if (num == 1)
                    {
                        string[] array2 = new string[2];
                        if (text.Length != 13)
                        {
                            throw new Exception("Error");
                        }
                        array2[0] = text.Substring(0, 8);
                        array2[1] = text.Substring(8, 5);
                        if (array2[1].Trim() != perso.Rows[0]["Factory_ID"].ToString().Trim())
                        {
                            throw new Exception("Error");
                        }
                        for (int i = 0; i < array2.Length; i++)
                        {
                            int num2 = i + 1;
                            if (StringUtil.IsEmpty(array2[i]))
                            {
                                if (num == 1 && num2 == 1)
                                {
                                    throw new Exception("Error");
                                }
                                if (num == 1 && num2 == 2)
                                {
                                    throw new Exception("Error");
                                }
                            }
                            else
                            {
                                this.strErr += this.CheckFileOneColumn(array2[i], num2, num);
                            }
                        }
                        if (!StringUtil.IsEmpty(this.strErr))
                        {
                            throw new Exception("Error");
                        }
                        num++;
                    }
                    else
                    {
                        if (!StringUtil.IsEmpty(text))
                        {
                            if (StringUtil.GetByteLength(text) != 50)
                            {
                                throw new Exception("Error");
                            }
                            string[] array3 = new string[3];
                            int begin = 0;
                            array3[0] = StringUtil.GetSubstringByByte(text, begin, 9, out begin).Trim();
                            array3[1] = StringUtil.GetSubstringByByte(text, begin, 30, out begin).Trim();
                            array3[2] = StringUtil.GetSubstringByByte(text, begin, 11, out begin).Trim();
                            string[] array2 = new string[]
                            {
                                array3[0].Substring(0, 3),
                                array3[0].Substring(3, 4),
                                array3[0].Substring(7, 2),
                                array3[1],
                                array3[2].Substring(0, 2),
                                array3[2].Substring(2, 9)
                            };
                            for (int j = 0; j < array2.Length; j++)
                            {
                                int num3 = j + 1;
                                if (StringUtil.IsEmpty(array2[j]))
                                {
                                    throw new Exception("Error");
                                }
                                if (this.CheckFileColumn(array2[j], num3, num).Trim().Length > 0)
                                {
                                    throw new Exception("Error");
                                }
                            }
                            DataRow dataRow = dataTable.NewRow();
                            DataRow dataRow2 = dataTable.NewRow();
                            for (int k = 0; k < array2.Length; k++)
                            {
                                if (StringUtil.IsEmpty(array2[k]))
                                {
                                    throw new Exception("Error");
                                }
                            }
                            for (int l = 0; l < array2.Length; l++)
                            {
                                if (l == 0)
                                {
                                    DataTable expr_330 = this.getCARD_TYPE_ALL().Tables[0];
                                    if (expr_330.Rows.Count == 0)
                                    {
                                        throw new Exception("Error");
                                    }
                                    if (expr_330.Select(string.Concat(new string[]
                                    {
                                        "TYPE = '",
                                        array2[0],
                                        "' AND AFFINITY = '",
                                        array2[1],
                                        "' AND PHOTO = '",
                                        array2[2],
                                        "'"
                                    })).Length == 0)
                                    {
                                        throw new Exception("Error");
                                    }
                                }
                                if (l == 3)
                                {
                                    array2[3] = array2[3].Trim();
                                }
                                if (l == 4 && dataSet2.Tables[0].Rows.Count != 0)
                                {
                                    for (int m = 0; m < dataSet2.Tables[0].Rows.Count; m++)
                                    {
                                        arrayList.Add(dataSet2.Tables[0].Rows[m]["Status_Code"].ToString());
                                    }
                                    if (!arrayList.Contains(array2[4]))
                                    {
                                        throw new Exception("Error");
                                    }
                                    DataRow[] array4 = dataSet2.Tables[0].Select("Status_Code='" + array2[4] + "'");
                                    array2[4] = array4[0]["RID"].ToString();
                                }
                                dataRow[l] = array2[l];
                                dataRow2[l] = array2[l];
                            }
                            if (dataSet3.Tables[0].Rows.Count != 0)
                            {
                                for (int n = 0; n < dataSet3.Tables[0].Rows.Count; n++)
                                {
                                    if (dataSet3.Tables[0].Rows[n]["TYPE"].ToString() == dataRow.ItemArray[0].ToString() && dataSet3.Tables[0].Rows[n]["AFFINITY"].ToString() == dataRow.ItemArray[1].ToString() && dataSet3.Tables[0].Rows[n]["PHOTO"].ToString() == dataRow.ItemArray[2].ToString() && dataSet3.Tables[0].Rows[n]["Name"].ToString() == dataRow.ItemArray[3].ToString())
                                    {
                                        dataSet3.Tables[0].Rows[n]["End_Time"].ToString();
                                        Convert.ToDateTime(dataSet3.Tables[0].Rows[n]["End_Time"].ToString());
                                        string text2 = Convert.ToDateTime(dataSet3.Tables[0].Rows[n]["End_Time"].ToString()).ToString("yyyyMMdd");
                                        DateTime.ParseExact(text2, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                                        if (DateTime.ParseExact(text2, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo) <= DateTime.ParseExact(Date, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo) && text2 != "19000101")
                                        {
                                            throw new Exception("Error");
                                        }
                                        if (dataSet3.Tables[0].Rows[n]["Is_Using"].ToString() == "N")
                                        {
                                            throw new Exception("Error");
                                        }
                                    }
                                }
                            }
                            bool flag = false;
                            if (dataSet4.Tables[0].Rows.Count != 0)
                            {
                                for (int num4 = 0; num4 < dataSet4.Tables[0].Rows.Count; num4++)
                                {
                                    if (dataSet4.Tables[0].Rows[num4]["Space_Short_Name"].ToString() == dataRow.ItemArray[3].ToString() && Convert.ToInt16(dataSet4.Tables[0].Rows[num4]["Status_Code"].ToString()) == Convert.ToInt16(dataRow.ItemArray[4]) && dataSet4.Tables[0].Rows[num4]["Perso_Factory_RID"].ToString() == this.GetFactory_RID(perso.Rows[0]["Factory_ID"].ToString().Trim()))
                                    {
                                        flag = true;
                                    }
                                }
                            }
                            if (flag)
                            {
                                throw new Exception("Error");
                            }
                            if (dataTable.Select(string.Concat(new string[]
                            {
                                "TYPE='",
                                array2[0],
                                "' and AFFINITY='",
                                array2[1],
                                "' and PHOTO='",
                                array2[2],
                                "' and Status_RID='",
                                array2[4],
                                "'"
                            })).Length != 0)
                            {
                                throw new Exception("Error");
                            }
                            dataRow2[array2.Length] = perso.Rows[0]["RID"].ToString();
                            dataTable.Rows.Add(dataRow2);
                        }
                        num++;
                    }
                }
                if (!StringUtil.IsEmpty(this.strErr))
                {
                    Warning.SetWarning("53", new object[]
                    {
                        perso.Rows[0]["Factory_ShortName_EN"].ToString(),
                        this.strErr
                    });
                }
                dataSet.Tables.Add(dataTable);
            }
            catch
            {
                Warning.SetWarning("53", new object[]
                {
                    perso.Rows[0]["Factory_ShortName_EN"].ToString(),
                    this.strErr
                });
            }
            finally
            {
                streamReader.Close();
            }
            return dataSet;
        }

        public string ImportCardTypeChange(DataTable dt, string Date, string SaveType)
        {
            string text = "";
            try
            {
                base.dao.OpenConnection();
                DateTime date_Time;
                if (SaveType == "1")
                {
                    date_Time = DateTime.ParseExact(Date, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                }
                else
                {
                    date_Time = DateTime.ParseExact(Date, "yyyy/MM/dd", DateTimeFormatInfo.InvariantInfo);
                }
                FACTORY_CHANGE_REPLACE_IMPORT fACTORY_CHANGE_REPLACE_IMPORT = new FACTORY_CHANGE_REPLACE_IMPORT();
                dt.Columns.Add(new DataColumn("Mk_No", Type.GetType("System.String")));
                int num = 0;
                foreach (DataRow dataRow in dt.Rows)
                {
                    fACTORY_CHANGE_REPLACE_IMPORT.TYPE = dataRow["TYPE"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.AFFINITY = dataRow["AFFINITY"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.PHOTO = dataRow["PHOTO"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.Space_Short_Name = dataRow["Space_Short_Name"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.Status_RID = Convert.ToInt32(dataRow["Status_RID"]);
                    fACTORY_CHANGE_REPLACE_IMPORT.Number = Convert.ToInt32(dataRow["Number"].ToString().Replace(",", ""));
                    fACTORY_CHANGE_REPLACE_IMPORT.Date_Time = date_Time;
                    fACTORY_CHANGE_REPLACE_IMPORT.Perso_Factory_RID = Convert.ToInt32(dataRow["Perso_Factory_RID"].ToString());
                    fACTORY_CHANGE_REPLACE_IMPORT.Is_Auto_Import = "Y";
                    fACTORY_CHANGE_REPLACE_IMPORT.RCU = this.strUserID;
                    fACTORY_CHANGE_REPLACE_IMPORT.RUU = this.strUserID;
                    dataRow["Mk_No"] = DateTime.Now.AddSeconds((double)num).ToString("HHmmss");
                    fACTORY_CHANGE_REPLACE_IMPORT.Mk_No = dataRow["Mk_No"].ToString();
                    base.dao.Add<FACTORY_CHANGE_REPLACE_IMPORT>(fACTORY_CHANGE_REPLACE_IMPORT, "RID");
                    num++;
                }
                DataTable cardNum = new DataTable();
                cardNum = this.GetReplaceCard(dt);
                DataTable dataTable = new DataTable();
                dataTable = this.GetMadeCardNum(cardNum);
                FACTORY_CHANGE_IMPORT fACTORY_CHANGE_IMPORT = new FACTORY_CHANGE_IMPORT();
                foreach (DataRow dataRow2 in dataTable.Rows)
                {
                    fACTORY_CHANGE_IMPORT.TYPE = dataRow2["TYPE"].ToString();
                    fACTORY_CHANGE_IMPORT.AFFINITY = dataRow2["AFFINITY"].ToString();
                    fACTORY_CHANGE_IMPORT.PHOTO = dataRow2["PHOTO"].ToString();
                    fACTORY_CHANGE_IMPORT.Space_Short_Name = dataRow2["Space_Short_Name"].ToString();
                    fACTORY_CHANGE_IMPORT.Status_RID = Convert.ToInt32(dataRow2["Status_RID"]);
                    fACTORY_CHANGE_IMPORT.Number = Convert.ToInt32(dataRow2["Number"].ToString().Replace(",", ""));
                    fACTORY_CHANGE_IMPORT.Date_Time = date_Time;
                    fACTORY_CHANGE_IMPORT.Perso_Factory_RID = Convert.ToInt32(dataRow2["Perso_Factory_RID"].ToString());
                    fACTORY_CHANGE_IMPORT.Is_Auto_Import = "Y";
                    fACTORY_CHANGE_IMPORT.RCU = this.strUserID;
                    fACTORY_CHANGE_IMPORT.RUU = this.strUserID;
                    fACTORY_CHANGE_IMPORT.Mk_No = dataRow2["Mk_No"].ToString();
                    base.dao.Add<FACTORY_CHANGE_IMPORT>(fACTORY_CHANGE_IMPORT, "RID");
                }
                base.dao.Commit();
                this.SendWarningPerso(dataTable, dt.Rows[0]["Perso_Factory_RID"].ToString());
                this.BL000.Material_Used_Warnning(dt.Rows[0]["Perso_Factory_RID"].ToString(), DateTime.Now, "2");
            }
            catch (Exception ex)
            {
                text = "匯入檔保存失敗";
                base.dao.Rollback();
                LogFactory.Write(ex.ToString(), "ErrLog");
                if (!(SaveType == "1"))
                {
                    throw new Exception(ex.Message);
                }
                Warning.SetWarning("53", new object[]
                {
                    this.GetFactory_NamebyRID(dt.Rows[0]["Perso_Factory_RID"].ToString()),
                    text
                });
            }
            finally
            {
                base.dao.CloseConnection();
            }
            return text;
        }

        private void SendWarningPerso(DataTable dtImport, string sFactoryRid)
        {
            try
            {
                DataTable dataTable = this.getCARD_TYPE_ALL().Tables[0];
                DataTable dataTable2 = this.GetFactoryList().Tables[0];
                DataTable dataTable3 = new DataTable();
                dataTable3.Columns.Add("card");
                dataTable3.Columns.Add("factory");
                DataTable dataTable4 = base.dao.GetList("select CardType_RID from dbo.GROUP_CARD_TYPE a inner join CARD_GROUP b on a.Group_rid=b.rid where b.Group_Name = '虛擬卡'").Tables[0];
                foreach (DataRow dataRow in dtImport.Rows)
                {
                    DataRow[] array = dataTable.Select(string.Concat(new string[]
                    {
                        "TYPE='",
                        dataRow["TYPE"].ToString(),
                        "' and AFFINITY='",
                        dataRow["AFFINITY"].ToString(),
                        "' and PHOTO='",
                        dataRow["PHOTO"].ToString(),
                        "'"
                    }));
                    if (array.Length >= 0)
                    {
                        int cardTypeRid;
                        if (dataRow["Status_RID"].ToString() == "4")
                        {
                            if (array[0]["Change_Space_RID"].ToString() != "0")
                            {
                                cardTypeRid = int.Parse(array[0]["Change_Space_RID"].ToString());
                            }
                            else if (array[0]["Replace_Space_RID"].ToString() != "0")
                            {
                                cardTypeRid = int.Parse(array[0]["Replace_Space_RID"].ToString());
                            }
                            else
                            {
                                cardTypeRid = int.Parse(array[0]["RID"].ToString());
                            }
                        }
                        else if (dataRow["Status_RID"].ToString().ToUpper() == "1" || dataRow["Status_RID"].ToString().ToUpper() == "2" || dataRow["Status_RID"].ToString().ToUpper() == "3")
                        {
                            if (array[0]["Replace_Space_RID"].ToString() != "0")
                            {
                                cardTypeRid = int.Parse(array[0]["Replace_Space_RID"].ToString());
                            }
                            else
                            {
                                cardTypeRid = int.Parse(array[0]["RID"].ToString());
                            }
                        }
                        else
                        {
                            cardTypeRid = int.Parse(array[0]["RID"].ToString());
                        }
                        DataRow[] array2 = dataTable2.Select("RID='" + sFactoryRid + "'");
                        if (dataTable3.Select(string.Concat(new string[]
                        {
                            "card='",
                            cardTypeRid.ToString(),
                            "' and factory='",
                            sFactoryRid.ToString(),
                            "'"
                        })).Length == 0 && (dataTable4.Rows.Count <= 0 || dataTable4.Select("CardType_RID = '" + cardTypeRid.ToString() + "'").Length == 0))
                        {
                            DataRow dataRow2 = dataTable3.NewRow();
                            dataRow2[0] = cardTypeRid.ToString();
                            dataRow2[1] = sFactoryRid.ToString();
                            dataTable3.Rows.Add(dataRow2);
                            if (new CardTypeManager().getCurrentStockPerso(Convert.ToInt32(sFactoryRid), cardTypeRid, DateTime.Now.Date.AddDays(1.0).AddSeconds(-1.0)) < 0)
                            {
                                object[] array3 = new object[2];
                                array3[0] = array2[0]["Factory_Shortname_CN"];
                                DataRow[] array4 = dataTable.Select("RID=" + cardTypeRid.ToString());
                                if (array4.Length != 0)
                                {
                                    array3[1] = array4[0]["NAME"];
                                }
                                else
                                {
                                    array3[1] = "";
                                }
                                Warning.SetWarning("56", array3);
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        public DataTable CreatTable()
        {
            return new DataTable
            {
                Columns =
                {
                    new DataColumn("TYPE", Type.GetType("System.String")),
                    new DataColumn("AFFINITY", Type.GetType("System.String")),
                    new DataColumn("PHOTO", Type.GetType("System.String")),
                    new DataColumn("Space_Short_Name", Type.GetType("System.String")),
                    new DataColumn("Status_RID", Type.GetType("System.String")),
                    new DataColumn("Number", Type.GetType("System.String")),
                    new DataColumn("Perso_Factory_RID", Type.GetType("System.String"))
                }
            };
        }

        private string CheckFileColumn(string strColumn, int num, int count)
        {
            string result = "";
            switch (num)
            {
                case 1:
                    {
                        string pattern = "^\\d{3}$";
                        if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
                        {
                            result = string.Concat(new string[]
                            {
                        "第",
                        count.ToString(),
                        "行第",
                        num.ToString(),
                        "列格式必須為3位數字;\\n"
                            });
                        }
                        break;
                    }
                case 2:
                    {
                        string pattern = "^\\d{4}$";
                        if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
                        {
                            result = string.Concat(new string[]
                            {
                        "第",
                        count.ToString(),
                        "行第",
                        num.ToString(),
                        "列格式必須為4位數字;\\n"
                            });
                        }
                        break;
                    }
                case 3:
                    {
                        string pattern = "^\\d{2}$";
                        if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
                        {
                            result = string.Concat(new string[]
                            {
                        "第",
                        count.ToString(),
                        "行第",
                        num.ToString(),
                        "列格式必須為2位數字;\\n"
                            });
                        }
                        break;
                    }
                case 4:
                    if (strColumn.Length == 0)
                    {
                        result = string.Concat(new string[]
                        {
                        "第",
                        count.ToString(),
                        "行第",
                        num.ToString(),
                        "列格式錯誤,匯入文件的卡種必須為30位;\\n"
                        });
                    }
                    break;
                case 5:
                    {
                        string pattern = "^\\d{2}$";
                        if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
                        {
                            result = string.Concat(new string[]
                            {
                        "第",
                        count.ToString(),
                        "行第",
                        num.ToString(),
                        "列格式必須為2位數字;\\n"
                            });
                        }
                        break;
                    }
                case 6:
                    {
                        string pattern = "^\\d{9}|[-]\\d{8}$";
                        if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
                        {
                            result = string.Concat(new string[]
                            {
                        "第",
                        count.ToString(),
                        "行第",
                        num.ToString(),
                        "列格式必須為9位數字;\\n"
                            });
                        }
                        break;
                    }
            }
            return result;
        }

        private DataSet CheckFileStatus()
        {
            DataSet result = null;
            this.dirValues.Clear();
            try
            {
                result = base.dao.GetList("SELECT RID,Status_Code,Status_Name FROM CARDTYPE_STATUS WHERE RST='A' ");
            }
            catch (Exception arg_20_0)
            {
                LogFactory.Write(arg_20_0.ToString(), "ErrLog");
            }
            return result;
        }

        private DataSet CheckCARD_TYPE_EndTime()
        {
            DataSet result = null;
            this.dirValues.Clear();
            try
            {
                result = base.dao.GetList("SELECT TYPE,AFFINITY,PHOTO,Name,End_Time,Is_Using FROM CARD_TYPE WHERE RST='A'");
            }
            catch (Exception arg_20_0)
            {
                LogFactory.Write(arg_20_0.ToString(), "ErrLog");
            }
            return result;
        }

        private DataSet CheckFACTORY_CHANGE_IMPORT(string Import_Date, string FactoryRID)
        {
            DataSet result = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("FactoryRID", FactoryRID);
                this.dirValues.Add("Import_Date_Start", Import_Date + " 00:00:00");
                this.dirValues.Add("Import_Date_End", Import_Date + " 23:59:59");
                result = base.dao.GetList("SELECT FCI.Space_Short_Name,CS.Status_Code,FCI.Perso_Factory_RID,FCI.Date_Time FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @FactoryRID AND  FCI.Date_Time >= @Import_Date_Start AND FCI.Date_Time <= @Import_Date_End", this.dirValues);
            }
            catch (Exception arg_6D_0)
            {
                LogFactory.Write(arg_6D_0.ToString(), "ErrLog");
            }
            return result;
        }

        private string CheckFileOneColumn(string strColumn, int num, int count)
        {
            string result = "";
            string pattern;
            if (num != 1)
            {
                if (num != 2)
                {
                    return result;
                }
            }
            else
            {
                pattern = "^\\d{8}$";
                if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
                {
                    result = "匯入文件的匯入日期必須為8位數字\\n";
                    return result;
                }
                try
                {
                    Convert.ToDateTime(string.Concat(new string[]
                    {
                        strColumn.Substring(0, 4),
                        "/",
                        strColumn.Substring(4, 2),
                        "/",
                        strColumn.Substring(6, 2)
                    }));
                    return result;
                }
                catch (Exception)
                {
                    result = "匯入文件的匯入日期格式不正確\\n";
                    return result;
                }
            }
            pattern = "^\\d{5}$";
            if (Regex.Matches(strColumn, pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture).Count == 0)
            {
                result = "匯入文件的Perso廠商必須為5位數字\\n";
            }
            return result;
        }

        private bool CheckFile(string FileName)
        {
            string[] array = this.CheckFileNameFormat(FileName);
            if (array == null)
            {
                return false;
            }
            if (array[0].ToString().ToUpper() != "OLDCARD")
            {
                return false;
            }
            string enPerso = array[2].Substring(0, array[2].Length - 4);
            if (!this.CheckEnPersoExist(enPerso))
            {
                return false;
            }
            string text = array[1].ToString();
            return new WorkDate().CheckWorkDate(Convert.ToDateTime(string.Concat(new string[]
            {
                text.Substring(0, 4),
                "/",
                text.Substring(4, 2),
                "/",
                text.Substring(6, 2)
            }))) && !this.CheckImportDate(text) && !this.CheckImportFile(FileName);
        }

        private string[] CheckFileNameFormat(string FileName)
        {
            string[] array = FileName.Split(new char[]
            {
                '-'
            });
            if (array == null || array.Length != 3)
            {
                return null;
            }
            return array;
        }

        private bool CheckEnPersoExist(string EnPerso)
        {
            bool result;
            try
            {
                this.dtPeso = new DataTable();
                this.dirValues.Clear();
                this.dirValues.Add("Factory_ShortName_EN", EnPerso);
                this.dtPeso = base.dao.GetList("SELECT * FROM FACTORY WHERE RST = 'A' AND Is_Perso='Y' AND Factory_ShortName_EN = @Factory_ShortName_EN", this.dirValues).Tables[0];
                if (this.dtPeso.Rows.Count > 0)
                {
                    result = true;
                }
                else
                {
                    result = false;
                }
            }
            catch (Exception arg_69_0)
            {
                LogFactory.Write(arg_69_0.ToString(), "ErrLog");
                result = false;
            }
            return result;
        }

        private bool CheckImportDate(string ImportEDate)
        {
            new DataSet();
            this.dirValues.Clear();
            ImportEDate = string.Concat(new string[]
            {
                ImportEDate.Substring(0, 4),
                "/",
                ImportEDate.Substring(4, 2),
                "/",
                ImportEDate.Substring(6, 2)
            });
            this.dirValues.Add("Stock_Date", ImportEDate);
            return !(base.dao.GetList("SELECT COUNT(*) FROM CARDTYPE_STOCKS WHERE RST='A' AND Stock_Date>=@Stock_Date", this.dirValues).Tables[0].Rows[0][0].ToString() == "0");
        }

        private bool CheckImportFile(string ImportFileName)
        {
            string[] array = ImportFileName.Split(new char[]
            {
                '-'
            });
            new DataSet();
            this.dirValues.Clear();
            this.dirValues.Add("Factory_ID", this.dtPeso.Rows[0]["Factory_ID"].ToString());
            this.dirValues.Add("Date_Time", array[1].ToString());
            return base.dao.GetList("SELECT FCI.* FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN FACTORY F ON FCI.Perso_Factory_RID = F.RID AND F.RST = 'A' WHERE FCI.RST='A' AND F.Factory_ID = @Factory_ID AND (CONVERT(char(10), FCI.Date_Time, 111) = @Date_Time)", this.dirValues).Tables[0].Rows.Count > 0;
        }

        private bool CheckValue(string commondSring, Dictionary<string, object> dirValues)
        {
            bool result;
            try
            {
                if (Convert.ToInt32(base.dao.ExecuteScalar(commondSring, dirValues)) == 0)
                {
                    result = false;
                }
                else
                {
                    result = true;
                }
            }
            catch (Exception)
            {
                LogFactory.Write("讀取資料庫錯誤：" + commondSring, "ErrLog");
                result = false;
            }
            return result;
        }

        private DataTable GetPerso(string EnPerso)
        {
            DataTable result;
            try
            {
                this.dtPeso = new DataTable();
                this.dirValues.Clear();
                this.dirValues.Add("Factory_ShortName_EN", EnPerso);
                this.dtPeso = base.dao.GetList("SELECT * FROM FACTORY WHERE RST = 'A' AND Is_Perso='Y' AND Factory_ShortName_EN = @Factory_ShortName_EN", this.dirValues).Tables[0];
                if (this.dtPeso.Rows.Count > 0)
                {
                    result = this.dtPeso;
                }
                else
                {
                    result = null;
                }
            }
            catch (Exception arg_6E_0)
            {
                LogFactory.Write(arg_6E_0.ToString(), "ErrLog");
                result = null;
            }
            return result;
        }

        public DataSet GetFactoryList()
        {
            DataSet list;
            try
            {
                this.dirValues.Clear();
                list = base.dao.GetList("SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' order by RID", this.dirValues);
            }
            catch (Exception arg_24_0)
            {
                LogFactory.Write(arg_24_0.ToString(), "ErrLog");
                throw new Exception("初始化頁面失敗");
            }
            return list;
        }

        private DataTable GetCardType(string Type, string Affinity, string Photo)
        {
            DataTable result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("TYPE", Type);
                this.dirValues.Add("Affinity", Affinity);
                this.dirValues.Add("Photo", Photo);
                result = base.dao.GetList("SELECT * FROM CARD_TYPE WHERE RST='A' AND TYPE=@TYPE AND AFFINITY=@AFFINITY AND PHOTO=@PHOTO", this.dirValues).Tables[0];
            }
            catch (Exception arg_62_0)
            {
                LogFactory.Write(arg_62_0.ToString(), "ErrLog");
                result = null;
            }
            return result;
        }

        public bool ImportFactoryChangeStart()
        {
            bool result;
            try
            {
                DataSet list = base.dao.GetList("SELECT COUNT(*) FROM BATCH_MANAGE WHERE (RID = 1 OR RID = 4 OR RID = 5) AND Status = 'Y'");
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToInt32(list.Tables[0].Rows[0][0]) > 0)
                    {
                        result = false;
                    }
                    else
                    {
                        base.dao.ExecuteNonQuery(this.UPDATE_BATCH_MANAGE_START);
                        result = true;
                    }
                }
                else
                {
                    result = false;
                }
            }
            catch (Exception arg_7E_0)
            {
                LogFactory.Write(arg_7E_0.ToString(), "ErrLog");
                throw new Exception("初始化頁面失敗");
            }
            return result;
        }

        public string Import(string strPath, DataTable dtblFileImp, string Date, string FactoryRID)
        {
            StreamReader streamReader = null;
            string result;
            try
            {
                DataTable dataTable = this.CheckFileStatus().Tables[0];
                DataSet cARD_TYPE_ALL = this.getCARD_TYPE_ALL();
                DataSet fACTORY_CHANGE_IMPORT_ALL = this.getFACTORY_CHANGE_IMPORT_ALL(FactoryRID, Date);
                string value = "";
                int num = 0;
                streamReader = new StreamReader(strPath, Encoding.Default);
                int num2 = 1;
                StringBuilder stringBuilder = new StringBuilder("");
                StringBuilder stringBuilder2 = new StringBuilder("");
                string text;
                while ((text = streamReader.ReadLine()) != null)
                {
                    if (num2 == 1)
                    {
                        if (text.Length != 13)
                        {
                            stringBuilder2.Append("匯入文件缺少匯入日期或Perso厰商，文件無法匯入！\\n");
                        }
                        else
                        {
                            string[] array = new string[]
                            {
                                text.Substring(0, 8),
                                text.Substring(8, 5)
                            };
                            for (int i = 0; i < array.Length; i++)
                            {
                                if (StringUtil.IsEmpty(array[i]))
                                {
                                    if (i == 0)
                                    {
                                        stringBuilder2.Append("第" + num2.ToString() + "行匯入文件的匯入日期不能為空！\\n");
                                    }
                                    else if (i == 1)
                                    {
                                        stringBuilder2.Append("第" + num2.ToString() + "行匯入文件的Perso厰商代號不能為空！\\n");
                                    }
                                }
                                else
                                {
                                    int num3 = i + 1;
                                    stringBuilder2.Append(this.CheckFileOneColumn(array[i], num3, num2));
                                }
                            }
                            if (array[0] != Date.Replace("/", ""))
                            {
                                stringBuilder2.Append("頁面輸入的匯入日期與匯入文件第1行中的匯入日期不符；\\n");
                            }
                            string factory_RID = this.GetFactory_RID(array[1]);
                            if (factory_RID == "" || factory_RID != FactoryRID)
                            {
                                stringBuilder2.Append("頁面輸入的Perso廠商與匯入文件第1行中的廠商代碼不符；\\n");
                            }
                        }
                    }
                    else
                    {
                        stringBuilder = new StringBuilder("");
                        if (StringUtil.GetByteLength(text) != 50)
                        {
                            stringBuilder.Append("第" + num2.ToString() + "行匯入文件格式不正確;\\n");
                            stringBuilder2.Append(stringBuilder);
                        }
                        else
                        {
                            string[] array = new string[3];
                            int begin = 0;
                            array[0] = StringUtil.GetSubstringByByte(text, begin, 9, out begin).Trim();
                            array[1] = StringUtil.GetSubstringByByte(text, begin, 30, out begin).Trim();
                            array[2] = StringUtil.GetSubstringByByte(text, begin, 11, out begin).Trim();
                            string[] array2 = new string[6];
                            if (array[0].Length != 9 || array[2].Length != 11)
                            {
                                stringBuilder.Append("第" + num2.ToString() + "行匯入文件格式不正確;\\n");
                            }
                            else
                            {
                                array2[0] = array[0].Substring(0, 3);
                                array2[1] = array[0].Substring(3, 4);
                                array2[2] = array[0].Substring(7, 2);
                                array2[3] = array[1];
                                array2[4] = array[2].Substring(0, 2);
                                array2[5] = array[2].Substring(2, 9);
                                for (int j = 0; j < array2.Length; j++)
                                {
                                    int num4 = j + 1;
                                    if (StringUtil.IsEmpty(array2[j]))
                                    {
                                        stringBuilder.Append(string.Concat(new string[]
                                        {
                                            "第",
                                            num2.ToString(),
                                            "行第",
                                            num4.ToString(),
                                            "列為空;\\n"
                                        }));
                                    }
                                    else
                                    {
                                        stringBuilder.Append(this.CheckFileColumn(array2[j], num4, num2));
                                    }
                                }
                                if (stringBuilder.Length == 0)
                                {
                                    bool flag = false;
                                    value = "";
                                    for (int k = 0; k < dataTable.Rows.Count; k++)
                                    {
                                        if (Convert.ToInt32(dataTable.Rows[k]["Status_Code"]) == Convert.ToInt32(array2[4]))
                                        {
                                            num = Convert.ToInt32(dataTable.Rows[k]["RID"].ToString());
                                            value = dataTable.Rows[k]["Status_Name"].ToString();
                                            flag = true;
                                            break;
                                        }
                                    }
                                    if (!flag)
                                    {
                                        stringBuilder.Append("第" + num2.ToString() + "行第5列的狀況代碼不存在;\\n");
                                    }
                                    DataRow[] array3 = cARD_TYPE_ALL.Tables[0].Select(string.Concat(new string[]
                                    {
                                        "Type = ",
                                        array2[0],
                                        " AND Affinity = ",
                                        array2[1],
                                        " AND Photo = ",
                                        array2[2]
                                    }));
                                    if (array3.Length == 0)
                                    {
                                        stringBuilder.Append("第" + num2.ToString() + "行的卡種不存在;\\n");
                                    }
                                    else
                                    {
                                        if (array3[0]["Is_Using"].ToString() != "Y")
                                        {
                                            stringBuilder.Append("第" + num2.ToString() + "行的卡種已經停用;\\n");
                                        }
                                        if (Convert.ToDateTime(array3[0]["Begin_Time"].ToString()) > Convert.ToDateTime(Date) || (Convert.ToDateTime(array3[0]["End_Time"].ToString()) < Convert.ToDateTime(Date) && Convert.ToDateTime(array3[0]["End_Time"].ToString()).ToString("yyyy-MM-dd") != "1900-01-01"))
                                        {
                                            stringBuilder.Append("第" + num2.ToString() + "行的卡種不在有效期內;\\n");
                                        }
                                    }
                                    if (fACTORY_CHANGE_IMPORT_ALL.Tables[0].Select("Space_Short_Name = '" + array2[3] + "' AND Status_Code = " + array2[4]).Length != 0)
                                    {
                                        stringBuilder.Append("第" + num2.ToString() + "行的廠商庫存異動資訊已經存在;\\n");
                                    }
                                    if (flag && dtblFileImp.Select(string.Concat(new string[]
                                    {
                                        "TYPE='",
                                        array2[0],
                                        "' and AFFINITY='",
                                        array2[1],
                                        "' and PHOTO='",
                                        array2[2],
                                        "' and Status_RID='",
                                        num.ToString(),
                                        "'"
                                    })).Length != 0)
                                    {
                                        stringBuilder.Append("第" + num2.ToString() + "行的廠商庫存異動資訊不能重複匯入;\\n");
                                    }
                                }
                            }
                            if (stringBuilder.Length > 0)
                            {
                                stringBuilder2.Append(stringBuilder);
                            }
                            else
                            {
                                DataRow dataRow = dtblFileImp.NewRow();
                                dataRow["TYPE"] = array2[0];
                                dataRow["AFFINITY"] = array2[1];
                                dataRow["PHOTO"] = array2[2];
                                dataRow["Space_Short_Name"] = array2[3];
                                dataRow["Status_RID"] = num;
                                dataRow["Status_Name"] = value;
                                dataRow["Number"] = Convert.ToInt32(array2[5]);
                                dataRow["Perso_Factory_RID"] = Convert.ToInt32(FactoryRID);
                                dtblFileImp.Rows.Add(dataRow);
                            }
                        }
                    }
                    num2++;
                }
                if (stringBuilder2.Length == 0)
                {
                    stringBuilder2.Append(this.ImportCardTypeChange(dtblFileImp, Date, "2"));
                    if (stringBuilder2.Length == 0)
                    {
                        base.dao.OpenConnection();
                        base.SetOprLog("11");
                        base.AddLog("2", strPath.Substring(strPath.LastIndexOf('\\') + 1));
                        base.dao.Commit();
                    }
                }
                result = stringBuilder2.ToString();
            }
            catch (Exception ex)
            {
                base.dao.Rollback();
                ExceptionFactory.CreateCustomSaveException("初始化頁面失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("初始化頁面失敗");
            }
            finally
            {
                if (streamReader != null)
                {
                    streamReader.Close();
                }
                base.dao.CloseConnection();
            }
            return result;
        }

        private DataSet getCARD_TYPE_ALL()
        {
            DataSet result = null;
            this.dirValues.Clear();
            try
            {
                result = base.dao.GetList("SELECT * FROM CARD_TYPE WHERE RST='A'");
            }
            catch (Exception ex)
            {
                ExceptionFactory.CreateCustomSaveException("初始化頁面失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("初始化頁面失敗");
            }
            return result;
        }

        private DataSet getFACTORY_CHANGE_IMPORT_ALL(string FactoryRID, string Import_Date)
        {
            DataSet result = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("FactoryRID", FactoryRID);
                this.dirValues.Add("Import_Date_Start", Import_Date + " 00:00:00");
                this.dirValues.Add("Import_Date_End", Import_Date + " 23:59:59");
                result = base.dao.GetList("SELECT FCI.Space_Short_Name,CS.Status_Code,FCI.Perso_Factory_RID,FCI.Date_Time FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @FactoryRID AND  FCI.Date_Time >= @Import_Date_Start AND FCI.Date_Time <= @Import_Date_End", this.dirValues);
            }
            catch (Exception ex)
            {
                ExceptionFactory.CreateCustomSaveException("初始化頁面失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("初始化頁面失敗");
            }
            return result;
        }

        public string GetFactory_RID(string Factory_ID)
        {
            string result = "";
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("factory_id", Factory_ID);
                DataSet list = base.dao.GetList("SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND Factory_ID = @factory_id order by RID", this.dirValues);
                if (list.Tables[0].Rows.Count != 0)
                {
                    result = list.Tables[0].Rows[0]["RID"].ToString();
                }
            }
            catch (Exception arg_7C_0)
            {
                LogFactory.Write(arg_7C_0.ToString(), "ErrLog");
            }
            return result;
        }

        public string GetFactory_Name(string Factory_ID)
        {
            string result = "";
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("factory_id", Factory_ID);
                DataSet list = base.dao.GetList("SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND Factory_ID = @factory_id order by RID", this.dirValues);
                if (list.Tables[0].Rows.Count != 0)
                {
                    result = list.Tables[0].Rows[0]["Factory_ShortName_CN"].ToString();
                }
            }
            catch (Exception arg_7C_0)
            {
                LogFactory.Write(arg_7C_0.ToString(), "ErrLog");
            }
            return result;
        }

        public string GetFactory_NamebyRID(string Factory_RID)
        {
            string result = "";
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("RID", Factory_RID);
                DataSet list = base.dao.GetList("SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND RID = @RID ", this.dirValues);
                if (list.Tables[0].Rows.Count != 0)
                {
                    result = list.Tables[0].Rows[0]["Factory_ShortName_CN"].ToString();
                }
            }
            catch (Exception arg_7C_0)
            {
                LogFactory.Write(arg_7C_0.ToString(), "ErrLog");
            }
            return result;
        }

        public void ImportFactoryChangeEnd()
        {
            try
            {
                base.dao.ExecuteNonQuery(this.UPDATE_BATCH_MANAGE_END);
            }
            catch (Exception ex)
            {
                ExceptionFactory.CreateCustomSaveException("初始化頁面失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("初始化頁面失敗");
            }
        }

        public int Delete(Dictionary<string, object> Input)
        {
            int result = 0;
            try
            {
                base.dao.OpenConnection();
                this.dirValues.Clear();
                this.dirValues.Add("check_date_start", ((DateTime)Input["date_time"]).ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("check_date_end", ((DateTime)Input["date_time"]).ToString("yyyy/MM/dd 23:59:59"));
                this.dirValues.Add("perso_factory_rid", Input["perso_factory_rid"]);
                int arg_C8_0 = base.dao.ExecuteNonQuery("DELETE FROM FACTORY_CHANGE_IMPORT WHERE RST='A' AND Is_Check='N' AND Perso_Factory_RID = @perso_factory_rid AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ", this.dirValues);
                int num = base.dao.ExecuteNonQuery("DELETE FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE RST='A' AND Is_Check='N' AND Perso_Factory_RID = @perso_factory_rid AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ", this.dirValues);
                base.SetOprLog("13");
                result = arg_C8_0 + num;
                base.dao.Commit();
            }
            catch (AlertException arg_E2_0)
            {
                base.dao.Rollback();
                throw arg_E2_0;
            }
            catch (Exception ex)
            {
                base.dao.Rollback();
                ExceptionFactory.CreateCustomSaveException("刪除失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("刪除失敗");
            }
            finally
            {
                base.dao.CloseConnection();
            }
            return result;
        }

        public void GetCompareFactoryReplace(DateTime Date, ref DataTable dtFaReplaceCompare)
        {
            DataTable cardNum = new DataTable();
            cardNum = this.GetFactory_Change(Date);
            DataTable dataTable = new DataTable();
            dataTable = this.GetFactory_Change_Replace(Date);
            dataTable = this.GetReplaceCard(dataTable);
            DataTable dataTable2 = new DataTable();
            DataTable dataTable3 = new DataTable();
            dataTable2 = this.GetMadeCardNum(cardNum);
            dataTable3 = this.GetMadeCardNum(dataTable);
            DataTable dataTable4 = new DataTable();
            if (dataTable2.Rows.Count > 0 || dataTable3.Rows.Count > 0)
            {
                dataTable4 = this.GetCompareMadeCardNum(dataTable2, dataTable3);
                if (dataTable4.Rows.Count > 0)
                {
                    dtFaReplaceCompare = dataTable4;
                }
            }
        }

        public DataTable GetReplaceCard(DataTable ReplaceCard)
        {
            DataTable dataTable = this.getCARD_TYPE_ALL().Tables[0];
            if (ReplaceCard.Rows.Count > 0)
            {
                foreach (DataRow dataRow in ReplaceCard.Rows)
                {
                    if (Convert.ToInt32(dataRow["Status_RID"]) == 5 || Convert.ToInt32(dataRow["Status_RID"]) == 6 || Convert.ToInt32(dataRow["Status_RID"]) == 7)
                    {
                        DataRow[] array = dataTable.Select(string.Concat(new string[]
                        {
                            "TYPE='",
                            dataRow["TYPE"].ToString(),
                            "' and AFFINITY='",
                            dataRow["AFFINITY"].ToString(),
                            "' and PHOTO='",
                            dataRow["PHOTO"].ToString(),
                            "'"
                        }));
                        if (array.Length >= 0)
                        {
                            if (Convert.ToInt32(array[0]["Change_Space_RID"]) != 0)
                            {
                                DataRow[] array2 = dataTable.Select("RID='" + array[0]["Change_Space_RID"].ToString() + "'");
                                if (array2.Length >= 0)
                                {
                                    dataRow["TYPE"] = array2[0]["TYPE"].ToString();
                                    dataRow["AFFINITY"] = array2[0]["AFFINITY"].ToString();
                                    dataRow["PHOTO"] = array2[0]["PHOTO"].ToString();
                                    dataRow["Space_Short_Name"] = array2[0]["Name"].ToString();
                                }
                            }
                            else if (Convert.ToInt32(array[0]["Replace_Space_RID"]) != 0)
                            {
                                DataRow[] array3 = dataTable.Select("RID='" + array[0]["Replace_Space_RID"].ToString() + "'");
                                if (array3.Length >= 0)
                                {
                                    dataRow["TYPE"] = array3[0]["TYPE"].ToString();
                                    dataRow["AFFINITY"] = array3[0]["AFFINITY"].ToString();
                                    dataRow["PHOTO"] = array3[0]["PHOTO"].ToString();
                                    dataRow["Space_Short_Name"] = array3[0]["Name"].ToString();
                                }
                            }
                        }
                    }
                }
            }
            return ReplaceCard;
        }

        public DataTable GetMadeCardNum(DataTable CardNum)
        {
            DataTable dataTable = new DataTable();
            if (CardNum.Rows.Count > 0)
            {
                dataTable = CardNum.Copy();
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    DataRow dataRow = dataTable.Rows[i];
                    if (dataRow["Status_RID"].ToString().Trim() == "5" || dataRow["Status_RID"].ToString().Trim() == "6" || dataRow["Status_RID"].ToString().Trim() == "7")
                    {
                        for (int j = i + 1; j < CardNum.Rows.Count; j++)
                        {
                            DataRow dataRow2 = CardNum.Rows[j];
                            if (dataRow["Status_RID"].ToString().Trim() == dataRow2["Status_RID"].ToString().Trim() && dataRow["Perso_Factory_RID"].ToString().Trim() == dataRow2["Perso_Factory_RID"].ToString().Trim() && dataRow["TYPE"].ToString().Trim() == dataRow2["TYPE"].ToString().Trim() && dataRow["AFFINITY"].ToString().Trim() == dataRow2["AFFINITY"].ToString().Trim() && dataRow["PHOTO"].ToString().Trim() == dataRow2["PHOTO"].ToString().Trim())
                            {
                                dataRow["Number"] = Convert.ToString(int.Parse(dataRow["Number"].ToString()) + int.Parse(dataRow2["Number"].ToString()));
                                dataRow["Mk_No"] = dataRow["Mk_No"].ToString() + "," + dataRow2["Mk_No"].ToString();
                                if (dataRow["Mk_No"].ToString().Length >= 50)
                                {
                                    dataRow["Mk_No"] = dataRow["Mk_No"].ToString().Substring(0, 49);
                                }
                                CardNum.Rows.RemoveAt(j);
                                dataTable.Rows.RemoveAt(j);
                                j--;
                            }
                        }
                    }
                }
            }
            return dataTable;
        }

        private DataTable getDTFactory()
        {
            return new DataTable
            {
                Columns =
                {
                    new DataColumn("Perso_Factory_RID", Type.GetType("System.Int32")),
                    new DataColumn("Factory_ShortName_CN", Type.GetType("System.String")),
                    new DataColumn("TYPE", Type.GetType("System.String")),
                    new DataColumn("AFFINITY", Type.GetType("System.String")),
                    new DataColumn("PHOTO", Type.GetType("System.String")),
                    new DataColumn("Name", Type.GetType("System.String")),
                    new DataColumn("Status_RID", Type.GetType("System.Int32")),
                    new DataColumn("Status_Name", Type.GetType("System.String")),
                    new DataColumn("Number", Type.GetType("System.Int32")),
                    new DataColumn("ReplaceNumber", Type.GetType("System.Int32"))
                }
            };
        }

        public DataTable GetCompareMadeCardNum(DataTable MadeCardNum, DataTable MadeCardNumReplace)
        {
            DataTable dTFactory = this.getDTFactory();
            for (int i = 0; i < MadeCardNumReplace.Rows.Count; i++)
            {
                if (MadeCardNumReplace.Rows[i]["Number"].ToString().Trim() == "0")
                {
                    MadeCardNumReplace.Rows.RemoveAt(i);
                    i--;
                }
            }
            for (int j = 0; j < MadeCardNum.Rows.Count; j++)
            {
                if (MadeCardNum.Rows[j]["Number"].ToString().Trim() == "0")
                {
                    MadeCardNum.Rows.RemoveAt(j);
                    j--;
                }
            }
            if (MadeCardNum.Rows.Count > 0 && MadeCardNumReplace.Rows.Count > 0)
            {
                for (int k = 0; k < MadeCardNumReplace.Rows.Count; k++)
                {
                    DataRow dataRow = MadeCardNumReplace.Rows[k];
                    for (int l = 0; l < MadeCardNum.Rows.Count; l++)
                    {
                        DataRow dataRow2 = MadeCardNum.Rows[l];
                        if (dataRow["Status_RID"].ToString().Trim() == dataRow2["Status_RID"].ToString().Trim() && dataRow["Perso_Factory_RID"].ToString().Trim() == dataRow2["Perso_Factory_RID"].ToString().Trim() && dataRow["TYPE"].ToString().Trim() == dataRow2["TYPE"].ToString().Trim() && dataRow["AFFINITY"].ToString().Trim() == dataRow2["AFFINITY"].ToString().Trim() && dataRow["PHOTO"].ToString().Trim() == dataRow2["PHOTO"].ToString().Trim())
                        {
                            if (dataRow["Number"].ToString().Trim() != dataRow2["Number"].ToString().Trim())
                            {
                                DataRow dataRow3 = dTFactory.NewRow();
                                dataRow3["Perso_Factory_RID"] = dataRow2["Perso_Factory_RID"];
                                dataRow3["Factory_ShortName_CN"] = dataRow2["Factory_ShortName_CN"];
                                dataRow3["TYPE"] = dataRow2["TYPE"];
                                dataRow3["AFFINITY"] = dataRow2["AFFINITY"];
                                dataRow3["PHOTO"] = dataRow2["PHOTO"];
                                dataRow3["Name"] = dataRow2["Name"];
                                dataRow3["Status_RID"] = dataRow2["Status_RID"];
                                dataRow3["Status_Name"] = dataRow2["Status_Name"];
                                dataRow3["Number"] = dataRow2["Number"];
                                dataRow3["ReplaceNumber"] = dataRow["Number"];
                                dTFactory.Rows.Add(dataRow3);
                            }
                            MadeCardNumReplace.Rows.RemoveAt(k);
                            k--;
                            MadeCardNum.Rows.RemoveAt(l);
                            l--;
                            break;
                        }
                    }
                }
            }
            if (MadeCardNumReplace.Rows.Count > 0)
            {
                foreach (DataRow dataRow4 in MadeCardNumReplace.Rows)
                {
                    DataRow dataRow5 = dTFactory.NewRow();
                    dataRow5["Perso_Factory_RID"] = dataRow4["Perso_Factory_RID"];
                    dataRow5["Factory_ShortName_CN"] = dataRow4["Factory_ShortName_CN"];
                    dataRow5["TYPE"] = dataRow4["TYPE"];
                    dataRow5["AFFINITY"] = dataRow4["AFFINITY"];
                    dataRow5["PHOTO"] = dataRow4["PHOTO"];
                    dataRow5["Name"] = dataRow4["Name"];
                    dataRow5["Status_RID"] = dataRow4["Status_RID"];
                    dataRow5["Status_Name"] = dataRow4["Status_Name"];
                    dataRow5["Number"] = "0";
                    dataRow5["ReplaceNumber"] = dataRow4["Number"];
                    dTFactory.Rows.Add(dataRow5);
                }
            }
            if (MadeCardNum.Rows.Count > 0)
            {
                foreach (DataRow dataRow6 in MadeCardNum.Rows)
                {
                    DataRow dataRow7 = dTFactory.NewRow();
                    dataRow7["Perso_Factory_RID"] = dataRow6["Perso_Factory_RID"];
                    dataRow7["Factory_ShortName_CN"] = dataRow6["Factory_ShortName_CN"];
                    dataRow7["TYPE"] = dataRow6["TYPE"];
                    dataRow7["AFFINITY"] = dataRow6["AFFINITY"];
                    dataRow7["PHOTO"] = dataRow6["PHOTO"];
                    dataRow7["Name"] = dataRow6["Name"];
                    dataRow7["Status_RID"] = dataRow6["Status_RID"];
                    dataRow7["Status_Name"] = dataRow6["Status_Name"];
                    dataRow7["Number"] = dataRow6["Number"];
                    dataRow7["ReplaceNumber"] = "0";
                    dTFactory.Rows.Add(dataRow7);
                }
            }
            return dTFactory;
        }

        public DataTable GetFactory_Change_Replace(DateTime Date)
        {
            DataSet dataSet = new DataSet();
            DataTable result = new DataTable();
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("date_time_start", Date.ToString("yyyy-MM-dd 00:00:00"));
                this.dirValues.Add("date_time_end", Date.ToString("yyyy-MM-dd 23:59:59"));
                dataSet = base.dao.GetList(" SELECT FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.Change_Space_RID,CT.Replace_SPace_RID,FCI.Status_RID,CS.Status_Name,SUM(Number) as Number FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN  FACTORY FC ON FC.RID=FCI.PerSo_Factory_RID   LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID  WHERE FCI.RST = 'A' AND Date_Time >= @date_time_start AND Date_Time<= @date_time_end   GROUP BY FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.Change_Space_RID,CT.Replace_SPace_RID,FCI.Status_RID,CS.Status_Name ", this.dirValues);
                if (dataSet.Tables[0].Rows.Count != 0)
                {
                    result = dataSet.Tables[0];
                }
            }
            catch (Exception arg_8D_0)
            {
                LogFactory.Write(arg_8D_0.ToString(), "ErrLog");
            }
            return result;
        }

        public DataTable GetFactory_Change(DateTime Date)
        {
            DataSet dataSet = new DataSet();
            DataTable result = new DataTable();
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("date_time_start", Date.ToString("yyyy-MM-dd 00:00:00"));
                this.dirValues.Add("date_time_end", Date.ToString("yyyy-MM-dd 23:59:59"));
                dataSet = base.dao.GetList(" SELECT FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,FCI.Status_RID,CS.Status_Name,SUM(Number) as Number FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN  FACTORY FC ON FC.RID=FCI.PerSo_Factory_RID   LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID  WHERE FCI.RST = 'A' AND Date_Time >= @date_time_start AND Date_Time<= @date_time_end   GROUP BY FCI.Perso_Factory_RID,FC.Factory_ShortName_CN,CT.RID,CT.NAME,CT.TYPE,CT.AFFINITY,CT.PHOTO,FCI.Status_RID,CS.Status_Name ", this.dirValues);
                if (dataSet.Tables[0].Rows.Count != 0)
                {
                    result = dataSet.Tables[0];
                }
            }
            catch (Exception arg_8D_0)
            {
                LogFactory.Write(arg_8D_0.ToString(), "ErrLog");
            }
            return result;
        }

        public DataTable List(Dictionary<string, object> searchInput, string firstRowNumber, string lastRowNumber, string sortField, string sortType, out int rowCount)
        {
            int num = 0;
            string sortField2 = (sortField == "null") ? "RID" : sortField;
            StringBuilder stringBuilder = new StringBuilder("SELECT FCI.RID,FCI.Perso_Factory_RID,FCI.TYPE,FCI.AFFINITY,FCI.PHOTO,FCI.Date_Time,CT.RID as Space_Short_RID,FCI.Status_RID,FCI.Number,P.Param_Name,CG.Group_Name,CG.RID AS CGRID,Space_Short_Name,P.Param_Code,FCI.Is_Auto_Import,CS.Status_Name FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO INNER JOIN GROUP_CARD_TYPE GCT ON CT.RID = GCT.CardType_RID INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND GCT.Group_RID = CG.RID INNER JOIN PARAM P ON P.RST = 'A' AND CG.Param_Code = P.Param_Code AND P.Param_Code = 'use2' INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @perso_factory_rid AND FCI.Date_Time >= @date_time_start AND FCI.Date_Time <= @date_time_end ");
            DataTable dataTable = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("date_time_start", Convert.ToDateTime(searchInput["date_time"]).ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("date_time_end", Convert.ToDateTime(searchInput["date_time"]).ToString("yyyy/MM/dd 23:59:59"));
                this.dirValues.Add("perso_factory_rid", searchInput["perso_factory_rid"].ToString().Trim());
                dataTable = base.dao.GetList(stringBuilder.ToString(), this.dirValues, firstRowNumber, lastRowNumber, sortField2, sortType, out num).Tables[0];
                dataTable.Columns.Add("IsSurplused", Type.GetType("System.String"));
                if (this.isCheckDate(Convert.ToDateTime(searchInput["date_time"])))
                {
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        dataTable.Rows[i]["IsSurplused"] = "yes";
                    }
                }
                else
                {
                    for (int j = 0; j < dataTable.Rows.Count; j++)
                    {
                        dataTable.Rows[j]["IsSurplused"] = "no";
                    }
                }
                rowCount = num;
            }
            catch (AlertException arg_184_0)
            {
                throw new Exception(arg_184_0.Message);
            }
            catch (Exception ex)
            {
                ExceptionFactory.CreateCustomSaveException("查詢失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("查詢失敗");
            }
            return dataTable;
        }

        public void DelFactory_Change_Import(int strRID)
        {
            try
            {
                FACTORY_CHANGE_REPLACE_IMPORT model = base.dao.GetModel<FACTORY_CHANGE_REPLACE_IMPORT, int>("RID", strRID);
                base.dao.OpenConnection();
                this.dirValues.Clear();
                this.dirValues.Add("RID", strRID);
                base.dao.ExecuteNonQuery("DELETE FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE RID = @RID AND RST='A' AND Is_Check='N' ", this.dirValues);
                if (model.Status_RID == 5 || model.Status_RID == 6 || model.Status_RID == 7)
                {
                    this.SaveFactoryChange(model.Perso_Factory_RID, model.Date_Time, model.Status_RID, model.Status_RID);
                }
                else
                {
                    this.dirValues.Clear();
                    this.dirValues.Add("Perso_Factory_RID", model.Perso_Factory_RID);
                    this.dirValues.Add("Date_Time_Start", model.Date_Time.ToString("yyyy/MM/dd 00:00:00"));
                    this.dirValues.Add("Date_Time_End", model.Date_Time.ToString("yyyy/MM/dd 23:59:59"));
                    this.dirValues.Add("Status_RID", model.Status_RID);
                    this.dirValues.Add("TYPE", model.TYPE);
                    this.dirValues.Add("AFFINITY", model.AFFINITY);
                    this.dirValues.Add("PHOTO", model.PHOTO);
                    base.dao.ExecuteNonQuery("DELETE FROM FACTORY_CHANGE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and  Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND RST='A' AND Is_Check='N' AND  Status_RID=@Status_RID AND TYPE = @TYPE AND AFFINITY = @AFFINITY AND PHOTO = @PHOTO", this.dirValues);
                }
                base.SetOprLog("4");
                base.dao.Commit();
            }
            catch (Exception ex)
            {
                base.dao.Rollback();
                ExceptionFactory.CreateCustomSaveException("初始化頁面失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("刪除資料失敗！");
            }
            finally
            {
                base.dao.CloseConnection();
            }
        }

        public void SaveFACTORY_CHANGE_IMPORT(Dictionary<string, object> searchInput)
        {
            try
            {
                base.dao.OpenConnection();
                FACTORY_CHANGE_REPLACE_IMPORT fACTORY_CHANGE_REPLACE_IMPORT = new FACTORY_CHANGE_REPLACE_IMPORT();
                fACTORY_CHANGE_REPLACE_IMPORT.Date_Time = Convert.ToDateTime(searchInput["Date_Time"]);
                fACTORY_CHANGE_REPLACE_IMPORT.Perso_Factory_RID = Convert.ToInt32(searchInput["Perso_Factory_RID"]);
                fACTORY_CHANGE_REPLACE_IMPORT.Is_Check = searchInput["Is_Check"].ToString();
                fACTORY_CHANGE_REPLACE_IMPORT.Status_RID = Convert.ToInt32(searchInput["Status_RID"]);
                fACTORY_CHANGE_REPLACE_IMPORT.TYPE = searchInput["TYPE"].ToString();
                fACTORY_CHANGE_REPLACE_IMPORT.AFFINITY = searchInput["AFFINITY"].ToString();
                fACTORY_CHANGE_REPLACE_IMPORT.PHOTO = searchInput["PHOTO"].ToString();
                fACTORY_CHANGE_REPLACE_IMPORT.Number = Convert.ToInt32(searchInput["Number"]);
                fACTORY_CHANGE_REPLACE_IMPORT.Space_Short_Name = searchInput["Space_Short_Name"].ToString();
                fACTORY_CHANGE_REPLACE_IMPORT.Check_Date = Convert.ToDateTime(searchInput["Check_Date"]);
                fACTORY_CHANGE_REPLACE_IMPORT.Is_Auto_Import = searchInput["Is_Auto_Import"].ToString();
                fACTORY_CHANGE_REPLACE_IMPORT.RCU = this.strUserID;
                fACTORY_CHANGE_REPLACE_IMPORT.RUU = this.strUserID;
                fACTORY_CHANGE_REPLACE_IMPORT.Mk_No = DateTime.Now.ToString("HHmmss");
                base.dao.Add<FACTORY_CHANGE_REPLACE_IMPORT>(fACTORY_CHANGE_REPLACE_IMPORT, "RID");
                if (fACTORY_CHANGE_REPLACE_IMPORT.Status_RID == 5 || fACTORY_CHANGE_REPLACE_IMPORT.Status_RID == 6 || fACTORY_CHANGE_REPLACE_IMPORT.Status_RID == 7)
                {
                    this.SaveFactoryChange(fACTORY_CHANGE_REPLACE_IMPORT.Perso_Factory_RID, fACTORY_CHANGE_REPLACE_IMPORT.Date_Time, fACTORY_CHANGE_REPLACE_IMPORT.Status_RID, fACTORY_CHANGE_REPLACE_IMPORT.Status_RID);
                }
                else
                {
                    FACTORY_CHANGE_IMPORT fACTORY_CHANGE_IMPORT = new FACTORY_CHANGE_IMPORT();
                    fACTORY_CHANGE_IMPORT.Date_Time = fACTORY_CHANGE_REPLACE_IMPORT.Date_Time;
                    fACTORY_CHANGE_IMPORT.Perso_Factory_RID = fACTORY_CHANGE_REPLACE_IMPORT.Perso_Factory_RID;
                    fACTORY_CHANGE_IMPORT.Is_Check = fACTORY_CHANGE_REPLACE_IMPORT.Is_Check;
                    fACTORY_CHANGE_IMPORT.Status_RID = fACTORY_CHANGE_REPLACE_IMPORT.Status_RID;
                    fACTORY_CHANGE_IMPORT.TYPE = fACTORY_CHANGE_REPLACE_IMPORT.TYPE;
                    fACTORY_CHANGE_IMPORT.AFFINITY = fACTORY_CHANGE_REPLACE_IMPORT.AFFINITY;
                    fACTORY_CHANGE_IMPORT.PHOTO = fACTORY_CHANGE_REPLACE_IMPORT.PHOTO;
                    fACTORY_CHANGE_IMPORT.Number = fACTORY_CHANGE_REPLACE_IMPORT.Number;
                    fACTORY_CHANGE_IMPORT.Space_Short_Name = fACTORY_CHANGE_REPLACE_IMPORT.Space_Short_Name;
                    fACTORY_CHANGE_IMPORT.Check_Date = fACTORY_CHANGE_REPLACE_IMPORT.Check_Date;
                    fACTORY_CHANGE_IMPORT.Is_Auto_Import = fACTORY_CHANGE_REPLACE_IMPORT.Is_Auto_Import;
                    fACTORY_CHANGE_IMPORT.RCU = this.strUserID;
                    fACTORY_CHANGE_IMPORT.RUU = this.strUserID;
                    fACTORY_CHANGE_IMPORT.Mk_No = fACTORY_CHANGE_REPLACE_IMPORT.Mk_No;
                    base.dao.Add<FACTORY_CHANGE_IMPORT>(fACTORY_CHANGE_IMPORT, "RID");
                }
                base.SetOprLog("2");
                base.dao.Commit();
            }
            catch (Exception ex)
            {
                base.dao.Rollback();
                ExceptionFactory.CreateCustomSaveException("初始化頁面失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("新增資料失敗！");
            }
            finally
            {
                base.dao.CloseConnection();
            }
        }

        public void Update(int rid, Dictionary<string, object> searchInput)
        {
            try
            {
                base.dao.OpenConnection();
                FACTORY_CHANGE_REPLACE_IMPORT model = base.dao.GetModel<FACTORY_CHANGE_REPLACE_IMPORT, int>("RID", rid);
                FACTORY_CHANGE_REPLACE_IMPORT model2 = base.dao.GetModel<FACTORY_CHANGE_REPLACE_IMPORT, int>("RID", rid);
                model.Date_Time = Convert.ToDateTime(searchInput["Date_Time"]);
                model.Perso_Factory_RID = Convert.ToInt32(searchInput["Perso_Factory_RID"]);
                model.Status_RID = Convert.ToInt32(searchInput["Status_RID"]);
                model.TYPE = searchInput["TYPE"].ToString();
                model.AFFINITY = searchInput["AFFINITY"].ToString();
                model.PHOTO = searchInput["PHOTO"].ToString();
                model.Number = Convert.ToInt32(searchInput["Number"]);
                model.Space_Short_Name = searchInput["Space_Short_Name"].ToString();
                model.RUU = this.strUserID;
                base.dao.Update<FACTORY_CHANGE_REPLACE_IMPORT>(model, "RID");
                if (model.Status_RID == 5 || model.Status_RID == 6 || model.Status_RID == 7 || model2.Status_RID == 5 || model2.Status_RID == 6 || model2.Status_RID == 7)
                {
                    this.SaveFactoryChange(model.Perso_Factory_RID, model.Date_Time, model.Status_RID, model2.Status_RID);
                }
                else
                {
                    int factory_Change_Import_RID = this.GetFactory_Change_Import_RID(model2);
                    FACTORY_CHANGE_IMPORT model3 = base.dao.GetModel<FACTORY_CHANGE_IMPORT, int>("RID", factory_Change_Import_RID);
                    model3.Date_Time = model.Date_Time;
                    model3.Perso_Factory_RID = model.Perso_Factory_RID;
                    model3.Status_RID = model.Status_RID;
                    model3.TYPE = model.TYPE;
                    model3.AFFINITY = model.AFFINITY;
                    model3.PHOTO = model.PHOTO;
                    model3.Number = model.Number;
                    model3.Space_Short_Name = model.Space_Short_Name;
                    model3.RUU = this.strUserID;
                    base.dao.Update<FACTORY_CHANGE_IMPORT>(model3, "RID");
                }
                base.SetOprLog("3");
                base.dao.Commit();
            }
            catch (Exception ex)
            {
                base.dao.Rollback();
                ExceptionFactory.CreateCustomSaveException("儲存失敗，請稍後再按一次儲存", ex.Message, base.dao.LastCommands);
                throw new Exception("儲存失敗，請稍後再按一次儲存");
            }
            finally
            {
                base.dao.CloseConnection();
            }
        }

        public void SaveFactoryChange(int Perso_Factory_RID, DateTime Date_Time, int Status_RID, int OldStatus_RID)
        {
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("Perso_Factory_RID", Perso_Factory_RID);
                this.dirValues.Add("Date_Time_Start", Date_Time.ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("Date_Time_End", Date_Time.ToString("yyyy/MM/dd 23:59:59"));
                this.dirValues.Add("Status_RID", Status_RID);
                this.dirValues.Add("OldStatus_RID", OldStatus_RID);
                base.dao.ExecuteNonQuery("DELETE FROM FACTORY_CHANGE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and  Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND RST='A' AND Is_Check='N' AND ( Status_RID=@OldStatus_RID OR Status_RID=@Status_RID) ", this.dirValues);
                DataTable dataTable = base.dao.GetList("SELECT *  FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND ( Status_RID=@OldStatus_RID OR Status_RID=@Status_RID) ", this.dirValues).Tables[0];
                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    DataTable cardNum = new DataTable();
                    cardNum = this.GetReplaceCard(dataTable);
                    new DataTable();
                    DataTable arg_F7_0 = this.GetMadeCardNum(cardNum);
                    FACTORY_CHANGE_IMPORT fACTORY_CHANGE_IMPORT = new FACTORY_CHANGE_IMPORT();
                    foreach (DataRow dataRow in arg_F7_0.Rows)
                    {
                        fACTORY_CHANGE_IMPORT.TYPE = dataRow["TYPE"].ToString();
                        fACTORY_CHANGE_IMPORT.AFFINITY = dataRow["AFFINITY"].ToString();
                        fACTORY_CHANGE_IMPORT.PHOTO = dataRow["PHOTO"].ToString();
                        fACTORY_CHANGE_IMPORT.Space_Short_Name = dataRow["Space_Short_Name"].ToString();
                        fACTORY_CHANGE_IMPORT.Status_RID = Convert.ToInt32(dataRow["Status_RID"]);
                        fACTORY_CHANGE_IMPORT.Number = Convert.ToInt32(dataRow["Number"].ToString().Replace(",", ""));
                        fACTORY_CHANGE_IMPORT.Date_Time = Convert.ToDateTime(dataRow["Date_Time"].ToString());
                        fACTORY_CHANGE_IMPORT.Perso_Factory_RID = Convert.ToInt32(dataRow["Perso_Factory_RID"].ToString());
                        fACTORY_CHANGE_IMPORT.Is_Auto_Import = dataRow["Is_Auto_Import"].ToString();
                        fACTORY_CHANGE_IMPORT.RCU = dataRow["RCU"].ToString();
                        fACTORY_CHANGE_IMPORT.RUU = this.strUserID;
                        fACTORY_CHANGE_IMPORT.Mk_No = dataRow["Mk_No"].ToString();
                        base.dao.Add<FACTORY_CHANGE_IMPORT>(fACTORY_CHANGE_IMPORT, "RID");
                    }
                }
            }
            catch (Exception expr_26F)
            {
                LogFactory.Write(expr_26F.ToString(), "ErrLog");
                throw expr_26F;
            }
        }

        public bool isCheckDate(DateTime CheckDate)
        {
            bool result;
            try
            {
                this.dirValues.Add("CheckDate", CheckDate);
                result = base.dao.Contains("SELECT count(*) from cardtype_stocks where rst = 'A' and Stock_Date = @CheckDate", this.dirValues);
            }
            catch (Exception ex)
            {
                ExceptionFactory.CreateCustomSaveException("查詢失敗", ex.Message, base.dao.LastCommands);
                throw new Exception("查詢失敗");
            }
            return result;
        }

        public int GetFactory_Change_Import_RID(FACTORY_CHANGE_REPLACE_IMPORT OldfciModel)
        {
            int result = 0;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("Perso_Factory_RID", OldfciModel.Perso_Factory_RID);
                this.dirValues.Add("Date_Time_Start", OldfciModel.Date_Time.ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("Date_Time_End", OldfciModel.Date_Time.ToString("yyyy/MM/dd 23:59:59"));
                this.dirValues.Add("Status_RID", OldfciModel.Status_RID);
                this.dirValues.Add("TYPE", OldfciModel.TYPE);
                this.dirValues.Add("AFFINITY", OldfciModel.AFFINITY);
                this.dirValues.Add("PHOTO", OldfciModel.PHOTO);
                DataTable dataTable = base.dao.GetList("SELECT *  FROM FACTORY_CHANGE_IMPORT WHERE Perso_Factory_RID=@Perso_Factory_RID and  Date_Time>=@Date_Time_Start and Date_Time<=@Date_Time_End  AND RST='A' AND Is_Check='N' AND  Status_RID=@Status_RID AND TYPE = @TYPE AND AFFINITY = @AFFINITY AND PHOTO = @PHOTO ", this.dirValues).Tables[0];
                if (dataTable != null && dataTable.Rows.Count > 0)
                {
                    result = Convert.ToInt32(dataTable.Rows[0]["RID"].ToString());
                }
            }
            catch (Exception arg_121_0)
            {
                throw new Exception(arg_121_0.Message);
            }
            return result;
        }

        public void AddReplaceImp()
        {
            string commandText = " select * from FACTORY_CHANGE_IMPORT where Mk_No is  null ";
            string commandText2 = " delete from FACTORY_CHANGE_REPLACE_IMPORT where Mk_No is  null ";
            try
            {
                new DataTable();
                DataTable arg_47_0 = base.dao.GetList(commandText).Tables[0];
                base.dao.OpenConnection();
                base.dao.ExecuteNonQuery(commandText2);
                FACTORY_CHANGE_REPLACE_IMPORT fACTORY_CHANGE_REPLACE_IMPORT = new FACTORY_CHANGE_REPLACE_IMPORT();
                foreach (DataRow dataRow in arg_47_0.Rows)
                {
                    fACTORY_CHANGE_REPLACE_IMPORT.TYPE = dataRow["TYPE"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.AFFINITY = dataRow["AFFINITY"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.PHOTO = dataRow["PHOTO"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.Space_Short_Name = dataRow["Space_Short_Name"].ToString();
                    fACTORY_CHANGE_REPLACE_IMPORT.Status_RID = Convert.ToInt32(dataRow["Status_RID"]);
                    fACTORY_CHANGE_REPLACE_IMPORT.Number = Convert.ToInt32(dataRow["Number"].ToString().Replace(",", ""));
                    fACTORY_CHANGE_REPLACE_IMPORT.Date_Time = DateTime.Parse(dataRow["Date_Time"].ToString());
                    fACTORY_CHANGE_REPLACE_IMPORT.Perso_Factory_RID = Convert.ToInt32(dataRow["Perso_Factory_RID"].ToString());
                    fACTORY_CHANGE_REPLACE_IMPORT.Is_Auto_Import = "Y";
                    fACTORY_CHANGE_REPLACE_IMPORT.RCU = this.strUserID;
                    fACTORY_CHANGE_REPLACE_IMPORT.RUU = this.strUserID;
                    base.dao.Add<FACTORY_CHANGE_REPLACE_IMPORT>(fACTORY_CHANGE_REPLACE_IMPORT, "RID");
                }
                base.dao.Commit();
            }
            catch (Exception ex)
            {
                base.dao.Rollback();
                ExceptionFactory.CreateCustomSaveException("儲存失敗，請稍後再按一次儲存", ex.Message, base.dao.LastCommands);
                throw new Exception("儲存失敗，請稍後再按一次儲存");
            }
            finally
            {
                base.dao.CloseConnection();
            }
        }

    }
}
