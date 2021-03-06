//*****************************************
//*  作    者：YangKun
//*  功能說明：匯入程式公共方法
//*  創建日期：2009-09-14
//*  修改日期：
//*  修改記錄：
//*            □2009-11-26
//*              修改 楊昆
//*                      1.增加日結時計算每日物料庫存結余算法SaveSurplusSystemNum()
//*                      2.物料的消耗數統計算法修改，統計耗用量時不再加上耗損率進行計算
//*****************************************

//**************************
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Text.RegularExpressions;
using System.IO;
using CIMSClass;
using CIMSClass.Model;
using CIMSClass.FTP;
using CIMSClass.Mail;
using CIMSClass.Business;

namespace CIMSClass.Business
{
    public class InOut000BL : BaseLogic
    {
        private WorkDate WorkDateBL = new WorkDate();

        public const string SEL_FACTORY = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' order by RID";

        public const string SEL_BATCH_MANAGE = "SELECT COUNT(*) FROM BATCH_MANAGE WHERE (RID = 1 OR RID = 4 OR RID = 5) AND Status = 'Y'";

        public string UPDATE_BATCH_MANAGE_START = "UPDATE BATCH_MANAGE SET Status = 'Y',RUU='InOut007BL.cs',RUT='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' WHERE (RID = 1 OR RID = 4 OR RID = 5)";

        public const string SEL_CARDTYPE_PERSO_FACTORY = "SELECT * FROM FACTORY WHERE RST = 'A' AND Is_Perso='Y' AND Factory_ShortName_EN = @Factory_ShortName_EN";

        public const string CON_CARDTYPE_SURPLUS = "SELECT COUNT(*) FROM CARDTYPE_STOCKS WHERE RST='A' AND Stock_Date>=@Stock_Date";

        public const string CON_IMPORT_CARDTYPE_CHANGE_CHECK = "SELECT FCI.* FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN FACTORY F ON FCI.Perso_Factory_RID = F.RID AND F.RST = 'A' WHERE FCI.RST='A' AND F.Factory_ID = @Factory_ID AND (CONVERT(char(10), FCI.Date_Time, 111) = @Date_Time)";

        public const string CON_CARDTYPE_STATUS = "SELECT COUNT(*) FROM CARDTYPE_STATUS WHERE RST='A'";

        public const string SEL_CARD_TYPE = "SELECT * FROM CARD_TYPE WHERE RST='A'";

        public const string SEL_FACTORY_CHANGE_ALERT = "SELECT * FROM WARNING_CONFIGURATION WC INNER JOIN WARNING_USER WU ON WC.RID = WU.Warning_RID LEFT JOIN USERS U ON  WU.UserID = U.UserID WHERE WC.RST = 'A' AND WC.RID =";

        public const string SEL_CHECK_DATE = "SELECT RID FROM CARDTYPE_STOCKS WHERE RST='A' AND Stock_Date = @check_date";

        public const string SEL_FACTORY_CHANGE_IMPORT_ALL = "SELECT FCI.Space_Short_Name,CS.Status_Code,FCI.Perso_Factory_RID,FCI.Date_Time FROM FACTORY_CHANGE_REPLACE_IMPORT FCI LEFT JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID WHERE FCI.RST='A' AND FCI.Perso_Factory_RID = @FactoryRID AND  FCI.Date_Time >= @Import_Date_Start AND FCI.Date_Time <= @Import_Date_End";

        public const string SEL_CARDTYPE_STATUS = "SELECT RID,Status_Code,Status_Name FROM CARDTYPE_STATUS WHERE RST='A' ";

        public const string SEL_CARDTYPE_End_Time = "SELECT TYPE,AFFINITY,PHOTO,Name,End_Time,Is_Using FROM CARD_TYPE WHERE RST='A'";

        public const string SEL_FACTORY_RID = "SELECT F.RID,F.Factory_ID,F.Factory_ShortName_CN FROM FACTORY AS F WHERE F.RST = 'A' AND F.Is_Perso = 'Y' AND Factory_ID = @factory_id order by RID";

        public const string SEL_MADE_CARD_WARNNING = "SELECT * FROM (SELECT SI.Perso_Factory_RID,CT.RID,CS.RID AS Status_RID,SUM(SI.Number) AS Number FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name  WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time GROUP BY SI.Perso_Factory_RID,CT.RID,CS.RID UNION ALL SELECT FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID,SUM(FCI.Number) AS Number FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN')  INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time GROUP BY FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID ) A ORDER BY Perso_Factory_RID,RID,Status_RID ";

        public const string SEL_EXPRESSIONS_DEFINE_WARNNING = "SELECT ED.Operate,CS.RID FROM EXPRESSIONS_DEFINE ED INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND ED.Type_RID = CS.RID WHERE ED.RST = 'A' AND ED.Expressions_RID = 1";

        public const string DEL_TEMP_MADE_CARD = "DELETE FROM TEMP_MADE_CARD WHERE Perso_Factory_RID = @perso_factory_rid";

        public const string INSERT_INTO_TEMP_MADE_CARD = "INSERT INTO TEMP_MADE_CARD(Perso_Factory_RID,CardType_RID,Number)values(@Perso_Factory_RID,@CardType_RID,@Number)";

        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD = " SELECT EI.Serial_Number AS EI_Number,CE.Serial_Number as CE_Number,TMC.Perso_Factory_RID,TMC.Number FROM TEMP_MADE_CARD TMC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND TMC.CardType_RID = CT.RID LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE TMC.Perso_Factory_RID = @perso_factory_rid";

        public const string SEL_MATERIAL_BY_TEMP_MADE_CARD_DM = " SELECT DI.Serial_Number AS DI_Number,A.Perso_Factory_RID,A.Number  FROM TEMP_MADE_CARD A  INNER JOIN DM_CARDTYPE DCT ON DCT.RST = 'A' AND A.CardType_RID = DCT.CardType_RID  INNER JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND DCT.DM_RID = DI.RID  WHERE A.Perso_Factory_RID = @perso_factory_rid ";

        public const string SEL_MATERIEL_STOCKS_MANAGER = "SELECT Top 1 MSM.Stock_Date,MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number,CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number WHERE Type = '4' AND MSM.Perso_Factory_RID = @perso_factory_rid AND MSM.Serial_Number = @serial_number ORDER BY Stock_Date Desc";

        public const string SEL_MATERIEL_STOCKS_DAY = "SELECT MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number,CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME FROM MATERIEL_STOCKS MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number WHERE Convert(varchar(10),Stock_Date,111) = @stock_date AND MSM.Serial_Number IN (";

        public const string SEL_MATERIEL_USED = "SELECT isnull(SUM(Number),0) as Number FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Serial_Number = @serial_number  AND Stock_Date>@from_stock_date AND Stock_Date<=@end_stock_date ";

        public const string SEL_MATERIEL_USED_DAY = "SELECT isnull(sum(Number),0) FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Serial_Number = @serial_number AND Convert(varchar(10),Stock_Date,111) = @stock_date ";

        public const string SEL_LAST_SURPLUS_DAY = "SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS WHERE RST = 'A' ORDER BY  Stock_Date DESC";

        public const string SEL_ENVELOPE_INFO = "SELECT * FROM ENVELOPE_INFO WHERE RST = 'A' AND Serial_Number = @serial_number";

        public const string SEL_CARD_EXPONENT = "SELECT * FROM CARD_EXPONENT WHERE RST = 'A' AND Serial_Number = @serial_number";

        public const string SEL_DMTYPE_INFO = "SELECT * FROM DMTYPE_INFO WHERE RST = 'A' AND Serial_Number = @serial_number";

        public const string SEL_MATERIEL_STOCKS_USED = "select isnull(sum(Number),0) as System_Num from MATERIEL_STOCKS_USED where rst='A' AND Serial_Number=@Serial_Number AND Perso_Factory_RID=@Perso_Factory_RID AND Stock_Date > @lastSurplusDateTime AND Stock_Date <= @thisSurplusDateTime";

        public const string SEL_CARDTYPE_ALL = "SELECT * FROM CARD_TYPE WHERE RST='A'";

        public string UPDATE_BATCH_MANAGE_END = "UPDATE BATCH_MANAGE SET Status = 'N',RUU='InOut007BL.cs',RUT='" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "' WHERE (RID = 1 OR RID = 5 OR RID = 4)";

        public const string DEL_CHANGE_IMPORT = "DELETE FROM FACTORY_CHANGE_REPLACE_IMPORT WHERE RST='A' AND Is_Check='N' AND Perso_Factory_RID = @perso_factory_rid AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";

        public const string DEL_MATERIEL_STOCKS = "DELETE  FROM MATERIEL_STOCKS  WHERE RST = 'A' and Perso_Factory_RID=@Perso_Factory_RID and Serial_Number=@Serial_Number  and Stock_Date > @lastSurplusDateTime ";

        public const string SEL_MATERIEL_USED_NUMBER = "SELECT Stock_Date,isnull(sum(Number),0)  as Number FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Stock_Date > @lastSurplusDateTime and Perso_Factory_RID=@Perso_Factory_RID and Serial_Number=@Serial_Number Group By Stock_Date Order By Stock_Date ";

        public const string SEL_MATERIEL_SERIAL_NUMBER = "select Perso_Factory_RID,Serial_Number from ( SELECT MS.Perso_Factory_RID,MS.Serial_Number from MATERIEL_STOCKS MS where MS.RST = 'A' and MS.Stock_Date>=@lastSurplusDateTime union SELECT MSU.Perso_Factory_RID,MSU.Serial_Number from MATERIEL_STOCKS_USED MSU where MSU.RST = 'A' and MSU.Stock_Date>=@lastSurplusDateTime ) A order by Perso_Factory_RID,Serial_Number ";

        public const string SEL_MATERIEL_STOCKS_MANAGE_LASTSURPLUS = "select MSM.Perso_Factory_RID,MSM.Serial_number,MSM.Stock_Date,MSM.Number  from MATERIEL_STOCKS_MANAGE MSM  inner join (  select Perso_Factory_RID,Serial_number,max(Stock_Date) as Stock_Date  from MATERIEL_STOCKS_MANAGE   where type=4  group by Perso_Factory_RID,Serial_number  ) A  on A.Perso_Factory_RID=MSM.Perso_Factory_RID and A.Serial_number=MSM.Serial_number and A.Stock_Date=MSM.Stock_Date and MSM.type = 4  order by MSM.Perso_Factory_RID,MSM.Serial_number  ";

        public const string SEL_MATERIEL_STOCKS_MOVE_IN = "SELECT Move_Date,Move_Number FROM MATERIEL_STOCKS_MOVE WHERE RST = 'A' AND To_Factory_RID = @perso_factory_rid AND Serial_Number = @Serial_Number AND Move_Date > @lastSurplusDateTime AND Move_Date <= @thisSurplusDateTime ";

        public const string SEL_MATERIEL_STOCKS_MOVE_OUT = "SELECT Move_Date,Move_Number FROM MATERIEL_STOCKS_MOVE WHERE RST = 'A' AND From_Factory_RID = @perso_factory_rid AND Serial_Number = @Serial_Number AND Move_Date > @lastSurplusDateTime AND Move_Date <= @thisSurplusDateTime ";

        public const string SEL_MATERIAL_USED_ENVELOPE_REPLACE = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,EI.Serial_Number,CS.RID as Status_RID,SI.Number  FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN')  INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name   INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE SI.RST = 'A' AND SI.Perso_Factory_RID = @Perso_Factory_RID AND SI.Date_Time>@From_Date_Time AND SI.Date_Time<=@End_Date_Time AND EI.Serial_Number = @Serial_Number UNION All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,EI.Serial_Number,FCI.Status_RID  ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time AND EI.Serial_Number = @Serial_Number ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID";

        public const string SEL_MATERIAL_USED_CARD_EXPONENT_REPLACE = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,CE.Serial_Number,CS.RID as Status_RID,SI.Number FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name  INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE SI.RST = 'A' AND SI.Perso_Factory_RID = @Perso_Factory_RID AND SI.Date_Time>@From_Date_Time AND SI.Date_Time<=@End_Date_Time AND CE.Serial_Number = @Serial_Number UNION  All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,CE.Serial_Number,FCI.Status_RID ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID  WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time AND CE.Serial_Number = @Serial_Number ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID";

        public const string SEL_MATERIAL_USED_DM_REPLACE = " SELECT B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID,SUM(B.Number) AS Number FROM (  SELECT A.Perso_Factory_RID,A.Date_Time as Stock_Date, DI.Serial_Number,A.Status_RID, A.Number1 AS Number  FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID FROM  SUBTOTAL_REPLACE_IMPORT  SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time) AND (SI.Date_Time <= @End_Date_Time) AND    (SI.Perso_Factory_RID = @Perso_Factory_RID)) A  INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID  INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID   WHERE (DI.Card_Type_Link_Type = '1') AND (DI.Serial_Number = @Serial_Number)   UNION  All   SELECT A_1.Perso_Factory_RID,A_1.Date_Time as Stock_Date, DI.Serial_Number,A_1.Status_RID, A_1.Number1 AS Number   FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID  FROM   SUBTOTAL_REPLACE_IMPORT SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE  (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time)  AND (SI.Date_Time <= @End_Date_Time) AND (SI.Perso_Factory_RID = @Perso_Factory_RID)) A_1  INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID  INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID  INNER JOIN DM_CARDTYPE  DCT  ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID  WHERE (DI.Card_Type_Link_Type = '2') AND (DI.Serial_Number = @Serial_Number) ) B  GROUP BY B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID";

        public const string SEL_MATERIAL_USED_ENVELOPE_S_REPLACE = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,EI.Serial_Number,CS.RID as Status_RID,SI.Number  FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN')  INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name   INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE SI.RST = 'A'  AND SI.Date_Time<=@End_Date_Time AND SI.Date_Time>=@Start_Date_Time  UNION  All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,EI.Serial_Number,FCI.Status_RID  ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE FCI.RST = 'A'  AND FCI.Date_Time<=@End_Date_Time  AND FCI.Date_Time>=@Start_Date_Time  ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID";

        public const string SEL_MATERIAL_USED_CARD_EXPONENT_S_REPLACE = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,CE.Serial_Number,CS.RID as Status_RID,SI.Number FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name  INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE SI.RST = 'A'  AND SI.Date_Time<=@End_Date_Time  AND SI.Date_Time>=@Start_Date_Time  UNION  All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,CE.Serial_Number,FCI.Status_RID ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID  WHERE FCI.RST = 'A'  AND FCI.Date_Time<=@End_Date_Time   AND FCI.Date_Time>=@Start_Date_Time  ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID";

        public const string SEL_MATERIAL_USED_DM_S_REPLACE = " SELECT B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID,SUM(B.Number) AS Number FROM (  SELECT A.Perso_Factory_RID,A.Date_Time as Stock_Date, DI.Serial_Number,A.Status_RID, A.Number1 AS Number  FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID FROM  SUBTOTAL_REPLACE_IMPORT  SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE (SI.RST = 'A') AND (SI.Date_Time<=@End_Date_Time  AND SI.Date_Time>=@Start_Date_Time)     ) A  INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID  INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID   WHERE (DI.Card_Type_Link_Type = '1')    UNION   All  SELECT A_1.Perso_Factory_RID,A_1.Date_Time as Stock_Date, DI.Serial_Number,A_1.Status_RID, A_1.Number1 AS Number   FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID  FROM   SUBTOTAL_REPLACE_IMPORT SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE  (SI.RST = 'A')   AND (SI.Date_Time<=@End_Date_Time  AND SI.Date_Time>=@Start_Date_Time) ) A_1  INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID  INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID  INNER JOIN DM_CARDTYPE  DCT  ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID  WHERE (DI.Card_Type_Link_Type = '2') ) B  GROUP BY B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID";

        public const string SEL_MATERIAL_USED_ENVELOPE_DAYLY = " SELECT DM.Perso_Factory_RID,DM.CDay  AS Stock_Date,EI.Serial_Number,'1' as Status_RID,sum(DM.BNumber) +sum(DM.CNumber) + sum(DM.D1Number)+ sum(DM.D2Number)+ sum(DM.ENumber) as Number  FROM DAYLY_MONITOR DM  INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND DM.TYPE = CT.TYPE AND DM.AFFINITY = CT.AFFINITY AND DM.PHOTO = CT.PHOTO  INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE DM.RST = 'A'  AND  DM.Xtype='12' AND DM.Perso_Factory_RID = @Perso_Factory_RID AND DM.CDay>@From_Date_Time AND DM.CDay<=@End_Date_Time AND EI.Serial_Number = @Serial_Number  group by DM.Perso_Factory_RID,DM.CDay ,EI.Serial_Number ";

        public const string SEL_MATERIAL_USED_EXPONENT_DAYLY = " SELECT DM.Perso_Factory_RID,DM.CDay  AS Stock_Date,EI.Serial_Number,'1' as Status_RID,sum(DM.BNumber) +sum(DM.CNumber) + sum(DM.D1Number)+ sum(DM.D2Number)+ sum(DM.ENumber) as Number  FROM DAYLY_MONITOR DM  INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND DM.TYPE = CT.TYPE AND DM.AFFINITY = CT.AFFINITY AND DM.PHOTO = CT.PHOTO  INNER JOIN CARD_EXPONENT EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE DM.RST = 'A'  AND  DM.Xtype='12'  AND DM.Perso_Factory_RID = @Perso_Factory_RID AND DM.CDay>@From_Date_Time AND DM.CDay<=@End_Date_Time AND EI.Serial_Number = @Serial_Number  group by DM.Perso_Factory_RID,DM.CDay ,EI.Serial_Number ";

        public const string SEL_MATERIEL_STOCKS_TRANSACTION_INFO = "SELECT Serial_Number FROM MATERIEL_STOCKS_TRANSACTION WHERE Transaction_Date = @Transaction_Date";

        public const string SEL_MATERIEL_STOCKS_TRANSACTION_STOCKS = "SELECT MST.Transaction_Date,MST.Transaction_Amount \r\n                                    FROM MATERIEL_STOCKS_TRANSACTION MST\r\n                                    inner join PARAM P ON MST.PARAM_RID = P.RID \r\n                                    WHERE MST.RST = 'A' \r\n                                    AND P.RST = 'A' \r\n                                    AND MST.Factory_RID = @perso_factory_rid \r\n                                    AND MST.Serial_Number = @Serial_Number \r\n                                    AND MST.Transaction_Date > @lastSurplusDateTime \r\n                                    AND MST.Transaction_Date <= @thisSurplusDateTime \r\n                                    AND P.ParamType_Code = 'matType1' \r\n                                    AND P.Param_Code = @Param_Code";

        public const string DEL_MATERIEL_STOCKS_MANAGE = "DELETE \r\n                                    FROM MATERIEL_STOCKS_MANAGE \r\n                                    WHERE RST = 'A' and Perso_Factory_RID=@Perso_Factory_RID and Serial_Number=@Serial_Number \r\n                                    and Stock_Date > @lastSurplusDateTime ";

        public DataTable dtPeso;

        public string strErr;

        private Dictionary<string, object> dirValues = new Dictionary<string, object>();

        public InOut000BL()
        {
            string arg_79_0 = ConfigurationManager.AppSettings["FTPCardModifyReplace"];
            string arg_89_0 = ConfigurationManager.AppSettings["FTPCardModifyReplacePath"];
        }

        public void Material_Used_Warnning(string strFactory_RID, DateTime importDate, string WarningType)
        {
            try
            {
                DateTime lastSurplusDate = this.getLastSurplusDate();
                this.dirValues.Clear();
                this.dirValues.Add("Perso_Factory_RID", strFactory_RID);
                this.dirValues.Add("From_Date_Time", lastSurplusDate.ToString("yyyy/MM/dd 23:59:59"));
                this.dirValues.Add("End_Date_Time", importDate.ToString("yyyy/MM/dd 23:59:59"));
                DataSet arg_C3_0 = base.dao.GetList("SELECT * FROM (SELECT SI.Perso_Factory_RID,CT.RID,CS.RID AS Status_RID,SUM(SI.Number) AS Number FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name  WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time GROUP BY SI.Perso_Factory_RID,CT.RID,CS.RID UNION ALL SELECT FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID,SUM(FCI.Number) AS Number FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN')  INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time GROUP BY FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID ) A ORDER BY Perso_Factory_RID,RID,Status_RID ", this.dirValues);
                DataSet list = base.dao.GetList("SELECT ED.Operate,CS.RID FROM EXPRESSIONS_DEFINE ED INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND ED.Type_RID = CS.RID WHERE ED.RST = 'A' AND ED.Expressions_RID = 1");
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("Perso_Factory_RID");
                dataTable.Columns.Add("CardType_RID");
                dataTable.Columns.Add("Number");
                int num = 0;
                int num2 = 0;
                int num3 = 0;
                foreach (DataRow dataRow in arg_C3_0.Tables[0].Rows)
                {
                    if (Convert.ToInt32(dataRow["RID"]) != num || Convert.ToInt32(dataRow["Perso_Factory_RID"]) != num2)
                    {
                        if (num != 0 && num2 != 0 && num3 != 0)
                        {
                            DataRow dataRow2 = dataTable.NewRow();
                            dataRow2["Number"] = num3.ToString();
                            dataRow2["Perso_Factory_RID"] = num2.ToString();
                            dataRow2["CardType_RID"] = num.ToString();
                            dataTable.Rows.Add(dataRow2);
                        }
                        num3 = 0;
                        DataRow[] array = list.Tables[0].Select("RID = " + dataRow["Status_RID"].ToString());
                        if (array.Length != 0)
                        {
                            if (array[0]["Operate"].ToString() == "+")
                            {
                                num3 += Convert.ToInt32(dataRow["Number"]);
                                num = Convert.ToInt32(dataRow["RID"]);
                                num2 = Convert.ToInt32(dataRow["Perso_Factory_RID"]);
                            }
                            else if (array[0]["Operate"].ToString() == "-")
                            {
                                num3 -= Convert.ToInt32(dataRow["Number"]);
                                num = Convert.ToInt32(dataRow["RID"]);
                                num2 = Convert.ToInt32(dataRow["Perso_Factory_RID"]);
                            }
                        }
                    }
                    else
                    {
                        DataRow[] array2 = list.Tables[0].Select("RID = " + dataRow["Status_RID"].ToString());
                        if (array2.Length != 0)
                        {
                            if (array2[0]["Operate"].ToString() == "+")
                            {
                                num3 += Convert.ToInt32(dataRow["Number"]);
                            }
                            else if (array2[0]["Operate"].ToString() == "-")
                            {
                                num3 -= Convert.ToInt32(dataRow["Number"]);
                            }
                        }
                    }
                }
                if (num != 0 && num2 != 0 && num3 != 0)
                {
                    DataRow dataRow3 = dataTable.NewRow();
                    dataRow3["Number"] = num3.ToString();
                    dataRow3["Perso_Factory_RID"] = num2.ToString();
                    dataRow3["CardType_RID"] = num.ToString();
                    dataTable.Rows.Add(dataRow3);
                }
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", strFactory_RID);
                base.dao.ExecuteNonQuery("DELETE FROM TEMP_MADE_CARD WHERE Perso_Factory_RID = @perso_factory_rid", this.dirValues);
                foreach (DataRow dataRow4 in dataTable.Rows)
                {
                    this.dirValues.Clear();
                    this.dirValues.Add("Perso_Factory_RID", dataRow4["Perso_Factory_RID"].ToString());
                    this.dirValues.Add("CardType_RID", dataRow4["CardType_RID"].ToString());
                    this.dirValues.Add("Number", dataRow4["Number"].ToString());
                    base.dao.ExecuteNonQuery("INSERT INTO TEMP_MADE_CARD(Perso_Factory_RID,CardType_RID,Number)values(@Perso_Factory_RID,@CardType_RID,@Number)", this.dirValues);
                }
                DataTable materialUsed = this.getMaterialUsed(strFactory_RID, importDate);
                this.getMaterielStocks(lastSurplusDate, strFactory_RID, importDate, materialUsed, WarningType);
            }
            catch (Exception expr_49A)
            {
                LogFactory.Write(expr_49A.ToString(), "ErrLog");
                throw new Exception(expr_49A.Message);
            }
        }

        public void getMaterielStocks(DateTime dtLastWorkDate, string strFactory_RID, DateTime importDate, DataTable dtMATERIAL_USED, string WarningType)
        {
            try
            {
                foreach (DataRow dataRow in dtMATERIAL_USED.Rows)
                {
                    this.dirValues.Clear();
                    this.dirValues.Add("perso_factory_rid", strFactory_RID);
                    this.dirValues.Add("serial_number", dataRow["Serial_Number"].ToString());
                    DataSet list = base.dao.GetList("SELECT Top 1 MSM.Stock_Date,MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number,CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number WHERE Type = '4' AND MSM.Perso_Factory_RID = @perso_factory_rid AND MSM.Serial_Number = @serial_number ORDER BY Stock_Date Desc", this.dirValues);
                    if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                    {
                        this.dirValues.Clear();
                        this.dirValues.Add("perso_factory_rid", strFactory_RID);
                        this.dirValues.Add("serial_number", dataRow["Serial_Number"].ToString());
                        this.dirValues.Add("from_stock_date", Convert.ToDateTime(list.Tables[0].Rows[0]["Stock_Date"]).ToString("yyyy/MM/dd 23:59:59"));
                        this.dirValues.Add("end_stock_date", dtLastWorkDate.ToString("yyyy/MM/dd 23:59:59"));
                        DataSet list2 = base.dao.GetList("SELECT isnull(SUM(Number),0) as Number FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Serial_Number = @serial_number  AND Stock_Date>@from_stock_date AND Stock_Date<=@end_stock_date ", this.dirValues);
                        if (list2 != null && list2.Tables.Count > 0 && list2.Tables[0].Rows.Count > 0)
                        {
                            int num = Convert.ToInt32(list.Tables[0].Rows[0]["Number"].ToString());
                            int num2 = 0;
                            if (list2.Tables[0].Rows[0]["Number"] != DBNull.Value)
                            {
                                num2 = Convert.ToInt32(list2.Tables[0].Rows[0]["Number"]);
                            }
                            int num3 = Convert.ToInt32(dataRow["Number"]);
                            int num4 = 0;
                            int factory_RID = int.Parse(strFactory_RID);
                            string serial_Number = dataRow["Serial_Number"].ToString();
                            DateTime lastSurplusDateTime = DateTime.Parse(Convert.ToDateTime(list.Tables[0].Rows[0]["Stock_Date"]).ToString("yyyy/MM/dd 23:59:59"));
                            DateTime thisSurplusDateTime = DateTime.Parse(dtLastWorkDate.ToString("yyyy/MM/dd 23:59:59"));
                            DataTable dataTable = new DataTable();
                            dataTable = this.StocksMoveIn(factory_RID, serial_Number, lastSurplusDateTime, thisSurplusDateTime);
                            if (dataTable != null && dataTable.Rows.Count > 0)
                            {
                                foreach (DataRow dataRow2 in dataTable.Rows)
                                {
                                    num4 -= int.Parse(dataRow2["Move_Number"].ToString());
                                }
                            }
                            DataTable dataTable2 = new DataTable();
                            dataTable2 = this.StocksMoveOut(factory_RID, serial_Number, lastSurplusDateTime, thisSurplusDateTime);
                            if (dataTable2 != null && dataTable2.Rows.Count > 0)
                            {
                                foreach (DataRow dataRow3 in dataTable2.Rows)
                                {
                                    num4 += int.Parse(dataRow3["Move_Number"].ToString());
                                }
                            }
                            DataTable dataTable3 = new DataTable();
                            dataTable3 = this.StocksTransatrion(factory_RID, serial_Number, "1", lastSurplusDateTime, thisSurplusDateTime);
                            if (dataTable3 != null && dataTable3.Rows.Count > 0)
                            {
                                foreach (DataRow dataRow4 in dataTable3.Rows)
                                {
                                    num4 -= int.Parse(dataRow4["Transaction_Amount"].ToString());
                                }
                            }
                            DataTable dataTable4 = new DataTable();
                            dataTable4 = this.StocksTransatrion(factory_RID, serial_Number, "2", lastSurplusDateTime, thisSurplusDateTime);
                            if (dataTable4 != null && dataTable4.Rows.Count > 0)
                            {
                                foreach (DataRow dataRow5 in dataTable4.Rows)
                                {
                                    num4 += int.Parse(dataRow5["Transaction_Amount"].ToString());
                                }
                            }
                            DataTable dataTable5 = new DataTable();
                            dataTable5 = this.StocksTransatrion(factory_RID, serial_Number, "3", lastSurplusDateTime, thisSurplusDateTime);
                            if (dataTable5 != null && dataTable5.Rows.Count > 0)
                            {
                                foreach (DataRow dataRow6 in dataTable5.Rows)
                                {
                                    num4 += int.Parse(dataRow6["Transaction_Amount"].ToString());
                                }
                            }
                            if (num < num2 + num3 + num4)
                            {
                                if (this.DmNotSafe_Type(dataRow["Serial_Number"].ToString()))
                                {
                                    string[] arg = new string[]
                                    {
                                        list.Tables[0].Rows[0]["Name"].ToString()
                                    };
                                    if (WarningType == "1")
                                    {
                                        Warning.SetWarning("43", arg);
                                    }
                                    else
                                    {
                                        Warning.SetWarning("54", arg);
                                    }
                                }
                            }
                            else
                            {
                                DataSet materiel = this.GetMateriel(dataRow["Serial_Number"].ToString());
                                if (materiel != null && materiel.Tables.Count > 0 && materiel.Tables[0].Rows.Count > 0)
                                {
                                    if ("1" == Convert.ToString(materiel.Tables[0].Rows[0]["Safe_Type"]))
                                    {
                                        if (num - num3 - num2 < Convert.ToInt32(materiel.Tables[0].Rows[0]["Safe_Number"]))
                                        {
                                            string[] arg2 = new string[]
                                            {
                                                materiel.Tables[0].Rows[0]["Name"].ToString()
                                            };
                                            if (WarningType == "1")
                                            {
                                                Warning.SetWarning("44", arg2);
                                            }
                                            else
                                            {
                                                Warning.SetWarning("55", arg2);
                                            }
                                        }
                                    }
                                    else if ("2" == Convert.ToString(materiel.Tables[0].Rows[0]["Safe_Type"]) && !this.CheckMaterielSafeDays(dataRow["Serial_Number"].ToString(), Convert.ToInt32(dataRow["Perso_Factory_RID"].ToString()), Convert.ToInt32(materiel.Tables[0].Rows[0]["Safe_Number"]), num - num2 - num3))
                                    {
                                        string[] arg3 = new string[]
                                        {
                                            materiel.Tables[0].Rows[0]["Name"].ToString()
                                        };
                                        if (WarningType == "1")
                                        {
                                            Warning.SetWarning("44", arg3);
                                        }
                                        else
                                        {
                                            Warning.SetWarning("55", arg3);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expr_7D9)
            {
                LogFactory.Write(expr_7D9.ToString(), "ErrLog");
                throw new Exception(expr_7D9.Message);
            }
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
            catch (Exception expr_B6)
            {
                LogFactory.Write(expr_B6.ToString(), "ErrLog");
                throw new Exception(expr_B6.Message);
            }
            return result;
        }

        public bool CheckMaterielSafeDays(string Serial_Number, int Factory_RID, int Days, int Stock_Number)
        {
            bool result = true;
            Days++;
            DateTime lastSurplusDateTime = DateTime.Now.AddDays((double)(-(double)Days));
            DateTime thisSurplusDateTime = DateTime.Now.AddDays((double)(Days - 1));
            DataTable dataTable = new DataTable();
            if ("C" == Serial_Number.Substring(0, 1).ToUpper())
            {
                dataTable = this.MaterielUsedCount(Factory_RID, Serial_Number, lastSurplusDateTime, DateTime.Now);
            }
            else
            {
                dataTable = this.MaterielUsedCountDayly(Factory_RID, Serial_Number, DateTime.Now, thisSurplusDateTime);
            }
            int num = 0;
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    num += Convert.ToInt32(dataTable.Rows[i]["System_Num"]);
                }
                if (Stock_Number < num)
                {
                    result = false;
                }
            }
            return result;
        }

        public decimal GetWearRate(string Serial_Number)
        {
            decimal num = 0m;
            DataSet dataSet = null;
            decimal result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("Serial_Number", Serial_Number);
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
                if (dataSet != null && dataSet.Tables.Count > 0 && dataSet.Tables[0].Rows.Count > 0)
                {
                    num = Convert.ToDecimal(dataSet.Tables[0].Rows[0]["Wear_Rate"]);
                }
                result = num;
            }
            catch (Exception expr_10F)
            {
                LogFactory.Write(expr_10F.ToString(), "ErrLog");
                throw new Exception(expr_10F.Message);
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
            catch (Exception expr_76)
            {
                LogFactory.Write(expr_76.ToString(), "ErrLog");
                throw new Exception(expr_76.Message);
            }
            return result;
        }

        public DataTable getMaterialUsed(string strFactory_RID, DateTime importDate)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("Stock_Date", Type.GetType("System.DateTime"));
            dataTable.Columns.Add("Number", Type.GetType("System.Int32"));
            dataTable.Columns.Add("Serial_Number", Type.GetType("System.String"));
            dataTable.Columns.Add("Perso_Factory_RID", Type.GetType("System.Int32"));
            DataTable result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", strFactory_RID);
                foreach (DataRow dataRow in base.dao.GetList(" SELECT EI.Serial_Number AS EI_Number,CE.Serial_Number as CE_Number,TMC.Perso_Factory_RID,TMC.Number FROM TEMP_MADE_CARD TMC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND TMC.CardType_RID = CT.RID LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE TMC.Perso_Factory_RID = @perso_factory_rid", this.dirValues).Tables[0].Rows)
                {
                    if (dataRow["CE_Number"].ToString() != "")
                    {
                        DataRow[] array = dataTable.Select("Serial_Number = '" + dataRow["CE_Number"].ToString() + "'");
                        int num = Convert.ToInt32(dataRow["Number"]);
                        if (array.Length != 0)
                        {
                            array[0]["Number"] = Convert.ToInt32(array[0]["Number"]) + num;
                        }
                        else
                        {
                            DataRow dataRow2 = dataTable.NewRow();
                            dataRow2["Stock_Date"] = importDate;
                            dataRow2["Number"] = num;
                            dataRow2["Serial_Number"] = dataRow["CE_Number"].ToString();
                            dataRow2["Perso_Factory_RID"] = Convert.ToInt32(dataRow["Perso_Factory_RID"]);
                            dataTable.Rows.Add(dataRow2);
                        }
                    }
                    if (dataRow["EI_Number"].ToString() != "")
                    {
                        DataRow[] array2 = dataTable.Select("Serial_Number = '" + dataRow["EI_Number"].ToString() + "'");
                        int num2 = Convert.ToInt32(dataRow["Number"]);
                        if (array2.Length != 0)
                        {
                            array2[0]["Number"] = Convert.ToInt32(array2[0]["Number"]) + num2;
                        }
                        else
                        {
                            DataRow dataRow3 = dataTable.NewRow();
                            dataRow3["Stock_Date"] = importDate;
                            dataRow3["Number"] = num2;
                            dataRow3["Serial_Number"] = dataRow["EI_Number"].ToString();
                            dataRow3["Perso_Factory_RID"] = Convert.ToInt32(dataRow["Perso_Factory_RID"]);
                            dataTable.Rows.Add(dataRow3);
                        }
                    }
                }
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", strFactory_RID);
                foreach (DataRow dataRow4 in base.dao.GetList(" SELECT DI.Serial_Number AS DI_Number,A.Perso_Factory_RID,A.Number  FROM TEMP_MADE_CARD A  INNER JOIN DM_CARDTYPE DCT ON DCT.RST = 'A' AND A.CardType_RID = DCT.CardType_RID  INNER JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND DCT.DM_RID = DI.RID  WHERE A.Perso_Factory_RID = @perso_factory_rid ", this.dirValues).Tables[0].Rows)
                {
                    if (dataRow4["DI_Number"].ToString() != "")
                    {
                        DataRow[] array3 = dataTable.Select("Serial_Number = '" + dataRow4["DI_Number"].ToString() + "'");
                        int num3 = Convert.ToInt32(dataRow4["Number"]);
                        if (array3.Length != 0)
                        {
                            array3[0]["Number"] = Convert.ToInt32(array3[0]["Number"]) + num3;
                        }
                        else
                        {
                            DataRow dataRow5 = dataTable.NewRow();
                            dataRow5["Stock_Date"] = importDate;
                            dataRow5["Number"] = num3;
                            dataRow5["Serial_Number"] = dataRow4["DI_Number"].ToString();
                            dataRow5["Perso_Factory_RID"] = Convert.ToInt32(dataRow4["Perso_Factory_RID"]);
                            dataTable.Rows.Add(dataRow5);
                        }
                    }
                }
                result = dataTable;
            }
            catch (Exception expr_46A)
            {
                LogFactory.Write(expr_46A.ToString(), "ErrLog");
                throw new Exception(expr_46A.Message);
            }
            return result;
        }

        private void SendWarningPerso(DataTable dtImport, string sFactoryRid)
        {
            try
            {
                DataTable cardType = this.GetCardType();
                DataTable dataTable = this.GetFactoryList().Tables[0];
                DataTable dataTable2 = new DataTable();
                dataTable2.Columns.Add("card");
                dataTable2.Columns.Add("factory");
                DataTable dataTable3 = base.dao.GetList("select CardType_RID from dbo.GROUP_CARD_TYPE a inner join CARD_GROUP b on a.Group_rid=b.rid where b.Group_Name = '虛擬卡'").Tables[0];
                foreach (DataRow dataRow in dtImport.Rows)
                {
                    DataRow[] array = cardType.Select(string.Concat(new string[]
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
                        DataRow[] array2 = dataTable.Select("RID='" + sFactoryRid + "'");
                        if (dataTable2.Select(string.Concat(new string[]
                        {
                            "card='",
                            cardTypeRid.ToString(),
                            "' and factory='",
                            sFactoryRid.ToString(),
                            "'"
                        })).Length == 0 && (dataTable3.Rows.Count <= 0 || dataTable3.Select("CardType_RID = '" + cardTypeRid.ToString() + "'").Length == 0))
                        {
                            DataRow dataRow2 = dataTable2.NewRow();
                            dataRow2[0] = cardTypeRid.ToString();
                            dataRow2[1] = sFactoryRid.ToString();
                            dataTable2.Rows.Add(dataRow2);
                            if (new CardTypeManager().getCurrentStockPersoReplace(Convert.ToInt32(sFactoryRid), cardTypeRid, DateTime.Now.Date.AddDays(1.0).AddSeconds(-1.0)) < 0)
                            {
                                object[] array3 = new object[2];
                                array3[0] = array2[0]["Factory_Shortname_CN"];
                                DataRow[] array4 = cardType.Select("RID=" + cardTypeRid.ToString());
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
                    new DataColumn("Name", Type.GetType("System.String")),
                    new DataColumn("Status_RID", Type.GetType("System.String")),
                    new DataColumn("Number", Type.GetType("System.String")),
                    new DataColumn("Factory_RID", Type.GetType("System.String"))
                }
            };
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

        public string GetFactoryList_RID(string Factory_ID)
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
            catch (Exception expr_24)
            {
                LogFactory.Write(expr_24.ToString(), "ErrLog");
                throw new Exception(expr_24.Message);
            }
            return list;
        }

        private DataTable GetCardType()
        {
            DataTable result;
            try
            {
                DataTable dataTable = new DataTable();
                dataTable = base.dao.GetList("SELECT * FROM CARD_TYPE WHERE RST='A'").Tables[0];
                if (dataTable.Rows.Count > 0)
                {
                    result = dataTable;
                }
                else
                {
                    result = null;
                }
            }
            catch (Exception arg_38_0)
            {
                LogFactory.Write(arg_38_0.ToString(), "ErrLog");
                result = null;
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
            catch (Exception expr_20)
            {
                LogFactory.Write(expr_20.ToString(), "ErrLog");
                throw new Exception(expr_20.Message);
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
            catch (Exception expr_7C)
            {
                LogFactory.Write(expr_7C.ToString(), "ErrLog");
                throw new Exception(expr_7C.Message);
            }
            return result;
        }

        private int ComputeMaterialNumber(string MNumber, long MCount)
        {
            int result;
            try
            {
                decimal wearRate = this.GetWearRate(MNumber);
                result = Convert.ToInt32(MCount * (wearRate / 100m + decimal.One));
            }
            catch (Exception expr_32)
            {
                LogFactory.Write(expr_32.ToString(), "ErrLog");
                throw new Exception(expr_32.Message);
            }
            return result;
        }

        public DataTable MaterielUsedCount(int Factory_RID, string Serial_Number, DateTime lastSurplusDateTime, DateTime thisSurplusDateTime)
        {
            DataTable dataTable = null;
            DataTable result;
            try
            {
                string text = Serial_Number.Substring(0, 1).ToUpper();
                string commandText;
                if (text.Equals("A"))
                {
                    commandText = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,EI.Serial_Number,CS.RID as Status_RID,SI.Number  FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN')  INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name   INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE SI.RST = 'A' AND SI.Perso_Factory_RID = @Perso_Factory_RID AND SI.Date_Time>@From_Date_Time AND SI.Date_Time<=@End_Date_Time AND EI.Serial_Number = @Serial_Number UNION All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,EI.Serial_Number,FCI.Status_RID  ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time AND EI.Serial_Number = @Serial_Number ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID";
                }
                else if (text.Equals("B"))
                {
                    commandText = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,CE.Serial_Number,CS.RID as Status_RID,SI.Number FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name  INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE SI.RST = 'A' AND SI.Perso_Factory_RID = @Perso_Factory_RID AND SI.Date_Time>@From_Date_Time AND SI.Date_Time<=@End_Date_Time AND CE.Serial_Number = @Serial_Number UNION  All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,CE.Serial_Number,FCI.Status_RID ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID  WHERE FCI.RST = 'A' AND FCI.Perso_Factory_RID = @Perso_Factory_RID AND FCI.Date_Time>@From_Date_Time AND FCI.Date_Time<=@End_Date_Time AND CE.Serial_Number = @Serial_Number ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID";
                }
                else
                {
                    if (!text.Equals("C"))
                    {
                        result = null;
                        return result;
                    }
                    commandText = " SELECT B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID,SUM(B.Number) AS Number FROM (  SELECT A.Perso_Factory_RID,A.Date_Time as Stock_Date, DI.Serial_Number,A.Status_RID, A.Number1 AS Number  FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID FROM  SUBTOTAL_REPLACE_IMPORT  SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time) AND (SI.Date_Time <= @End_Date_Time) AND    (SI.Perso_Factory_RID = @Perso_Factory_RID)) A  INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID  INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID   WHERE (DI.Card_Type_Link_Type = '1') AND (DI.Serial_Number = @Serial_Number)   UNION  All   SELECT A_1.Perso_Factory_RID,A_1.Date_Time as Stock_Date, DI.Serial_Number,A_1.Status_RID, A_1.Number1 AS Number   FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID  FROM   SUBTOTAL_REPLACE_IMPORT SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE  (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time)  AND (SI.Date_Time <= @End_Date_Time) AND (SI.Perso_Factory_RID = @Perso_Factory_RID)) A_1  INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID  INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID  INNER JOIN DM_CARDTYPE  DCT  ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID  WHERE (DI.Card_Type_Link_Type = '2') AND (DI.Serial_Number = @Serial_Number) ) B  GROUP BY B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID";
                }
                this.dirValues.Clear();
                this.dirValues.Add("Perso_Factory_RID", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("From_Date_Time", lastSurplusDateTime);
                this.dirValues.Add("End_Date_Time", thisSurplusDateTime);
                DataSet list = base.dao.GetList(commandText, this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = this.Material_Used_Num(list.Tables[0]);
                }
                result = dataTable;
            }
            catch (Exception expr_111)
            {
                LogFactory.Write(expr_111.ToString(), "ErrLog");
                throw new Exception(expr_111.Message);
            }
            return result;
        }

        public DataTable MaterielUsedCountDayly(int Factory_RID, string Serial_Number, DateTime lastSurplusDateTime, DateTime thisSurplusDateTime)
        {
            DataTable dataTable = null;
            DataTable result;
            try
            {
                string text = Serial_Number.Substring(0, 1).ToUpper();
                string commandText;
                if (text.Equals("A"))
                {
                    commandText = " SELECT DM.Perso_Factory_RID,DM.CDay  AS Stock_Date,EI.Serial_Number,'1' as Status_RID,sum(DM.BNumber) +sum(DM.CNumber) + sum(DM.D1Number)+ sum(DM.D2Number)+ sum(DM.ENumber) as Number  FROM DAYLY_MONITOR DM  INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND DM.TYPE = CT.TYPE AND DM.AFFINITY = CT.AFFINITY AND DM.PHOTO = CT.PHOTO  INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE DM.RST = 'A'  AND  DM.Xtype='12' AND DM.Perso_Factory_RID = @Perso_Factory_RID AND DM.CDay>@From_Date_Time AND DM.CDay<=@End_Date_Time AND EI.Serial_Number = @Serial_Number  group by DM.Perso_Factory_RID,DM.CDay ,EI.Serial_Number ";
                }
                else
                {
                    if (!text.Equals("B"))
                    {
                        result = null;
                        return result;
                    }
                    commandText = " SELECT DM.Perso_Factory_RID,DM.CDay  AS Stock_Date,EI.Serial_Number,'1' as Status_RID,sum(DM.BNumber) +sum(DM.CNumber) + sum(DM.D1Number)+ sum(DM.D2Number)+ sum(DM.ENumber) as Number  FROM DAYLY_MONITOR DM  INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND DM.TYPE = CT.TYPE AND DM.AFFINITY = CT.AFFINITY AND DM.PHOTO = CT.PHOTO  INNER JOIN CARD_EXPONENT EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE DM.RST = 'A'  AND  DM.Xtype='12'  AND DM.Perso_Factory_RID = @Perso_Factory_RID AND DM.CDay>@From_Date_Time AND DM.CDay<=@End_Date_Time AND EI.Serial_Number = @Serial_Number  group by DM.Perso_Factory_RID,DM.CDay ,EI.Serial_Number ";
                }
                this.dirValues.Clear();
                this.dirValues.Add("Perso_Factory_RID", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("From_Date_Time", lastSurplusDateTime);
                this.dirValues.Add("End_Date_Time", thisSurplusDateTime);
                DataSet list = base.dao.GetList(commandText, this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = this.Material_Used_Num(list.Tables[0]);
                }
                result = dataTable;
            }
            catch (Exception expr_FC)
            {
                LogFactory.Write(expr_FC.ToString(), "ErrLog");
                throw new Exception(expr_FC.Message);
            }
            return result;
        }

        public DataTable DayMaterielUsedCount(int Factory_RID, string Serial_Number, DateTime lastSurplusDateTime, DateTime thisSurplusDateTime)
        {
            DataTable dataTable = null;
            DataTable result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("Perso_Factory_RID", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", lastSurplusDateTime);
                this.dirValues.Add("thisSurplusDateTime", thisSurplusDateTime);
                DataSet list = base.dao.GetList("select isnull(sum(Number),0) as System_Num from MATERIEL_STOCKS_USED where rst='A' AND Serial_Number=@Serial_Number AND Perso_Factory_RID=@Perso_Factory_RID AND Stock_Date > @lastSurplusDateTime AND Stock_Date <= @thisSurplusDateTime", this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = list.Tables[0];
                }
                result = dataTable;
            }
            catch (Exception expr_B3)
            {
                LogFactory.Write(expr_B3.ToString(), "ErrLog");
                throw new Exception(expr_B3.Message);
            }
            return result;
        }

        public bool DayCheckMaterielSafeDays(string Serial_Number, int Factory_RID, int Days, int Stock_Number)
        {
            bool result = true;
            Days++;
            DateTime lastSurplusDateTime = DateTime.Now.AddDays((double)(-(double)Days));
            DateTime thisSurplusDateTime = DateTime.Now.AddDays((double)(Days - 1));
            DataTable dataTable = new DataTable();
            if ("C" == Serial_Number.Substring(0, 1).ToUpper())
            {
                dataTable = this.DayMaterielUsedCount(Factory_RID, Serial_Number, lastSurplusDateTime, DateTime.Now);
            }
            else
            {
                dataTable = this.MaterielUsedCountDayly(Factory_RID, Serial_Number, DateTime.Now, thisSurplusDateTime);
            }
            int num = 0;
            if (dataTable != null && dataTable.Rows.Count > 0)
            {
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    num += Convert.ToInt32(dataTable.Rows[i]["System_Num"]);
                }
                if (Stock_Number < num)
                {
                    result = false;
                }
            }
            return result;
        }

        public List<string> SaveMaterielUsedCount(DateTime Date)
        {
            List<string> list = new List<string>();
            MATERIEL_STOCKS_USED mATERIEL_STOCKS_USED = new MATERIEL_STOCKS_USED();
            List<string> result;
            try
            {
                string commandText = " SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,EI.Serial_Number,CS.RID as Status_RID,SI.Number  FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN')  INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name   INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE SI.RST = 'A'  AND SI.Date_Time<=@End_Date_Time AND SI.Date_Time>=@Start_Date_Time  UNION  All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,EI.Serial_Number,FCI.Status_RID  ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID  WHERE FCI.RST = 'A'  AND FCI.Date_Time<=@End_Date_Time  AND FCI.Date_Time>=@Start_Date_Time  ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID SELECT A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID,SUM(A.Number) AS Number FROM (  SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,CE.Serial_Number,CS.RID as Status_RID,SI.Number FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') INNER JOIN CARDTYPE_STATUS CS ON CS.RST='A' AND CS.Status_Name=MCT.Type_Name  INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID WHERE SI.RST = 'A'  AND SI.Date_Time<=@End_Date_Time  AND SI.Date_Time>=@Start_Date_Time  UNION  All  SELECT FCI.Perso_Factory_RID,FCI.Date_Time  AS Stock_Date,CE.Serial_Number,FCI.Status_RID ,FCI.Number  FROM FACTORY_CHANGE_REPLACE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO  INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID AND CS.Status_Name NOT IN ('3D','DA','PM','RN') INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID  WHERE FCI.RST = 'A'  AND FCI.Date_Time<=@End_Date_Time   AND FCI.Date_Time>=@Start_Date_Time  ) A  GROUP BY A.Perso_Factory_RID,A.Stock_Date,A.Serial_Number,A.Status_RID SELECT B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID,SUM(B.Number) AS Number FROM (  SELECT A.Perso_Factory_RID,A.Date_Time as Stock_Date, DI.Serial_Number,A.Status_RID, A.Number1 AS Number  FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID FROM  SUBTOTAL_REPLACE_IMPORT  SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE (SI.RST = 'A') AND (SI.Date_Time<=@End_Date_Time  AND SI.Date_Time>=@Start_Date_Time)     ) A  INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID  INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID   WHERE (DI.Card_Type_Link_Type = '1')    UNION   All  SELECT A_1.Perso_Factory_RID,A_1.Date_Time as Stock_Date, DI.Serial_Number,A_1.Status_RID, A_1.Number1 AS Number   FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID,1 as Status_RID  FROM   SUBTOTAL_REPLACE_IMPORT SI  INNER JOIN CARD_TYPE  CT  ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO   WHERE  (SI.RST = 'A')   AND (SI.Date_Time<=@End_Date_Time  AND SI.Date_Time>=@Start_Date_Time) ) A_1  INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID  INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID  INNER JOIN DM_CARDTYPE  DCT  ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID  WHERE (DI.Card_Type_Link_Type = '2') ) B  GROUP BY B.Perso_Factory_RID,B.Stock_Date,B.Serial_Number,B.Status_RID";
                string commandText2 = "Delete MATERIEL_STOCKS_USED where Stock_Date >=@Start_Date_Time and Stock_Date <=@End_Date_Time ";
                this.dirValues.Clear();
                this.dirValues.Add("Start_Date_Time", Date.ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("End_Date_Time", Date.ToString("yyyy/MM/dd 23:59:59"));
                DataSet list2 = base.dao.GetList(commandText, this.dirValues);
                base.dao.ExecuteNonQuery(commandText2, this.dirValues);
                for (int i = 0; i < list2.Tables.Count; i++)
                {
                    if (list2 != null && list2.Tables.Count > 0 && list2.Tables[i].Rows.Count > 0)
                    {
                        DataTable dataTable = this.Material_Used_Num(list2.Tables[i]);
                        if (dataTable != null && dataTable.Rows.Count > 0)
                        {
                            for (int j = 0; j < dataTable.Rows.Count; j++)
                            {
                                if (dataTable.Rows[j]["Serial_Number"].ToString() != null && dataTable.Rows[j]["Serial_Number"].ToString() != "")
                                {
                                    if (-1 == list.IndexOf(dataTable.Rows[j]["Serial_Number"].ToString()))
                                    {
                                        list.Add(dataTable.Rows[j]["Serial_Number"].ToString());
                                    }
                                    if (dataTable.Rows[j]["Stock_Date"].ToString() != null && dataTable.Rows[j]["Stock_Date"].ToString() != "")
                                    {
                                        mATERIEL_STOCKS_USED.Stock_Date = DateTime.Parse(dataTable.Rows[j]["Stock_Date"].ToString());
                                    }
                                    mATERIEL_STOCKS_USED.Number = long.Parse(dataTable.Rows[j]["System_Num"].ToString());
                                    mATERIEL_STOCKS_USED.Serial_Number = dataTable.Rows[j]["Serial_Number"].ToString();
                                    mATERIEL_STOCKS_USED.Perso_Factory_RID = Convert.ToInt32(dataTable.Rows[j]["Perso_Factory_RID"].ToString());
                                    base.dao.Add<MATERIEL_STOCKS_USED>(mATERIEL_STOCKS_USED, "RID");
                                }
                            }
                        }
                    }
                }
                result = list;
            }
            catch (Exception expr_2B0)
            {
                LogFactory.Write(expr_2B0.ToString(), "ErrLog");
                throw new Exception(expr_2B0.Message);
            }
            return result;
        }

        public DataTable Material_Used_Num(DataTable dsMade_Card)
        {
            DataTable result;
            try
            {
                DataSet list = base.dao.GetList("SELECT ED.Operate,CS.RID FROM EXPRESSIONS_DEFINE ED INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND ED.Type_RID = CS.RID WHERE ED.RST = 'A' AND ED.Expressions_RID = 1");
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("Perso_Factory_RID");
                dataTable.Columns.Add("Stock_Date");
                dataTable.Columns.Add("Serial_Number");
                dataTable.Columns.Add("Number");
                DateTime dateTime = Convert.ToDateTime("1900-01-01");
                string text = "";
                int num = 0;
                int num2 = 0;
                foreach (DataRow dataRow in dsMade_Card.Rows)
                {
                    if (Convert.ToDateTime(dataRow["Stock_Date"]) != dateTime || Convert.ToString(dataRow["Serial_Number"]) != text || Convert.ToInt32(dataRow["Perso_Factory_RID"]) != num)
                    {
                        if (dateTime != Convert.ToDateTime("1900-01-01") && text != "" && num2 != 0 && num != 0)
                        {
                            DataRow dataRow2 = dataTable.NewRow();
                            dataRow2["Perso_Factory_RID"] = num.ToString();
                            dataRow2["Number"] = num2.ToString();
                            dataRow2["Stock_Date"] = dateTime.ToString();
                            dataRow2["Serial_Number"] = text.ToString();
                            dataTable.Rows.Add(dataRow2);
                        }
                        num2 = 0;
                        DataRow[] array = list.Tables[0].Select("RID = " + dataRow["Status_RID"].ToString());
                        if (array.Length != 0)
                        {
                            if (array[0]["Operate"].ToString() == "+")
                            {
                                num2 += Convert.ToInt32(dataRow["Number"]);
                                num = Convert.ToInt32(dataRow["Perso_Factory_RID"]);
                                dateTime = Convert.ToDateTime(dataRow["Stock_Date"]);
                                text = Convert.ToString(dataRow["Serial_Number"]);
                            }
                            else if (array[0]["Operate"].ToString() == "-")
                            {
                                num2 -= Convert.ToInt32(dataRow["Number"]);
                                num = Convert.ToInt32(dataRow["Perso_Factory_RID"]);
                                dateTime = Convert.ToDateTime(dataRow["Stock_Date"]);
                                text = Convert.ToString(dataRow["Serial_Number"]);
                            }
                        }
                    }
                    else
                    {
                        DataRow[] array2 = list.Tables[0].Select("RID = " + dataRow["Status_RID"].ToString());
                        if (array2.Length != 0)
                        {
                            if (array2[0]["Operate"].ToString() == "+")
                            {
                                num2 += Convert.ToInt32(dataRow["Number"]);
                            }
                            else if (array2[0]["Operate"].ToString() == "-")
                            {
                                num2 -= Convert.ToInt32(dataRow["Number"]);
                            }
                        }
                    }
                }
                if (dateTime != Convert.ToDateTime("1900-01-01") && text != "" && num2 != 0 && num != 0)
                {
                    DataRow dataRow3 = dataTable.NewRow();
                    dataRow3["Perso_Factory_RID"] = num.ToString();
                    dataRow3["Number"] = num2.ToString();
                    dataRow3["Stock_Date"] = dateTime.ToString();
                    dataRow3["Serial_Number"] = text.ToString();
                    dataTable.Rows.Add(dataRow3);
                }
                if (dataTable.Rows.Count > 0)
                {
                    dataTable.Columns.Add(new DataColumn("System_Num", Type.GetType("System.Int32")));
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        dataTable.Rows[i]["System_Num"] = Convert.ToInt32(dataTable.Rows[i]["Number"].ToString());
                    }
                }
                result = dataTable;
            }
            catch (Exception arg_453_0)
            {
                LogFactory.Write(arg_453_0.ToString(), "ErrLog");
                result = null;
            }
            return result;
        }

        public void getDayMaterielStocks(DateTime Surplus_Date, List<string> lstSerielNumber)
        {
            try
            {
                if (lstSerielNumber.Count > 0)
                {
                    string text = "'";
                    foreach (string current in lstSerielNumber)
                    {
                        text = text + current + "','";
                    }
                    text = text.Substring(0, text.Length - 2);
                    DateTime lastWorkDate = this.WorkDateBL.GetLastWorkDate(Surplus_Date);
                    this.dirValues.Clear();
                    this.dirValues.Add("stock_date", lastWorkDate.ToString("yyyy/MM/dd"));
                    DataSet list = base.dao.GetList("SELECT MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number,CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME FROM MATERIEL_STOCKS MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number WHERE Convert(varchar(10),Stock_Date,111) = @stock_date AND MSM.Serial_Number IN (" + text + ")", this.dirValues);
                    if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dataRow in list.Tables[0].Rows)
                        {
                            this.dirValues.Clear();
                            this.dirValues.Add("stock_date", Surplus_Date.ToString("yyyy/MM/dd"));
                            this.dirValues.Add("serial_number", dataRow["Serial_Number"].ToString());
                            this.dirValues.Add("perso_factory_rid", dataRow["Perso_Factory_RID"].ToString());
                            DataSet list2 = base.dao.GetList("SELECT isnull(sum(Number),0) FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Serial_Number = @serial_number AND Convert(varchar(10),Stock_Date,111) = @stock_date ", this.dirValues);
                            if (list2 != null && list2.Tables.Count > 0 && list2.Tables[0].Rows.Count > 0)
                            {
                                int num = 0;
                                int num2 = 0;
                                if (!StringUtil.IsEmpty(dataRow["Number"].ToString()))
                                {
                                    num = Convert.ToInt32(dataRow["Number"].ToString());
                                }
                                if (!StringUtil.IsEmpty(list2.Tables[0].Rows[0][0].ToString()))
                                {
                                    num2 = Convert.ToInt32(list2.Tables[0].Rows[0][0].ToString());
                                }
                                int factory_RID = int.Parse(dataRow["Perso_Factory_RID"].ToString());
                                string serial_Number = dataRow["Serial_Number"].ToString();
                                DateTime lastSurplusDateTime = DateTime.Parse(lastWorkDate.ToString("yyyy/MM/dd 23:59:59"));
                                DateTime thisSurplusDateTime = DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd 23:59:59"));
                                DataTable dataTable = new DataTable();
                                dataTable = this.StocksMoveIn(factory_RID, serial_Number, lastSurplusDateTime, thisSurplusDateTime);
                                if (dataTable != null && dataTable.Rows.Count > 0)
                                {
                                    foreach (DataRow dataRow2 in dataTable.Rows)
                                    {
                                        num2 -= int.Parse(dataRow2["Move_Number"].ToString());
                                    }
                                }
                                DataTable dataTable2 = new DataTable();
                                dataTable2 = this.StocksMoveOut(factory_RID, serial_Number, lastSurplusDateTime, thisSurplusDateTime);
                                if (dataTable2 != null && dataTable2.Rows.Count > 0)
                                {
                                    foreach (DataRow dataRow3 in dataTable2.Rows)
                                    {
                                        num2 += int.Parse(dataRow3["Move_Number"].ToString());
                                    }
                                }
                                DataTable dataTable3 = new DataTable();
                                dataTable3 = this.StocksTransatrion(factory_RID, serial_Number, "1", lastSurplusDateTime, thisSurplusDateTime);
                                if (dataTable3 != null && dataTable3.Rows.Count > 0)
                                {
                                    foreach (DataRow dataRow4 in dataTable3.Rows)
                                    {
                                        num2 -= int.Parse(dataRow4["Transaction_Amount"].ToString());
                                    }
                                }
                                DataTable dataTable4 = new DataTable();
                                dataTable4 = this.StocksTransatrion(factory_RID, serial_Number, "2", lastSurplusDateTime, thisSurplusDateTime);
                                if (dataTable4 != null && dataTable4.Rows.Count > 0)
                                {
                                    foreach (DataRow dataRow5 in dataTable4.Rows)
                                    {
                                        num2 += int.Parse(dataRow5["Transaction_Amount"].ToString());
                                    }
                                }
                                DataTable dataTable5 = new DataTable();
                                dataTable5 = this.StocksTransatrion(factory_RID, serial_Number, "3", lastSurplusDateTime, thisSurplusDateTime);
                                if (dataTable5 != null && dataTable5.Rows.Count > 0)
                                {
                                    foreach (DataRow dataRow6 in dataTable5.Rows)
                                    {
                                        num2 += int.Parse(dataRow6["Transaction_Amount"].ToString());
                                    }
                                }
                                if (num < num2)
                                {
                                    if (this.DmNotSafe_Type(dataRow["Serial_Number"].ToString()))
                                    {
                                        string[] arg = new string[]
                                        {
                                            dataRow["Name"].ToString()
                                        };
                                        Warning.SetWarning("45", arg);
                                        Warning.SetWarning("43", arg);
                                    }
                                }
                                else
                                {
                                    DataSet materiel = this.GetMateriel(dataRow["Serial_Number"].ToString());
                                    if (materiel != null && materiel.Tables.Count > 0 && materiel.Tables[0].Rows.Count > 0)
                                    {
                                        if ("1" == Convert.ToString(materiel.Tables[0].Rows[0]["Safe_Type"]))
                                        {
                                            if (num - num2 < Convert.ToInt32(materiel.Tables[0].Rows[0]["Safe_Number"]))
                                            {
                                                string[] arg2 = new string[]
                                                {
                                                    materiel.Tables[0].Rows[0]["Name"].ToString()
                                                };
                                                Warning.SetWarning("44", arg2);
                                                Warning.SetWarning("46", arg2);
                                            }
                                        }
                                        else if ("2" == Convert.ToString(materiel.Tables[0].Rows[0]["Safe_Type"]) && !this.DayCheckMaterielSafeDays(dataRow["Serial_Number"].ToString(), Convert.ToInt32(dataRow["Perso_Factory_RID"].ToString()), Convert.ToInt32(materiel.Tables[0].Rows[0]["Safe_Number"]), num - num2))
                                        {
                                            string[] arg3 = new string[]
                                            {
                                                materiel.Tables[0].Rows[0]["Name"].ToString()
                                            };
                                            Warning.SetWarning("44", arg3);
                                            Warning.SetWarning("46", arg3);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expr_7A9)
            {
                LogFactory.Write(expr_7A9.ToString(), "ErrLog");
                throw expr_7A9;
            }
        }

        public bool DmNotSafe_Type(string strSerial_Number)
        {
            bool result = true;
            try
            {
                if (strSerial_Number.Contains("C"))
                {
                    DataTable dataTable = base.dao.GetList("select * from dbo.DMTYPE_INFO where Serial_Number='" + strSerial_Number + "'").Tables[0];
                    if (dataTable.Rows.Count > 0 && dataTable.Rows[0]["Safe_Type"].ToString() == "3")
                    {
                        result = false;
                    }
                }
            }
            catch
            {
                return true;
            }
            return result;
        }

        public void SaveSurplusSystemNum(DateTime SurplusDate)
        {
            try
            {
                DateTime lastWorkDate = this.WorkDateBL.GetLastWorkDate(SurplusDate);
                this.dirValues.Clear();
                this.dirValues.Add("lastSurplusDateTime", lastWorkDate.ToString("yyyy/MM/dd 23:59:59"));
                DataTable dataTable = new DataTable();
                DataSet list = base.dao.GetList("select Perso_Factory_RID,Serial_Number from ( SELECT MS.Perso_Factory_RID,MS.Serial_Number from MATERIEL_STOCKS MS where MS.RST = 'A' and MS.Stock_Date>=@lastSurplusDateTime union SELECT MSU.Perso_Factory_RID,MSU.Serial_Number from MATERIEL_STOCKS_USED MSU where MSU.RST = 'A' and MSU.Stock_Date>=@lastSurplusDateTime ) A order by Perso_Factory_RID,Serial_Number ", this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = list.Tables[0];
                    DateTime dateTime = DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd 12:00:00"));
                    this.dirValues.Clear();
                    DataTable dataTable2 = new DataTable();
                    DataSet list2 = base.dao.GetList("select MSM.Perso_Factory_RID,MSM.Serial_number,MSM.Stock_Date,MSM.Number  from MATERIEL_STOCKS_MANAGE MSM  inner join (  select Perso_Factory_RID,Serial_number,max(Stock_Date) as Stock_Date  from MATERIEL_STOCKS_MANAGE   where type=4  group by Perso_Factory_RID,Serial_number  ) A  on A.Perso_Factory_RID=MSM.Perso_Factory_RID and A.Serial_number=MSM.Serial_number and A.Stock_Date=MSM.Stock_Date and MSM.type = 4  order by MSM.Perso_Factory_RID,MSM.Serial_number  ", this.dirValues);
                    if (list2 != null && list2.Tables.Count > 0 && list2.Tables[0].Rows.Count > 0)
                    {
                        dataTable2 = list2.Tables[0];
                    }
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        string text = Convert.ToString(dataTable.Rows[i]["Serial_Number"]);
                        int num = (int)Convert.ToInt16(dataTable.Rows[i]["Perso_Factory_RID"]);
                        DateTime dateTime2 = default(DateTime);
                        DateTime dateTime3 = default(DateTime);
                        int num2;
                        if (dataTable2.Select(string.Concat(new string[]
                        {
                            "Perso_Factory_RID='",
                            num.ToString(),
                            "' and Serial_Number='",
                            text,
                            "'"
                        })).Length != 0)
                        {
                            DataRow expr_1E5 = dataTable2.Select(string.Concat(new string[]
                            {
                                "Perso_Factory_RID='",
                                num.ToString(),
                                "' and Serial_Number='",
                                text,
                                "'"
                            }))[0];
                            num2 = Convert.ToInt32(expr_1E5["Number"].ToString());
                            dateTime3 = Convert.ToDateTime(expr_1E5["Stock_Date"].ToString());
                            dateTime2 = this.WorkDateBL.GetNextWorkDate(dateTime3);
                        }
                        else
                        {
                            num2 = 0;
                            dateTime3 = lastWorkDate;
                            dateTime2 = SurplusDate;
                        }
                        DataTable dataTable3 = new DataTable();
                        dataTable3 = this.StocksMoveIn(num, text, dateTime3, dateTime);
                        DataTable dataTable4 = new DataTable();
                        dataTable4 = this.StocksMoveOut(num, text, dateTime3, dateTime);
                        DataTable dataTable5 = new DataTable();
                        dataTable5 = this.StocksTransatrion(num, text, "1", dateTime3, dateTime);
                        DataTable dataTable6 = new DataTable();
                        dataTable6 = this.StocksTransatrion(num, text, "2", dateTime3, dateTime);
                        DataTable dataTable7 = new DataTable();
                        dataTable7 = this.StocksTransatrion(num, text, "3", dateTime3, dateTime);
                        this.dirValues.Clear();
                        this.dirValues.Add("lastSurplusDateTime", dateTime3.ToString("yyyy/MM/dd 23:59:59"));
                        this.dirValues.Add("Perso_Factory_RID", num.ToString());
                        this.dirValues.Add("Serial_Number", text.ToString());
                        DataTable dataTable8 = new DataTable();
                        DataSet list3 = base.dao.GetList("SELECT Stock_Date,isnull(sum(Number),0)  as Number FROM MATERIEL_STOCKS_USED WHERE RST = 'A' AND Stock_Date > @lastSurplusDateTime and Perso_Factory_RID=@Perso_Factory_RID and Serial_Number=@Serial_Number Group By Stock_Date Order By Stock_Date ", this.dirValues);
                        if (list3 != null && list3.Tables.Count > 0 && list3.Tables[0].Rows.Count > 0)
                        {
                            dataTable8 = list3.Tables[0];
                            this.dirValues.Clear();
                            this.dirValues.Add("lastSurplusDateTime", dateTime3.ToString("yyyy/MM/dd 23:59:59"));
                            this.dirValues.Add("Perso_Factory_RID", num.ToString());
                            this.dirValues.Add("Serial_Number", text.ToString());
                            base.dao.ExecuteNonQuery("DELETE  FROM MATERIEL_STOCKS  WHERE RST = 'A' and Perso_Factory_RID=@Perso_Factory_RID and Serial_Number=@Serial_Number  and Stock_Date > @lastSurplusDateTime ", this.dirValues);
                            while (dateTime2 <= dateTime)
                            {
                                if (this.WorkDateBL.CheckWorkDate(dateTime2))
                                {
                                    if (dataTable8.Select(" Stock_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'").Length != 0)
                                    {
                                        DataRow dataRow = dataTable8.Select(" Stock_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'")[0];
                                        num2 -= Convert.ToInt32(dataRow["Number"].ToString());
                                    }
                                    if (dataTable3 != null && dataTable3.Rows.Count > 0 && dataTable3.Select(" Move_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'").Length != 0)
                                    {
                                        DataRow dataRow2 = dataTable3.Select(" Move_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'")[0];
                                        num2 += Convert.ToInt32(dataRow2["Move_Number"].ToString());
                                    }
                                    if (dataTable4 != null && dataTable4.Rows.Count > 0 && dataTable4.Select(" Move_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'").Length != 0)
                                    {
                                        DataRow dataRow3 = dataTable4.Select(" Move_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'")[0];
                                        num2 -= Convert.ToInt32(dataRow3["Move_Number"].ToString());
                                    }
                                    if (dataTable5 != null && dataTable5.Rows.Count > 0 && dataTable5.Select(" Transaction_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'").Length != 0)
                                    {
                                        DataRow dataRow4 = dataTable5.Select(" Transaction_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'")[0];
                                        num2 += Convert.ToInt32(dataRow4["Transaction_Amount"].ToString());
                                        MATERIEL_STOCKS_MANAGE mATERIEL_STOCKS_MANAGE = new MATERIEL_STOCKS_MANAGE();
                                        mATERIEL_STOCKS_MANAGE.Stock_Date = DateTime.Parse(dateTime2.ToString("yyyy/MM/dd 00:00:00"));
                                        mATERIEL_STOCKS_MANAGE.Number = (long)Convert.ToInt32(dataRow4["Transaction_Amount"].ToString());
                                        mATERIEL_STOCKS_MANAGE.Perso_Factory_RID = num;
                                        mATERIEL_STOCKS_MANAGE.Type = "1";
                                        mATERIEL_STOCKS_MANAGE.Comment = "";
                                        mATERIEL_STOCKS_MANAGE.Serial_Number = text;
                                        mATERIEL_STOCKS_MANAGE.Invenroty_Remain = (long)num2;
                                        mATERIEL_STOCKS_MANAGE.Real_Wear_Rate = 0;
                                        base.dao.Add<MATERIEL_STOCKS_MANAGE>(mATERIEL_STOCKS_MANAGE, "RID");
                                    }
                                    if (dataTable6 != null && dataTable6.Rows.Count > 0 && dataTable6.Select(" Transaction_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'").Length != 0)
                                    {
                                        DataRow dataRow5 = dataTable6.Select(" Transaction_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'")[0];
                                        num2 -= Convert.ToInt32(dataRow5["Transaction_Amount"].ToString());
                                        MATERIEL_STOCKS_MANAGE mATERIEL_STOCKS_MANAGE2 = new MATERIEL_STOCKS_MANAGE();
                                        mATERIEL_STOCKS_MANAGE2.Stock_Date = DateTime.Parse(dateTime2.ToString("yyyy/MM/dd 00:00:00"));
                                        mATERIEL_STOCKS_MANAGE2.Number = (long)Convert.ToInt32(dataRow5["Transaction_Amount"].ToString());
                                        mATERIEL_STOCKS_MANAGE2.Perso_Factory_RID = num;
                                        mATERIEL_STOCKS_MANAGE2.Type = "2";
                                        mATERIEL_STOCKS_MANAGE2.Comment = "";
                                        mATERIEL_STOCKS_MANAGE2.Serial_Number = text;
                                        mATERIEL_STOCKS_MANAGE2.Invenroty_Remain = (long)num2;
                                        mATERIEL_STOCKS_MANAGE2.Real_Wear_Rate = 0;
                                        base.dao.Add<MATERIEL_STOCKS_MANAGE>(mATERIEL_STOCKS_MANAGE2, "RID");
                                    }
                                    if (dataTable7 != null && dataTable7.Rows.Count > 0 && dataTable7.Select(" Transaction_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'").Length != 0)
                                    {
                                        DataRow dataRow6 = dataTable7.Select(" Transaction_Date='" + dateTime2.ToString("yyyy/MM/dd 00:00:00") + "'")[0];
                                        num2 -= Convert.ToInt32(dataRow6["Transaction_Amount"].ToString());
                                        MATERIEL_STOCKS_MANAGE mATERIEL_STOCKS_MANAGE3 = new MATERIEL_STOCKS_MANAGE();
                                        mATERIEL_STOCKS_MANAGE3.Stock_Date = DateTime.Parse(dateTime2.ToString("yyyy/MM/dd 00:00:00"));
                                        mATERIEL_STOCKS_MANAGE3.Number = (long)Convert.ToInt32(dataRow6["Transaction_Amount"].ToString());
                                        mATERIEL_STOCKS_MANAGE3.Perso_Factory_RID = num;
                                        mATERIEL_STOCKS_MANAGE3.Type = "3";
                                        mATERIEL_STOCKS_MANAGE3.Comment = "";
                                        mATERIEL_STOCKS_MANAGE3.Serial_Number = text;
                                        mATERIEL_STOCKS_MANAGE3.Invenroty_Remain = (long)num2;
                                        mATERIEL_STOCKS_MANAGE3.Real_Wear_Rate = 0;
                                        base.dao.Add<MATERIEL_STOCKS_MANAGE>(mATERIEL_STOCKS_MANAGE3, "RID");
                                    }
                                    MATERIEL_STOCKS mATERIEL_STOCKS = new MATERIEL_STOCKS();
                                    mATERIEL_STOCKS.Number = (long)num2;
                                    mATERIEL_STOCKS.Perso_Factory_RID = num;
                                    mATERIEL_STOCKS.Serial_Number = text;
                                    mATERIEL_STOCKS.Stock_Date = dateTime2;
                                    base.dao.Add<MATERIEL_STOCKS>(mATERIEL_STOCKS, "RID");
                                }
                                dateTime2 = dateTime2.AddDays(1.0);
                            }
                        }
                    }
                }
            }
            catch (Exception expr_8DB)
            {
                LogFactory.Write(expr_8DB.ToString(), "ErrLog");
                throw new Exception(expr_8DB.Message);
            }
        }

        public DataTable StocksMoveOut(int Factory_RID, string Serial_Number, DateTime lastSurplusDateTime, DateTime thisSurplusDateTime)
        {
            DataTable dataTable = null;
            DataTable result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                DataSet list = base.dao.GetList("SELECT Move_Date,Move_Number FROM MATERIEL_STOCKS_MOVE WHERE RST = 'A' AND From_Factory_RID = @perso_factory_rid AND Serial_Number = @Serial_Number AND Move_Date > @lastSurplusDateTime AND Move_Date <= @thisSurplusDateTime ", this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = list.Tables[0];
                }
                result = dataTable;
            }
            catch (Exception expr_D2)
            {
                LogFactory.Write(expr_D2.ToString(), "ErrLog");
                throw expr_D2;
            }
            return result;
        }

        public DataTable StocksMoveIn(int Factory_RID, string Serial_Number, DateTime lastSurplusDateTime, DateTime thisSurplusDateTime)
        {
            DataTable dataTable = null;
            DataTable result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                DataSet list = base.dao.GetList("SELECT Move_Date,Move_Number FROM MATERIEL_STOCKS_MOVE WHERE RST = 'A' AND To_Factory_RID = @perso_factory_rid AND Serial_Number = @Serial_Number AND Move_Date > @lastSurplusDateTime AND Move_Date <= @thisSurplusDateTime ", this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = list.Tables[0];
                }
                result = dataTable;
            }
            catch (Exception expr_D2)
            {
                LogFactory.Write(expr_D2.ToString(), "ErrLog");
                throw expr_D2;
            }
            return result;
        }

        public List<string> SaveMaterielTransactionCount(DateTime Date, List<string> lstSerielNumber)
        {
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("Transaction_Date", Date.ToString("yyyy/MM/dd 00:00:00"));
                DataSet list = base.dao.GetList("SELECT Serial_Number FROM MATERIEL_STOCKS_TRANSACTION WHERE Transaction_Date = @Transaction_Date", this.dirValues);
                for (int i = 0; i < list.Tables.Count; i++)
                {
                    if (list != null && list.Tables.Count > 0 && list.Tables[i].Rows.Count > 0)
                    {
                        DataTable dataTable = list.Tables[i];
                        if (dataTable != null && dataTable.Rows.Count > 0)
                        {
                            for (int j = 0; j < dataTable.Rows.Count; j++)
                            {
                                if (dataTable.Rows[j]["Serial_Number"].ToString() != null && dataTable.Rows[j]["Serial_Number"].ToString() != "" && -1 == lstSerielNumber.IndexOf(dataTable.Rows[j]["Serial_Number"].ToString()))
                                {
                                    lstSerielNumber.Add(dataTable.Rows[j]["Serial_Number"].ToString());
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception expr_15D)
            {
                LogFactory.Write(expr_15D.ToString(), "ErrLog");
                throw new Exception(expr_15D.Message);
            }
            return lstSerielNumber;
        }

        public DataTable StocksTransatrion(int Factory_RID, string Serial_Number, string Param_Code, DateTime lastSurplusDateTime, DateTime thisSurplusDateTime)
        {
            DataTable dataTable = null;
            DataTable result;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("Param_Code", Param_Code);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                DataSet list = base.dao.GetList("SELECT MST.Transaction_Date,MST.Transaction_Amount \r\n                                    FROM MATERIEL_STOCKS_TRANSACTION MST\r\n                                    inner join PARAM P ON MST.PARAM_RID = P.RID \r\n                                    WHERE MST.RST = 'A' \r\n                                    AND P.RST = 'A' \r\n                                    AND MST.Factory_RID = @perso_factory_rid \r\n                                    AND MST.Serial_Number = @Serial_Number \r\n                                    AND MST.Transaction_Date > @lastSurplusDateTime \r\n                                    AND MST.Transaction_Date <= @thisSurplusDateTime \r\n                                    AND P.ParamType_Code = 'matType1' \r\n                                    AND P.Param_Code = @Param_Code", this.dirValues);
                if (list != null && list.Tables.Count > 0 && list.Tables[0].Rows.Count > 0)
                {
                    dataTable = list.Tables[0];
                }
                result = dataTable;
            }
            catch (Exception expr_E3)
            {
                LogFactory.Write(expr_E3.ToString(), "ErrLog");
                throw expr_E3;
            }
            return result;
        }
    }
}
