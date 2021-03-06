//*****************************************
//*  作    者：GaoAi
//*  功能說明：物料庫存控管
//*  創建日期：2008-12-10
//*  修改日期：
//*  修改記錄：
//*****************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
using System.Collections;
using System.Data.Common;
using Microsoft.Practices.EnterpriseLibrary.Data;
using System.Text.RegularExpressions;
using System.IO;
using CIMSBatch.Model;
using CIMSBatch.FTP;
using CIMSBatch.Mail;
using CIMSBatch.Public;
using CIMSClass.Business;

namespace CIMSBatch.Business
{
    class Depository010BL : BaseLogic
    {
        #region SQL語句
        public const string SEL_FACTORY_ALL = "SELECT RID,Factory_ShortName_CN "
                                        + "FROM FACTORY "
                                        + "WHERE RST = 'A' AND Is_Perso = 'Y' "
                                        + "ORDER BY RID";
        public const string SEL_MATERIEL_LAST_SURPLUS_DATE = "SELECT TOP 1 Stock_Date,Number "
                                        + "FROM (SELECT MSM.Serial_Number,Stock_Date,MSM.Number "
                                        + "FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number "
                                        + "LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number "
                                        + "LEFT JOIN DMTYPE_INFO DM ON DM.RST = 'A' AND MSM.Serial_Number = DM.Serial_Number "
                                        + "WHERE MSM.RST = 'A' AND MSM.Perso_Factory_RID = @perso_factory_rid AND Type = '4') A "
                                        + "WHERE Serial_Number = @serial_number "
                                        + "ORDER BY Stock_Date DESC";







        public const string SEL_MATERIEL_SURPLUS_DATE = "SELECT Stock_Date "
                                        + "FROM (SELECT MS.Serial_Number,Stock_Date "
                                        + "FROM MATERIEL_STOCKS MS LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MS.Serial_Number = CE.Serial_Number "
                                        + "LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MS.Serial_Number = EI.Serial_Number "
                                        + "LEFT JOIN DMTYPE_INFO DM ON DM.RST = 'A' AND MS.Serial_Number = DM.Serial_Number "
                                        + "WHERE MS.RST = 'A' AND MS.Perso_Factory_RID = @perso_factory_rid) A "
                                        + "WHERE Serial_Number = @serial_number "
                                        + "ORDER BY Stock_Date DESC";
        public const string SEL_ENVELOPE_INFO = "SELECT * "
                                        + "FROM ENVELOPE_INFO "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_CARD_EXPONENT = "SELECT * "
                                        + "FROM CARD_EXPONENT "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_DMTYPE_INFO = "SELECT * "
                                        + "FROM DMTYPE_INFO "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_MATERIEL_STOCKS_MANAGE = "SELECT TOP 1 * "
                                        + "FROM MATERIEL_STOCKS_MANAGE "
                                        + "WHERE RST = 'A' AND TYPE = '4' AND Perso_Factory_RID = @perso_factory_rid "
                                        + " AND Materiel_RID = @materiel_rid AND Materiel_Type = @materiel_type "
                                        + "ORDER BY Stock_Date DESC";
        public const string SEL_MATERIEL_STOCKS_MOVE_IN = "SELECT Move_Date,Move_Number "
                                        + "FROM MATERIEL_STOCKS_MOVE "
                                        + "WHERE RST = 'A' AND To_Factory_RID = @perso_factory_rid "
                                        + "AND Serial_Number = @Serial_Number "
                                        + "AND Move_Date > @lastSurplusDateTime "
                                        + "AND Move_Date <= @thisSurplusDateTime ";
        public const string SEL_MATERIEL_STOCKS_MOVE_OUT = "SELECT Move_Date,Move_Number "
                                        + "FROM MATERIEL_STOCKS_MOVE "
                                        + "WHERE RST = 'A' AND From_Factory_RID = @perso_factory_rid "
                                        + "AND Serial_Number = @Serial_Number "
                                        + "AND Move_Date > @lastSurplusDateTime "
                                        + "AND Move_Date <= @thisSurplusDateTime ";
        public const string SEL_CARD_INFO_ENVELOPE = "SELECT CT.RID "
                                        + "FROM Card_Type CT "
                                        + "WHERE CT.RST = 'A' AND CT.Envelope_RID = @envelope_rid ";
        public const string SEL_CARD_INFO_EXPONENT = "SELECT CT.RID "
                                        + "FROM Card_Type CT "
                                        + "WHERE CT.RST = 'A' AND CT.Exponent_RID = @exponent_rid ";
        public const string SEL_CARD_INFO_DM = "SELECT CT.RID "
                                        + "FROM DM_CARDTYPE DC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DC.CardType_RID = CT.RID "
                                        + "WHERE DC.RST = 'A' AND DC.DM_RID = @dm_rid ";
        public const string SEL_SUBTOTAL_IMPORT = "SELECT SI.PHOTO,SI.Number,SI.TYPE,SI.Affinity,SI.Date_Time,CT.RID "
                                        + "FROM SUBTOTAL_IMPORT AS SI "
                                        + "left join card_type as CT "
                                        + "on ct.rst='A' and ct.photo=si.photo and ct.type=si.type and ct.affinity=si.affinity "
                                        + "WHERE SI.RST = 'A' AND SI.Perso_Factory_RID = @perso_factory_rid "
                                        + "AND SI.Date_Time > @lastSurplusDateTime "
                                        + "AND SI.Date_Time <= @thisSurplusDateTime ";

        public const string GET_ENVELOPE_INFO_WEARRATE = "SELECT EI.Wear_Rate "
                                        + "FROM CARD_TYPE CT INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID "
                                        + "WHERE CT.RST = 'A' AND CT.TYPE = @total_type AND CT.PHOTO = @photo AND CT.Affinity = @Affinity ";
        public const string GET_CARD_EXPONENT_WEARRATE = "SELECT CE.Wear_Rate "
                                        + "FROM CARD_TYPE CT INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID "
                                        + "WHERE CT.RST = 'A' AND CT.TYPE = @total_type AND CT.PHOTO = @photo AND CT.Affinity = @Affinity ";
        public const string GET_DMTYPE_INFO_WEARRATE = "SELECT DM.Wear_Rate "
                                        + "FROM DM_CARDTYPE DC INNER JOIN CARD_TYPE CT ON DC.RST = 'A' AND CT.RST = 'A' AND DC.CardType_RID = CT.RID "
                                        + "INNER JOIN DMTYPE_INFO DM ON DM.RST = 'A' AND DC.DM_RID = DM.RID "
                                        + "WHERE CT.TYPE = @total_type AND CT.PHOTO = @photo AND CT.Affinity = @Affinity ";
        public const string GET_LAST_SURPLUS_BY_FACTORY = "SELECT TOP 1 * "
                                        + "FROM MATERIEL_STOCKS_MANAGE "
                                        + "WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Type = 4 "
                                        + "ORDER BY Stock_Date DESC";
        public const string GET_SURPLUS_BY_FACTORY = "SELECT distinct Serial_Number "
                                        + "FROM MATERIEL_STOCKS_MANAGE "
                                        + "WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND Type = 4 ";
        public const string DEL_MATERIEL_STOCKS_MANAGE = "DELETE "
                                        + "FROM MATERIEL_STOCKS_MANAGE "
                                        + "WHERE RST = 'A' AND RID = @rid ";
        public const string DEL_MATERIEL_STOCKS_SYS = "DELETE "
                                            + "FROM MATERIEL_STOCKS "
                                            + "WHERE RST = 'A' "
                                            + "AND Perso_Factory_RID = @perso_factory_rid "
                                            + "AND Serial_Number = @Serial_Number "
                                            + "AND Stock_Date > @lastSurplusDateTime ";
        //+ "AND Stock_Date <= @thisSurplusDateTime ";
        public const string DEL_MATERIEL_STOCKS_MANAGE_DEL = "DELETE "
                                            + "FROM MATERIEL_STOCKS_MANAGE "
                                            + "WHERE RST = 'A' "
                                            + "AND Perso_Factory_RID = @perso_factory_rid "
                                            + "AND Serial_Number = @Serial_Number "
                                            + "AND Stock_Date > @lastSurplusDateTime ";
        public const string DEL_MATERIEL_STOCKS_DEL = "DELETE "
                                            + "FROM MATERIEL_STOCKS "
                                            + "WHERE RST = 'A' "
                                            + "AND Perso_Factory_RID = @perso_factory_rid "
                                            + "AND Serial_Number = @Serial_Number "
                                            + "AND Stock_Date > @lastSurplusDateTime ";
        public const string SEL_MATERIEL_STOCKS_USED = "select * from MATERIEL_STOCKS_USED where rst='A' "
                                            + "AND Serial_Number=@Serial_Number "
                                            + "AND Perso_Factory_RID=@Perso_Factory_RID "
                                            + "AND Stock_Date > @lastSurplusDateTime "
                                            + "AND Stock_Date <= @thisSurplusDateTime";
        public const string CON_MATERIEL_STOCKS = "select * from IMPORT_HISTORY where File_Type='4' and File_Name=@File_Name";

        public const string SEL_ALL_WORK_DATE = "SELECT Date_Time FROM WORK_DATE WHERE RST = 'A' AND Is_WorkDay = 'Y' Order by Date_Time";

        public const string SEL_MATERIAL_USED_ENVELOPE = "SELECT SI.Date_Time AS Stock_Date,EI.Serial_Number,SI.Number "
                        + "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                        + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                        + "INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID "
                        + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time AND EI.Serial_Number = @Serial_Number";
        public const string SEL_MATERIAL_USED_CARD_EXPONENT = "SELECT SI.Date_Time AS Stock_Date,CE.Serial_Number,SI.Number "
                        + "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                        + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                        + "INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID "
                        + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time AND CE.Serial_Number = @Serial_Number";
        //public const string SEL_MATERIAL_USED_DM = "SELECT SI.Date_Time AS Stock_Date,DI.Serial_Number,SI.Number "
        //                + "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
        //                + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
        //                + "INNER JOIN DM_CARDTYPE DCT ON DCT.RST = 'A' AND CT.RID = DCT.CardType_RID "
        //                + "INNER JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND DCT.DM_RID = DI.RID "
        //                + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time AND DI.Serial_Number = @Serial_Number";

        public const string SEL_MATERIAL_USED_DM = "SELECT A.Date_Time as Stock_Date, DI.Serial_Number, A.Number1 AS Number "
                + " FROM (SELECT CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID "
                + " FROM  SUBTOTAL_IMPORT  SI "
                + " INNER JOIN CARD_TYPE  CT "
                + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                + " WHERE (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time) AND (SI.Date_Time <= @End_Date_Time) AND   "
                + " (SI.Perso_Factory_RID = @Perso_Factory_RID)) A "
                + " INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID "
                + " INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID  "
                + " WHERE (DI.Card_Type_Link_Type = '1') AND (DI.Serial_Number = @Serial_Number)  "
                + " UNION  "
                + " SELECT A_1.Date_Time as Stock_Date, DI.Serial_Number, A_1.Number1 AS Number  "
                + " FROM (SELECT CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID  "
                + " FROM   SUBTOTAL_IMPORT SI "
                + " INNER JOIN CARD_TYPE  CT "
                + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                + " WHERE  (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time) "
                + " AND (SI.Date_Time <= @End_Date_Time) AND (SI.Perso_Factory_RID = @Perso_Factory_RID)) A_1 "
                + " INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID "
                + " INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID "
                + " INNER JOIN DM_CARDTYPE  DCT "
                + " ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID "
                + " WHERE (DI.Card_Type_Link_Type = '2') AND (DI.Serial_Number = @Serial_Number)";

        public const string SEL_MATERIAL_USED_ENVELOPE_S = "SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,EI.Serial_Number,SI.Number "
          + "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
          + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
          + "INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID "
            //  + "WHERE SI.RST = 'A'  AND Date_Time>=@From_Date_Time AND Date_Time<=@End_Date_Time ";
           + "WHERE SI.RST = 'A'  AND  Date_Time=@End_Date_Time ";

        public const string SEL_MATERIAL_USED_CARD_EXPONENT_S = "SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,CE.Serial_Number,SI.Number "
                + "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                + "INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID "
            // + "WHERE SI.RST = 'A'  AND Date_Time>=@From_Date_Time AND Date_Time<=@End_Date_Time ";
             + "WHERE SI.RST = 'A'  AND  Date_Time=@End_Date_Time ";

        public const string SEL_MATERIAL_USED_DM_S = "SELECT A.Perso_Factory_RID,A.Date_Time as Stock_Date, DI.Serial_Number, A.Number1 AS Number "
                           + " FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID "
                           + " FROM  SUBTOTAL_IMPORT  SI "
                           + " INNER JOIN CARD_TYPE  CT "
                           + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                           + " WHERE (SI.RST = 'A') AND  (SI.Date_Time = @End_Date_Time)   "
                           + " ) A "
                           + " INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID "
                           + " INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID  "
                           + " WHERE (DI.Card_Type_Link_Type = '1')  "
                           + " UNION  "
                           + " SELECT A_1.Perso_Factory_RID,A_1.Date_Time as Stock_Date, DI.Serial_Number, A_1.Number1 AS Number  "
                           + " FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID  "
                           + " FROM   SUBTOTAL_IMPORT SI "
                           + " INNER JOIN CARD_TYPE  CT "
                           + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                           + " WHERE  (SI.RST = 'A')  "
                           + " AND (SI.Date_Time = @End_Date_Time) ) A_1 "
                           + " INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID "
                           + " INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID "
                           + " INNER JOIN DM_CARDTYPE  DCT "
                           + " ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID "
                           + " WHERE (DI.Card_Type_Link_Type = '2') ";

        #region 200908CR物料的消耗計算改為用小計檔的「替換前」版面計算
        //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 start
        public const string SEL_MATERIAL_USED_ENVELOPE_REPLACE = "SELECT SI.Date_Time AS Stock_Date,EI.Serial_Number,SI.Number "
               + "FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
               + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
               + "INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID "
               + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time AND EI.Serial_Number = @Serial_Number";
        //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 end
        public const string SEL_MATERIAL_USED_CARD_EXPONENT_REPLACE = "SELECT SI.Date_Time AS Stock_Date,CE.Serial_Number,SI.Number "
                + "FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                + "INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID "
                + "WHERE SI.RST = 'A' AND Perso_Factory_RID = @Perso_Factory_RID AND Date_Time>@From_Date_Time AND Date_Time<=@End_Date_Time AND CE.Serial_Number = @Serial_Number";


        public const string SEL_MATERIAL_USED_DM_REPLACE = "SELECT A.Date_Time as Stock_Date, DI.Serial_Number, A.Number1 AS Number "
                        + " FROM (SELECT CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID "
                        + " FROM  SUBTOTAL_REPLACE_IMPORT  SI "
                        + " INNER JOIN CARD_TYPE  CT "
                        + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                        + " WHERE (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time) AND (SI.Date_Time <= @End_Date_Time) AND   "
                        + " (SI.Perso_Factory_RID = @Perso_Factory_RID)) A "
                        + " INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID "
                        + " INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID  "
                        + " WHERE (DI.Card_Type_Link_Type = '1') AND (DI.Serial_Number = @Serial_Number)  "
                        + " UNION  "
                        + " SELECT A_1.Date_Time as Stock_Date, DI.Serial_Number, A_1.Number1 AS Number  "
                        + " FROM (SELECT CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID  "
                        + " FROM   SUBTOTAL_REPLACE_IMPORT SI "
                        + " INNER JOIN CARD_TYPE  CT "
                        + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                        + " WHERE  (SI.RST = 'A') AND (SI.Date_Time > @From_Date_Time) "
                        + " AND (SI.Date_Time <= @End_Date_Time) AND (SI.Perso_Factory_RID = @Perso_Factory_RID)) A_1 "
                        + " INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID "
                        + " INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID "
                        + " INNER JOIN DM_CARDTYPE  DCT "
                        + " ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID "
                        + " WHERE (DI.Card_Type_Link_Type = '2') AND (DI.Serial_Number = @Serial_Number)";

        public const string SEL_MATERIAL_USED_ENVELOPE_S_REPLACE = "SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,EI.Serial_Number,SI.Number "
              + "FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
              + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
              + "INNER JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT.Envelope_RID = EI.RID "
               + "WHERE SI.RST = 'A'  AND  Date_Time=@End_Date_Time ";

        public const string SEL_MATERIAL_USED_CARD_EXPONENT_S_REPLACE = "SELECT SI.Perso_Factory_RID,SI.Date_Time AS Stock_Date,CE.Serial_Number,SI.Number "
                + "FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST ='A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                + "INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND SI.MakeCardType_RID = MCT.RID AND MCT.Type_Name In ('3D','DA','PM','RN') "
                + "INNER JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT.Exponent_RID = CE.RID "
             + "WHERE SI.RST = 'A'  AND  Date_Time=@End_Date_Time ";

        public const string SEL_MATERIAL_USED_DM_S_REPLACE = "SELECT A.Perso_Factory_RID,A.Date_Time as Stock_Date, DI.Serial_Number, A.Number1 AS Number "
                           + " FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID "
                           + " FROM  SUBTOTAL_REPLACE_IMPORT  SI "
                           + " INNER JOIN CARD_TYPE  CT "
                           + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                           + " WHERE (SI.RST = 'A') AND  (SI.Date_Time = @End_Date_Time)   "
                           + " ) A "
                           + " INNER JOIN DM_MAKECARDTYPE  DMM ON DMM.MakeCardType_RID = A.MakeCardType_RID "
                           + " INNER JOIN DMTYPE_INFO DI ON DI.RID = DMM.DM_RID  "
                           + " WHERE (DI.Card_Type_Link_Type = '1')  "
                           + " UNION  "
                           + " SELECT A_1.Perso_Factory_RID,A_1.Date_Time as Stock_Date, DI.Serial_Number, A_1.Number1 AS Number  "
                           + " FROM (SELECT SI.Perso_Factory_RID,CT.RID, SI.Number AS Number1, SI.Date_Time, SI.MakeCardType_RID  "
                           + " FROM   SUBTOTAL_REPLACE_IMPORT SI "
                           + " INNER JOIN CARD_TYPE  CT "
                           + " ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO  "
                           + " WHERE  (SI.RST = 'A')  "
                           + " AND (SI.Date_Time = @End_Date_Time) ) A_1 "
                           + " INNER JOIN DM_MAKECARDTYPE DMM ON DMM.MakeCardType_RID = A_1.MakeCardType_RID "
                           + " INNER JOIN DMTYPE_INFO  DI ON DI.RID = DMM.DM_RID "
                           + " INNER JOIN DM_CARDTYPE  DCT "
                           + " ON DCT.RST = 'A' AND A_1.RID = DCT.CardType_RID AND DCT.DM_RID = DI.RID "
                           + " WHERE (DI.Card_Type_Link_Type = '2') ";
        #endregion

        //add by Ian Huang start
        public const string SEL_MATERIEL_STOCKS_TRANSACTION = @"SELECT MST.Transaction_Date,MST.Transaction_Amount,P.Param_Code 
                        FROM MATERIEL_STOCKS_TRANSACTION MST
                        inner join PARAM P on MST.PARAM_RID = P.RID
                        WHERE MST.RST = 'A' AND MST.Factory_RID = @perso_factory_rid 
                        AND MST.Serial_Number = @Serial_Number 
                        AND MST.Transaction_Date > @lastSurplusDateTime 
                        AND MST.Transaction_Date <= @thisSurplusDateTime ";

        public const string DELETE_MATERIEL_STOCKS_MANAGE = @"DELETE FROM MATERIEL_STOCKS_MANAGE 
                        where Stock_Date = @Stock_Date and RCU = @RCU and Perso_Factory_RID = @Perso_Factory_RID 
                        and Type = @Type and Serial_Number = @Serial_Number";

        public const string SEL_MATERIEL_STOCKS = @"select Number from MATERIEL_STOCKS
                        where Stock_date = @Stock_date and Serial_Number = @Serial_Number and Perso_Factory_RID = @Perso_Factory_RID";

        public const string SEL_MATERIEL_STOCKS_MOVE = @"select * from MATERIEL_STOCKS_MOVE
                        where Move_Date = @Move_Date and Serial_Number = @Serial_Number and Move_Number = @Move_Number ";

        public const string UPDATE_MATERIEL_STOCKS_MOVE_1 = @"update MATERIEL_STOCKS_MOVE set From_Factory_RID = 0
                        where From_Factory_RID = @Factory_RID and To_Factory_RID <> 0 and RCU = 'INPUT'";

        public const string UPDATE_MATERIEL_STOCKS_MOVE_2 = @"update MATERIEL_STOCKS_MOVE set To_Factory_RID = 0
                        where To_Factory_RID = @Factory_RID and From_Factory_RID <> 0 and RCU = 'INPUT'";

        public const string DEL_MATERIEL_STOCKS_MOVE_1 = @"delete from MATERIEL_STOCKS_MOVE
                        where From_Factory_RID = @Factory_RID and To_Factory_RID = 0 and RCU = 'INPUT'";

        public const string DEL_MATERIEL_STOCKS_MOVE_2 = @"delete from MATERIEL_STOCKS_MOVE
                        where To_Factory_RID = @Factory_RID and From_Factory_RID = 0 and RCU = 'INPUT'";

        public const string INSERT_MATERIEL_STOCKS_MOVE = @"INSERT INTO MATERIEL_STOCKS_MOVE
           (Move_Date,RCU,RUU,RCT,RUT,RST,Move_Number,From_Factory_RID,To_Factory_RID,Move_ID,Serial_Number) VALUES
           (@Move_Date,@RCU,@RUU,@RCT,@RUT,@RST,@Move_Number,@From_Factory_RID,@To_Factory_RID,@Move_ID,@Serial_Number)";

        public const string SEL_MAX_MOVE_ID = "SELECT TOP 1 Move_ID "
                                + "FROM MATERIEL_STOCKS_MOVE "
                                + "WHERE Move_Date >= @move_date1 AND Move_Date<=@move_date2 "
                                + "ORDER BY Move_ID DESC ";
        //add by Ian Huang end
        #endregion

        //參數
        public string strErr;
        Dictionary<string, object> dirValues = new Dictionary<string, object>();
        //下載檔案
        public ArrayList Download(string FactoryPath)
        {
            ArrayList FileNameList = new ArrayList();
            try  // 加try catch add by judy 2018/03/28
            {
                #region Attribute

                string FolderYear = DateTime.Now.ToString("yyyy");
                string FolderDate = DateTime.Now.ToString("MMdd");
                string FolderName = "";
                FTPFactory ftp = new FTPFactory(GlobalString.FtpString.MATERIAL);
                string ftpPath = ConfigurationManager.AppSettings["FTPMATERIEL"] + "/" + FactoryPath; ;
                string locPath = ConfigurationManager.AppSettings["LocalMATERIEL"];
                string[] fileList;
                string[] fileMethod;
                bool returnFlag;


                #endregion
                fileList = ftp.GetFileList(ftpPath);
                if (fileList != null)
                {
                    foreach (string FileName in fileList)
                    {
                        if (!CheckFile(FileName)) //檢查要下載的檔案規則，不滿足則跳過
                        {
                            continue;
                        }
                        fileMethod = FileName.Split('-');
                        if (fileMethod != null)
                        {
                            FolderName = locPath + "\\" + fileMethod[1] + "\\";

                            returnFlag = ftp.Download(ftpPath, FileName, FolderName, FileName);
                            if (returnFlag)
                            {
                                string[] FList = new string[2];
                                FList[0] = FolderName;
                                FList[1] = FileName;
                                FileNameList.Add(FList);
                                // Legend 2017/11/28 將此處刪檔注釋做UAT測試, 上線是再解開 todo
                                returnFlag = ftp.Delete(ftpPath, FileName);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogFactory.Write(" 匯入廠商物料異動下載FTP檔案Download方法報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
           
            return FileNameList;
        }
        /// <summary>
        /// 檢查FTP檔案規則
        /// </summary>
        /// <param name="FileName">檔案名稱(全名)</param>
        /// <returns></returns>
        private bool CheckFile(string FileName)
        {
            string[] fileSplit = FileName.Split('-');
            if (fileSplit.Length != 3)
            {
                return false;
            }
            if (fileSplit[0].Trim().ToLower() != "material")
            {
                return false;
            }
            if (fileSplit[1].Trim().Length != 8)
            {
                return false;
            }
            try
            {
                DateTime time = DateTime.ParseExact(fileSplit[1].Trim(), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                ;
            }
            catch (Exception ex)
            {
                return false;
            }

            BatchBL bl = new BatchBL();
            if (!(bl.CheckWorkDate(Convert.ToDateTime(fileSplit[1].Substring(0, 4) + "/" + fileSplit[1].Substring(4, 2) + "/" + fileSplit[1].Substring(6, 2))))) //非工作日直接返回，不執行批次
            {
                return false;
            }

            if (CheckImportFile(FileName))
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 檢查是否已匯入過
        /// </summary>
        /// <param name="ImportFileName"></param>
        /// <returns>已匯入過返回true</returns>
        private bool CheckImportFile(string ImportFileName)
        {

            string[] sr = ImportFileName.Split('-');
            DataSet dst = new DataSet();
            dirValues.Clear();
            dirValues.Add("File_Name", ImportFileName);
            dst = dao.GetList(CON_MATERIEL_STOCKS, dirValues);
            if (dst.Tables[0].Rows.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        /// <summary>
        /// 以廠商英文簡稱取廠商編號
        /// </summary>
        /// <param name="Factory_ShortName_CN">廠商英文簡稱</param>
        /// <returns></returns>
        private string getFactoryIDByShortName_EN(string Factory_ShortName_EN)
        {
            FACTORY mFACTORY = null;
            try
            {
                mFACTORY = dao.GetModel<FACTORY, String>("Factory_ShortName_EN", Factory_ShortName_EN);
                return mFACTORY.Factory_ID;

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("以廠商英文簡稱取廠商編號getFactoryIDByShortName_EN報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return "";
            }
        }
        /// <summary>
        /// 以廠商英文簡稱取廠商RID
        /// </summary>
        /// <param name="Factory_ShortName_CN">廠商英文簡稱</param>
        /// <returns></returns>
        private int getFactoryRIDByShortName_EN(string Factory_ShortName_EN)
        {
            FACTORY mFACTORY = null;
            try
            {
                mFACTORY = dao.GetModel<FACTORY, String>("Factory_ShortName_EN", Factory_ShortName_EN);
                return mFACTORY.RID;

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("以廠商英文簡稱取廠商RID, getFactoryRIDByShortName_EN報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return 0;
            }
        }
        /// <summary>
        /// 以結余類型編號，取結余類型名稱
        /// </summary>
        /// <param name="strType"></param>
        /// <returns></returns>
        public string getTypeName(string strType)
        {
            string strReturn = "";
            switch (strType)
            {
                case "01":
                    strReturn = "進貨";
                    break;
                case "02":
                    strReturn = "退貨";
                    break;
                case "03":
                    strReturn = "銷毀";
                    break;
                case "04":
                    strReturn = "結余";
                    break;
                //add by Ian Huang start
                case "05":
                    strReturn = "移轉出";
                    break;
                case "06":
                    strReturn = "抽驗";
                    break;
                case "07":
                    strReturn = "退件重寄";
                    break;
                case "08":
                    strReturn = "電訊單異動";
                    break;
                case "09":
                    strReturn = "移轉入";
                    break;
                //add by Ian Huang end
                default:
                    break;
            }
            return strReturn;
        }
        /// <summary>
        /// 取Perso廠商的物料最近一次的結余日期、結余數量
        /// </summary>
        /// <param name="Serial_Number">品名編號</param>
        /// <param name="Perso_Factory_RID">Perso廠RID</param>
        /// <returns><DateTime>物料的最近一次的結余日期</returns>
        public DataTable GetLastSurplusDateNum(string Serial_Number, int Perso_Factory_RID)
        {
            DataTable dtLastSurplusDateNum = null;
            try
            {
                // 取最近一次的結余日期、結余數量
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Perso_Factory_RID);
                this.dirValues.Add("serial_number", Serial_Number);
                DataSet dstSurplus = dao.GetList(SEL_MATERIEL_LAST_SURPLUS_DATE, this.dirValues);
                if (null != dstSurplus && dstSurplus.Tables.Count > 0 &&
                    dstSurplus.Tables[0].Rows.Count > 0)
                {
                    dtLastSurplusDateNum = dstSurplus.Tables[0];
                }

                return dtLastSurplusDateNum;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取Perso廠商的物料最近一次的結余日期、結余數量, GetLastSurplusDateNum報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
        }
        /// <summary>
        /// 通過卡片庫存表的時間判斷指定日期是否為日結時間
        /// </summary>
        /// <param name="Stock_Date"></param>
        /// <returns></returns>
        private bool isSurplusDate(DateTime Stock_Date)
        {
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("Stock_Date", Stock_Date);
                DataSet ds = dao.GetList("select * from CARDTYPE_STOCKS where rst='A' and Stock_Date=@Stock_Date", dirValues);
                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("通過卡片庫存表的時間判斷指定日期是否為日結時間isSurplusDate報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }

        public string GetSubstringByByte(string strReadLine, int begin, int length, out int nextBegin)
        {
            string strTemp1 = strReadLine.Substring(begin, length + length - StringUtil.GetByteLength(strReadLine.Substring(begin, length)));
            nextBegin = begin + strTemp1.Length;
            return strTemp1;
        }

        /// <summary>
        /// 檢查是否每一種物料都有上一次結余記錄，如果沒有，則新增一條結余為0的記錄！
        /// </summary>
        /// <param name="dstlMaterielStocksIn"></param>
        public void CheckMATERIEL_STOCKS(DataSet dstlMaterielStocksIn)
        {
            try
            {
                dao.OpenConnection();
                if (dstlMaterielStocksIn.Tables.Count != 2)
                    return;

                if (dstlMaterielStocksIn.Tables[0].Rows.Count == 0)
                    return;

                string strFactory = dstlMaterielStocksIn.Tables[0].Rows[0]["factory_rid"].ToString();

                //dstlMaterielStocksIn.Tables[1].DefaultView.Sort = "Serial_Number,Stock_Date";

                foreach (DataRow dr in dstlMaterielStocksIn.Tables[1].Rows)
                {
                    DataTable dtbl = dao.GetList("select count(*) from dbo.MATERIEL_STOCKS_MANAGE where Serial_Number='" + dr["Serial_Number"].ToString() + "' and Perso_Factory_rid=" + strFactory).Tables[0];

                    if (dtbl.Rows[0][0].ToString() == "0")
                    {
                        MATERIEL_STOCKS_MANAGE msModel = new MATERIEL_STOCKS_MANAGE();
                        msModel.Number = 0;
                        msModel.Type = "4";
                        msModel.Perso_Factory_RID = int.Parse(strFactory);
                        msModel.RST = "A";
                        msModel.Serial_Number = dr["Serial_Number"].ToString();
                        msModel.Stock_Date = Convert.ToDateTime(dr["Stock_Date"]).Date.AddDays(-1);
                        dao.Add<MATERIEL_STOCKS_MANAGE>(msModel, "RID");
                    }
                }


                dao.Commit();
            }
            catch (Exception ex)
            {
                dao.Rollback();
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("檢查是否每一種物料都有上一次結余記錄，如果沒有，則新增一條結余為0的記錄！CheckMATERIEL_STOCKS報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(BizMessage.BizCommMsg.ALT_CMN_InitPageFail);
            }
            finally
            {
                dao.CloseConnection();
            }
        }

        /// <summary>
        /// 匯入前對匯入文件進行格式檢查
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public DataSet CheckIn(Dictionary<string, object> values)
        {
            StreamReader sr = null;
            DataSet dstDataIn = null;
            try
            {
                if (!File.Exists(values["FileUpd"].ToString()))
                {
                    return null;
                }

                sr = new StreamReader(values["FileUpd"].ToString(), System.Text.Encoding.Default);
                string[] strLine;
                string strReadLine = "";
                int count = 1;
                strErr = "";
                int intPerso_Factory_Rid = 0;
                // add by Ian Huang start
                string strMaxDate = "";
                // add by Ian Huang end

                DataTable dtFactory = new DataTable();
                dtFactory.Columns.Add(new DataColumn("Factory_RID", Type.GetType("System.Int32")));
                dtFactory.Columns.Add(new DataColumn("Factory_ID", Type.GetType("System.String")));
                dtFactory.Columns.Add(new DataColumn("Factory_Name", Type.GetType("System.String")));

                DataTable dtDataIn = new DataTable();
                dtDataIn.Columns.Add(new DataColumn("Serial_Number", Type.GetType("System.String")));// 物品編號
                dtDataIn.Columns.Add(new DataColumn("Materiel_Name", Type.GetType("System.String")));// 物料品名
                dtDataIn.Columns.Add(new DataColumn("Stock_Date", Type.GetType("System.DateTime")));// 異動日期
                dtDataIn.Columns.Add(new DataColumn("Type", Type.GetType("System.Int32")));//異動類型
                dtDataIn.Columns.Add(new DataColumn("Number", Type.GetType("System.Int32")));// 檔案結余數量
                dtDataIn.Columns.Add(new DataColumn("Materiel_RID", Type.GetType("System.Int32")));//物料RID
                dtDataIn.Columns.Add(new DataColumn("System_Num", Type.GetType("System.Int32")));// 庫存結余數量
                dtDataIn.Columns.Add(new DataColumn("Last_Surplus_Date", Type.GetType("System.DateTime")));// 物品的上次結余日期
                dtDataIn.Columns.Add(new DataColumn("Last_Surplus_Num", Type.GetType("System.Int32")));// 物品的上次結余數量
                dtDataIn.Columns.Add(new DataColumn("Comment", Type.GetType("System.String")));// 備注

                #region 讀字符串，并檢查字符串格式(列數、每列的字符格式),并保存到臨時DataTable(dtDataIn)中
                while ((strReadLine = sr.ReadLine()) != null)
                {
                    if (count == 1)
                    {
                        // Perso廠一致性檢查
                        if (getFactoryIDByShortName_EN(Convert.ToString(values["Factory_ShortName_EN"])) != strReadLine.Trim())
                        {
                            strErr += "Perso廠不一致\n";
                        }

                        // 保存Perso廠商訊息
                        DataRow drFactory = dtFactory.NewRow();
                        drFactory["Factory_RID"] = getFactoryRIDByShortName_EN(Convert.ToString(values["Factory_ShortName_EN"]));
                        intPerso_Factory_Rid = Convert.ToInt32(drFactory["Factory_RID"]);
                        drFactory["Factory_ID"] = strReadLine.Trim();
                        drFactory["Factory_Name"] = Convert.ToString(values["Factory_ShortName_EN"]);
                        dtFactory.Rows.Add(drFactory);
                    }
                    else
                    {
                        // 不是空的行
                        if (!StringUtil.IsEmpty(strReadLine))
                        {
                            if (StringUtil.GetByteLength(strReadLine) != 25)//列數量檢查
                            {
                                strErr += "第" + count.ToString() + "行列數不正確。\n";
                            }
                            else
                            {

                                // 分割字符串
                                //strLine = strReadLine.Split(GlobalString.FileSplit.Split);
                                int nextBegin = 0;
                                strLine = new string[4];
                                strLine[0] = GetSubstringByByte(strReadLine, nextBegin, 6, out nextBegin).Trim();
                                strLine[1] = GetSubstringByByte(strReadLine, nextBegin, 8, out nextBegin).Trim();
                                strLine[2] = GetSubstringByByte(strReadLine, nextBegin, 2, out nextBegin).Trim();
                                strLine[3] = GetSubstringByByte(strReadLine, nextBegin, 9, out nextBegin).Trim();


                                // 列長度檢查
                                for (int i = 0; i < strLine.Length; i++)
                                {
                                    int num = i + 1;
                                    if (StringUtil.IsEmpty(strLine[i]))
                                        strErr += "第" + count.ToString() + "行第" + num.ToString() + "列為空;\n";
                                    else
                                        strErr += CheckFileOneColumn(strLine[i], num, count);

                                    if (i == 0)
                                    {
                                        dirValues.Clear();
                                        dirValues.Add("Serical_Number", strLine[i]);
                                        if (strLine[i].Contains("A"))
                                        {
                                            if (!dao.Contains("select count(*) from ENVELOPE_INFO where Serial_number=@Serical_Number", dirValues))
                                                strErr += "第" + count.ToString() + "行第" + num.ToString() + "列信封不存在;";
                                        }
                                        if (strLine[i].Contains("B"))
                                        {
                                            if (!dao.Contains("select count(*) from CARD_EXPONENT where Serial_number=@Serical_Number", dirValues))
                                                strErr += "第" + count.ToString() + "行第" + num.ToString() + "列寄卡單不存在;";
                                        }
                                        if (strLine[i].Contains("C"))
                                        {
                                            if (!dao.Contains("select count(*) from DMTYPE_INFO where Serial_number=@Serical_Number", dirValues))
                                                strErr += "第" + count.ToString() + "行第" + num.ToString() + "列DM不存在;";
                                        }
                                    }

                                    if (i == 1)
                                    {
                                        DataTable dtblMax = dao.GetList("select convert(varchar(20),max(stock_date),112) from MATERIEL_STOCKS_MANAGE where Type='4' and Perso_Factory_RID = " + intPerso_Factory_Rid.ToString()).Tables[0];
                                        if (dtblMax.Rows.Count > 0)
                                        {
                                            if (!StringUtil.IsEmpty(dtblMax.Rows[0][0].ToString()))
                                            {
                                                if (int.Parse(dtblMax.Rows[0][0].ToString()) >= int.Parse(strLine[i]))
                                                {
                                                    strErr += "第" + count.ToString() + "行第" + num.ToString() + "列日期不能小於最後結餘日期;";
                                                }
                                            }
                                        }
                                    }
                                    //if (!StringUtil.IsEmpty(strErr))
                                    //    throw new AlertException(strErr);
                                }

                                // 將訊息新增到Table中
                                DataRow drIn = dtDataIn.NewRow();
                                drIn["Serial_Number"] = strLine[0];
                                drIn["Stock_Date"] = Convert.ToDateTime(strLine[1].Substring(0, 4) + "/" + strLine[1].Substring(4, 2) + "/" + strLine[1].Substring(6, 2));
                                drIn["Type"] = Convert.ToInt32(strLine[2]);
                                drIn["Number"] = Convert.ToInt32(strLine[3]);
                                dtDataIn.Rows.Add(drIn);
                                // add by Ian Huang start
                                strMaxDate = "" == strMaxDate ? strLine[1] : int.Parse(strMaxDate) > int.Parse(strLine[1]) ? strMaxDate : strLine[1];
                                // add by Ian Huang end
                            }
                        }
                    }
                    count++;
                }
                #endregion 讀字符串，并檢查字符串格式(列數、每列的字符格式),并保存到臨時DataTable(dtDataIn)中

                // add by Ian Huang start
                #region 將 移轉出、移轉入 資料存入到DB
                // 將 移轉出、移轉入 資料存入到 MATERIEL_STOCKS_MOVE 表中以代替原 物料庫存移轉作業 功能
                string strFactoryRID = intPerso_Factory_Rid.ToString().Trim();

                DataRow[] drType59 = dtDataIn.Select(" Type=5 or Type=9 ", "Serial_Number,Stock_Date");

                for (int i = 0; i < drType59.Length; i++)
                {
                    string strSNum = drType59[i]["Serial_Number"].ToString().Trim();
                    string strSDate = DateTime.Parse(drType59[i]["Stock_Date"].ToString()).ToString("yyyy/MM/dd");
                    int iType = int.Parse(drType59[i]["Type"].ToString());
                    int iNumber = int.Parse(drType59[i]["Number"].ToString());

                    MATERIEL_STOCKS_MOVE moveModel = getexistModel(strSDate, strSNum, iNumber, iType, strFactoryRID);
                    bool bIsexist = isexistModel(strSDate, strSNum, iNumber, iType, strFactoryRID);

                    if (!bIsexist)
                    {
                        if (null == moveModel)
                        {
                            // add
                            moveModel = new MATERIEL_STOCKS_MOVE();
                            moveModel.Serial_Number = strSNum;
                            moveModel.Move_Number = iNumber;

                            if (5 == iType)
                            {
                                // 5 移轉出
                                moveModel.To_Factory_RID = 0;
                                moveModel.From_Factory_RID = Convert.ToInt32(strFactoryRID);
                            }
                            else
                            {
                                // 9 移轉入
                                moveModel.To_Factory_RID = Convert.ToInt32(strFactoryRID);
                                moveModel.From_Factory_RID = 0;
                            }

                            moveModel.Move_ID = GetMove_ID(strSDate.Trim());
                            moveModel.Move_Date = DateTime.Parse(strSDate);
                            moveModel.RCU = "INPUT";
                            AddM(moveModel);
                        }
                        else
                        {
                            // update
                            if (5 == iType)
                            {
                                // 5 移轉出
                                moveModel.From_Factory_RID = Convert.ToInt32(strFactoryRID);
                            }
                            else
                            {
                                // 9 移轉入
                                moveModel.To_Factory_RID = Convert.ToInt32(strFactoryRID);
                            }
                            Update(moveModel);
                        }
                    }
                }

                #endregion 將 移轉出、移轉入 資料存入到DB

                #region 沒有廠商結餘的資料新增廠商結餘
                List<string> Lstr = new List<string>(); // 如果沒有廠商結餘，則ADD Serial_Number

                Dictionary<string, int> dic01 = new Dictionary<string, int>();  //進貨
                Dictionary<string, int> dic02 = new Dictionary<string, int>();  //退貨
                Dictionary<string, int> dic03 = new Dictionary<string, int>();  //銷毀

                Dictionary<string, int> dic06 = new Dictionary<string, int>();  //抽驗
                Dictionary<string, int> dic07 = new Dictionary<string, int>();  //退件重寄
                Dictionary<string, int> dic08 = new Dictionary<string, int>();  //電訊單異動
                DataRow[] drType4 = dtDataIn.Select("", "Serial_Number");

                for (int i = 0; i < drType4.Length; i++)
                {
                    string strSNum = drType4[i]["Serial_Number"].ToString();
                    bool bHave04 = false;

                    for (int j = i; j < drType4.Length; j++)
                    {
                        if (strSNum != drType4[j]["Serial_Number"].ToString())
                        {
                            i = j - 1;
                            break;
                        }

                        if ("4" == drType4[j]["Type"].ToString())
                        {
                            bHave04 = true;
                        }

                        // 記錄 進貨 數量
                        if ("1" == drType4[j]["Type"].ToString())
                        {
                            if (dic01.ContainsKey(strSNum))
                            {
                                dic01[strSNum] += int.Parse(drType4[j]["Number"].ToString());
                            }
                            else
                            {
                                dic01.Add(strSNum, int.Parse(drType4[j]["Number"].ToString()));
                            }
                        }

                        // 記錄 退貨 數量
                        if ("2" == drType4[j]["Type"].ToString())
                        {
                            if (dic02.ContainsKey(strSNum))
                            {
                                dic02[strSNum] += int.Parse(drType4[j]["Number"].ToString());
                            }
                            else
                            {
                                dic02.Add(strSNum, int.Parse(drType4[j]["Number"].ToString()));
                            }
                        }

                        // 記錄 銷毀 數量
                        if ("3" == drType4[j]["Type"].ToString())
                        {
                            if (dic03.ContainsKey(strSNum))
                            {
                                dic03[strSNum] += int.Parse(drType4[j]["Number"].ToString());
                            }
                            else
                            {
                                dic03.Add(strSNum, int.Parse(drType4[j]["Number"].ToString()));
                            }
                        }

                        // 記錄 抽驗 數量
                        if ("6" == drType4[j]["Type"].ToString())
                        {
                            if (dic06.ContainsKey(strSNum))
                            {
                                dic06[strSNum] += int.Parse(drType4[j]["Number"].ToString());
                            }
                            else
                            {
                                dic06.Add(strSNum, int.Parse(drType4[j]["Number"].ToString()));
                            }
                        }

                        // 記錄 退件重寄 數量
                        if ("7" == drType4[j]["Type"].ToString())
                        {
                            if (dic07.ContainsKey(strSNum))
                            {
                                dic07[strSNum] += int.Parse(drType4[j]["Number"].ToString());
                            }
                            else
                            {
                                dic07.Add(strSNum, int.Parse(drType4[j]["Number"].ToString()));
                            }
                        }

                        // 記錄 電訊單異動 數量
                        if ("8" == drType4[j]["Type"].ToString())
                        {
                            if (dic08.ContainsKey(strSNum))
                            {
                                dic08[strSNum] += int.Parse(drType4[j]["Number"].ToString());
                            }
                            else
                            {
                                dic08.Add(strSNum, int.Parse(drType4[j]["Number"].ToString()));
                            }
                        }

                        i = j;

                    }

                    if (!bHave04)
                    {
                        Lstr.Add(strSNum);
                    }
                }

                // 如果有資料沒有廠商結餘
                if (Lstr.Count > 0)
                {
                    strMaxDate = strMaxDate.Substring(0, 4) + "/" + strMaxDate.Substring(4, 2) + "/" + strMaxDate.Substring(6, 2);
                    InOut000BL bl000 = new InOut000BL();
                    bl000.SaveSurplusSystemNum(Convert.ToDateTime(strMaxDate));

                    for (int i = 0; i < Lstr.Count; i++)
                    {
                        int iCount = 0;

                        if (dic01.ContainsKey(Lstr[i]))
                        {
                            iCount -= dic01[Lstr[i]];
                        }

                        if (dic02.ContainsKey(Lstr[i]))
                        {
                            iCount += dic02[Lstr[i]];
                        }

                        if (dic03.ContainsKey(Lstr[i]))
                        {
                            iCount += dic03[Lstr[i]];
                        }

                        if (dic06.ContainsKey(Lstr[i]))
                        {
                            iCount += dic06[Lstr[i]];
                        }

                        if (dic07.ContainsKey(Lstr[i]))
                        {
                            iCount += dic07[Lstr[i]];
                        }

                        if (dic08.ContainsKey(Lstr[i]))
                        {
                            iCount += dic08[Lstr[i]];
                        }

                        DataRow drIn = dtDataIn.NewRow();
                        drIn["Serial_Number"] = Lstr[i];
                        drIn["Type"] = 4;
                        drIn["Stock_Date"] = Convert.ToDateTime(strMaxDate);
                        drIn["Number"] = selectSTOCKS(strMaxDate, Lstr[i], int.Parse(strFactoryRID)) - iCount;
                        dtDataIn.Rows.Add(drIn);
                    }

                }

                // 排序 下方有檢查
                DataTable dtDataInSort = dtDataIn.Copy();
                dtDataInSort.DefaultView.Sort = "Serial_Number,Stock_Date,Type ASC";
                dtDataIn = dtDataInSort.DefaultView.ToTable();
                #endregion 沒有廠商結餘的資料新增廠商結餘


                // add by Ian Huang end

                #region 廠商資料及匯入資料有無檢查
                if (dtFactory.Rows.Count == 0)
                {
                    strErr += "匯入資料中沒有廠商編號！\n";
                }

                if (dtDataIn.Rows.Count == 0)
                {
                    strErr += "匯入資料中沒有庫存異動訊息！\n";
                }
                #endregion 廠商資料及匯入資料有無檢查

                //pan:2008/11/18
                #region 判斷物品是否存在
                for (int i = 0; i < dtDataIn.Rows.Count; i++)
                {
                    string str = dtDataIn.Rows[i]["Serial_Number"].ToString();
                    DataSet ds = new DataSet();

                    dirValues.Clear();
                    dirValues.Add("Serial_Number", str);
                    if ("A" == str.Substring(0, 1).ToUpper())// 信封
                    {
                        ds = dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                    }
                    else if ("B" == str.Substring(0, 1).ToUpper())// 卡單
                    {
                        ds = dao.GetList(SEL_CARD_EXPONENT, dirValues);
                    }
                    else if ("C" == str.Substring(0, 1).ToUpper())// DM
                    {
                        ds = dao.GetList(SEL_DMTYPE_INFO, dirValues);
                    }

                    if (ds == null || ds.Tables[0].Rows.Count == 0)
                    {
                        strErr += "物品編號'" + str + "'不存在！\n";
                    }
                    else
                    {
                        dtDataIn.Rows[i]["Materiel_RID"] = Convert.ToInt32(ds.Tables[0].Rows[0]["RID"]);
                    }
                }
                #endregion 判斷物品是否存在

                //pan:2008/10/29
                #region 判斷檔案內容各種物料(品名)是否結余（即type=4）
                for (int intRow = 0; intRow < dtDataIn.Rows.Count; intRow++)
                {
                    string str = dtDataIn.Rows[intRow]["Serial_Number"].ToString();
                    ArrayList al = new ArrayList();
                    for (int iRow = 0; iRow < dtDataIn.Rows.Count; iRow++)
                    {
                        if (str == dtDataIn.Rows[iRow]["Serial_Number"].ToString())
                        {
                            al.Add(dtDataIn.Rows[iRow]["Type"].ToString());
                        }
                    }
                    if (al.Count > 0)
                    {
                        ArrayList array = new ArrayList();
                        for (int i = 0; i < al.Count; i++)
                        {
                            if (al[i].ToString() == "4")
                            {
                                array.Add(al[i]);
                                break;
                            }
                        }
                        if (array.Count == 0)
                        {
                            strErr += "物料" + str + "沒有結余！\n";
                        }
                    }
                }
                #endregion 判斷檔案內容各種物料(品名)是否結余（即type=4）

                #region 檔案內容各種物料(品名)只能有一筆結余
                for (int intRow = 0; intRow < dtDataIn.Rows.Count - 1; intRow++)
                {
                    if (4 == Convert.ToInt32(dtDataIn.Rows[intRow]["Type"]))
                    {
                        for (int intRow1 = intRow + 1; intRow1 < dtDataIn.Rows.Count; intRow1++)
                        {
                            if (4 == Convert.ToInt32(dtDataIn.Rows[intRow1]["Type"]) &&
                                Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]) == Convert.ToString(dtDataIn.Rows[intRow1]["Serial_Number"]))
                            {
                                strErr += "檔案格式錯誤，[" + Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]) + "]只能有一筆結余。\n";
                            }
                        }
                    }
                }
                #endregion 檔案內容各種物料(品名)只能有一筆結余

                #region 廠商結余為負檢查
                for (int intRow = 0; intRow < dtDataIn.Rows.Count; intRow++)
                {
                    if (4 == Convert.ToInt32(dtDataIn.Rows[intRow]["Type"]))
                    {
                        if (Convert.ToInt32(dtDataIn.Rows[intRow]["Number"]) < 0)
                        {
                            strErr += "廠商餘為數不能為負\n";
                        }
                    }
                }
                #endregion 廠商結余為負檢查

                #region 排序檢查(按品名、日期、類別)
                DataTable dtSort = dtDataIn.Copy();
                dtSort.DefaultView.Sort = "Serial_Number,Stock_Date,Type ASC";
                dtSort = dtSort.DefaultView.ToTable();
                // 排序檢查
                for (int intRowNum = 0; intRowNum < dtDataIn.Rows.Count; intRowNum++)
                {
                    //if (Convert.ToString(dtDataIn.Rows[intRowNum]["Serial_Number"]) != Convert.ToString(dtSort.Rows[intRowNum]["Serial_Number"]) ||
                    //    Convert.ToDateTime(dtDataIn.Rows[intRowNum]["Stock_Date"]) != Convert.ToDateTime(dtSort.Rows[intRowNum]["Stock_Date"]) ||
                    //    Convert.ToInt32(dtDataIn.Rows[intRowNum]["Type"]) != Convert.ToInt32(dtSort.Rows[intRowNum]["Type"]))
                    //{
                    //    throw new AlertException(BizMessage.BizMsg.ALT_DEPOSITORY_010_02);
                    //}

                    if (Convert.ToString(dtDataIn.Rows[intRowNum]["Serial_Number"]) != Convert.ToString(dtSort.Rows[intRowNum]["Serial_Number"]))
                    {
                        strErr += "匯入文件中品名排序錯誤！\n";
                    }
                    else if (Convert.ToDateTime(dtDataIn.Rows[intRowNum]["Stock_Date"]) != Convert.ToDateTime(dtSort.Rows[intRowNum]["Stock_Date"]))
                    {
                        strErr += "匯入文件中日期排序錯誤！\n";
                    }
                    else if (Convert.ToInt32(dtDataIn.Rows[intRowNum]["Type"]) != Convert.ToInt32(dtSort.Rows[intRowNum]["Type"]))
                    {
                        strErr += "匯入文件中類別排序錯誤！\n";
                    }
                }
                #endregion 排序檢查(按品名、日期、類別)

                #region 每一Perso廠回饋檔，檔案內容各種物料(品名)的進貨、退貨、銷毀日期不能大於結余日期
                for (int intRow = 0; intRow < dtDataIn.Rows.Count; intRow++)
                {
                    // 結余
                    if (4 == Convert.ToInt32(dtDataIn.Rows[intRow]["Type"]))
                    {
                        for (int intRow1 = 0; intRow1 < intRow; intRow1++)
                        {
                            if (intRow != intRow1 &&
                                4 != Convert.ToInt32(dtDataIn.Rows[intRow1]["Type"]) &&
                                Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]) == Convert.ToString(dtDataIn.Rows[intRow1]["Serial_Number"]) &&
                                Convert.ToDateTime(dtDataIn.Rows[intRow]["Stock_Date"]) < Convert.ToDateTime(dtDataIn.Rows[intRow1]["Stock_Date"]))
                            {
                                strErr += "檔案格式錯誤，[" + Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]) + "]（" +
                                    getTypeName(Convert.ToString(dtDataIn.Rows[intRow1]["Type"])) + "）日期" +
                                    (Convert.ToDateTime(dtDataIn.Rows[intRow1]["Stock_Date"])).ToString("yyyy/MM/dd") + "大於結余日期" +
                                    (Convert.ToDateTime(dtDataIn.Rows[intRow]["Stock_Date"])).ToString("yyyy/MM/dd") + "。\n";
                            }
                        }
                    }
                }
                #endregion 每一Perso廠回饋檔，檔案內容各種物料(品名)的進貨、退貨、銷毀日期不能大於結余日期

                #region 各種物料(品名)之結余日期及任何異動日期皆不能小於上次結余日期
                string Serial_Number = "";
                for (int intRow = 0; intRow < dtDataIn.Rows.Count; intRow++)
                {
                    DataTable dtLastSurplusDate = null;
                    //if (Serial_Number != Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]))
                    if (Serial_Number != Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]) || Convert.ToInt16(dtDataIn.Rows[intRow]["Type"]) == 4)
                    {
                        Serial_Number = Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]);
                        // 取物料的上次結余日期及結余數量
                        dtLastSurplusDate = (DataTable)GetLastSurplusDateNum(Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]),
                                                        Convert.ToInt32(getFactoryRIDByShortName_EN(Convert.ToString(values["Factory_ShortName_EN"]))));
                    }

                    // 物料作過結余
                    if (null != dtLastSurplusDate &&
                        dtLastSurplusDate.Rows.Count > 0)
                    {
                        // 結余日期或異動日期不能小于上次結余日期
                        if (DateTime.Parse(Convert.ToDateTime(dtDataIn.Rows[intRow]["Stock_Date"]).ToString("yyyy/MM/dd 12:00:00")) <=
                            DateTime.Parse(Convert.ToDateTime(dtLastSurplusDate.Rows[0]["Stock_Date"]).ToString("yyyy/MM/dd 12:00:00")))
                        {
                            strErr += "檔案格式錯誤，[" + Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]) + "]結余日期或異動日期不能小于上次結余日期。\n";
                        }

                        // 該物料的最近一次結余日期
                        dtDataIn.Rows[intRow]["Last_Surplus_Date"] = Convert.ToDateTime(dtLastSurplusDate.Rows[0]["Stock_Date"]);
                        // 該物料的最近一次結余數量
                        dtDataIn.Rows[intRow]["Last_Surplus_Num"] = Convert.ToInt32(dtLastSurplusDate.Rows[0]["Number"]);
                    }
                    // 物料沒作過結余(理論上不存在此情況，數據庫中最少有一條結余記錄)
                    else
                    {
                        for (int i = 0; i < dtDataIn.Rows.Count; i++)
                        {
                            if (dtDataIn.Rows[i]["Serial_Number"].ToString() == Serial_Number)
                            {
                                // 該物料的最近一次結余日期
                                dtDataIn.Rows[intRow]["Last_Surplus_Date"] = Convert.ToDateTime(dtDataIn.Rows[i]["Stock_Date"]).AddDays(-1);
                                break;
                            }
                        }

                        //該物料的最近一次結余時間（數據庫中必須有一條結余記錄）
                        //dtDataIn.Rows[intRow]["Last_Surplus_Date"] = Convert.ToDateTime("1900-01-01");
                        // 該物料的最近一次結余數量
                        dtDataIn.Rows[intRow]["Last_Surplus_Num"] = 0;
                    }
                }
                #endregion 各種物料(品名)之結余日期及任何異動日期皆不能小於上次結余日期

                #region 同一檔案所有的物料(品名)結余日期必須一樣
                for (int intRow = 0; intRow < dtDataIn.Rows.Count; intRow++)
                {
                    if (4 == Convert.ToInt32(dtDataIn.Rows[intRow]["Type"]))
                    {
                        for (int intRow1 = intRow + 1; intRow1 < dtDataIn.Rows.Count; intRow1++)
                        {
                            if (4 == Convert.ToInt32(dtDataIn.Rows[intRow1]["Type"]) &&
                                Convert.ToString(dtDataIn.Rows[intRow]["Stock_Date"]) != Convert.ToString(dtDataIn.Rows[intRow1]["Stock_Date"])
                                )
                            {
                                strErr += "同一檔案所有的物料(品名)結余日期必須一樣";
                            }
                        }
                        break;
                    }
                }
                #endregion  同一檔案所有的物料(品名)結余日期必須一樣

                #region 匯入檔案時，要檢查檔案內的結余日期，必須是系統做過日結的日期
                for (int intRow = 0; intRow < dtDataIn.Rows.Count; intRow++)
                {
                    if (4 == Convert.ToInt32(dtDataIn.Rows[intRow]["Type"]))
                    {
                        //if (!isSurplusDate(Convert.ToString(dtDataIn.Rows[intRow]["Serial_Number"]),
                        //    Convert.ToDateTime(dtDataIn.Rows[intRow]["Stock_Date"]),
                        //    Convert.ToInt32(values["FactoryRID"])))
                        //{
                        //    throw new AlertException(BizMessage.BizMsg.ALT_DEPOSITORY_010_07);
                        //}

                        //pan:2008/11/18
                        if (!isSurplusDate(Convert.ToDateTime(dtDataIn.Rows[intRow]["Stock_Date"])))
                        {
                            strErr += "結余日期系統未做過日結！\n";
                        }
                    }
                }
                #endregion 匯入檔案時，要檢查檔案內的結余日期，必須是系統做過日結的日期

                dstDataIn = new DataSet();
                dstDataIn.Tables.Add(dtFactory);
                dstDataIn.Tables.Add(dtDataIn);

                if (strErr != "")
                {
                    string[] arg = new string[1];
                    arg[0] = dtFactory.Rows[0]["Factory_Name"].ToString();
                    Warning.SetWarning(GlobalString.WarningType.MaterialAutoDataIn, arg);
                }

                return dstDataIn;

            }

            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("匯入前對匯入文件進行格式檢查, CheckIn報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
            finally
            {
                // 關閉文件
                if (null != sr)
                {
                    sr.Close();
                }
            }
        }


        /// <summary>
        /// 驗證匯入字段是否滿足格式
        /// </summary>
        /// <param name="strColumn"></param>
        /// <param name="num"></param>
        /// <param name="count"></param>
        /// <returns></returns>
        private string CheckFileOneColumn(string strColumn, int num, int count)
        {
            string strErr = "";
            string Pattern = "";
            MatchCollection Matches;
            switch (num)
            {
                case 1:
                    Pattern = @"^[A-Z]\d{5}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為6位字符;";
                    }
                    break;
                case 2:
                    if (strColumn.Length != 8)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列時間格式不對;";
                        break;
                    }

                    string str1 = strColumn.Substring(0, 4) + "/" + strColumn.Substring(4, 2) + "/" + strColumn.Substring(6, 2);
                    try
                    {
                        DateTime dt = Convert.ToDateTime(str1);
                    }
                    catch
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列時間格式不對;";
                    }
                    break;
                case 3:
                    // edit by Ian Huang start
                    if (strColumn != "06" && strColumn != "07" && strColumn != "08" && strColumn != "04" && strColumn != "05" && strColumn != "09" && strColumn != "01" && strColumn != "02" && strColumn != "03")
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列必須為{01}{02}{03}{04}{05}{06}{07}{08}或{09}種的任意一字符串;";
                    }
                    // edit by Ian Huang end                    
                    break;
                case 4:
                    Pattern = @"^\d{9}$";
                    Matches = Regex.Matches(strColumn, Pattern, RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);
                    if (Matches.Count == 0)
                    {
                        strErr = "第" + count.ToString() + "行第" + num.ToString() + "列格式必須為9位以內數字;";
                    }
                    break;
                default:
                    break;
            }

            return strErr;
        }

        public string In(DataSet MaterielStocksIn)
        {
            //檢查MATERIEL_STOCKS
            CheckMATERIEL_STOCKS(MaterielStocksIn);

            try
            {


                DataTable dtFactory = MaterielStocksIn.Tables[0];
                DataTable dtMaterielStocksIn = MaterielStocksIn.Tables[1];

                // 物品的移入訊息
                List<object> listStocksMoveInOnDay = getStocksMoveInOnDay(dtMaterielStocksIn, dtFactory);

                // 物品的移出訊息
                List<object> listStocksMoveOutOnDay = getStocksMoveOutOnDay(dtMaterielStocksIn, dtFactory);

                // 從物料庫存消耗檔中取耗用量
                List<object> listMaterielUsedOnDay = getMaterielUsedOnDay(dtMaterielStocksIn, dtFactory);

                ////add by Ian Huang start
                //// 從物料庫存異動檔中取進貨、退貨、銷毀訊息
                //List<object> listStocksTransactionOnDay = getStocksTransactionOnDay(dtMaterielStocksIn, dtFactory);
                ////add by Ian Huang end

                if (listStocksMoveInOnDay != null)
                {
                    foreach (DataRow dr in listStocksMoveInOnDay)
                    {
                        dtMaterielStocksIn.Rows.Add(dr);
                    }
                }
                if (listStocksMoveOutOnDay != null)
                {
                    foreach (DataRow dr in listStocksMoveOutOnDay)
                    {
                        dtMaterielStocksIn.Rows.Add(dr);
                    }
                }
                if (listMaterielUsedOnDay != null)
                {
                    foreach (DataRow dr in listMaterielUsedOnDay)
                    {
                        dtMaterielStocksIn.Rows.Add(dr);
                    }
                }

                ////add by Ian Huang start
                //if (listStocksTransactionOnDay != null)
                //{
                //    foreach (DataRow dr in listStocksTransactionOnDay)
                //    {
                //        dtMaterielStocksIn.Rows.Add(dr);
                //    }
                //}
                ////add by Ian Huang end

                // 按物品編號、時間、類型排次序重新排序
                dtMaterielStocksIn.DefaultView.Sort = "Serial_Number,Stock_Date,Type ASC";
                dtMaterielStocksIn = dtMaterielStocksIn.DefaultView.ToTable();

                // 整理匯入資料訊息
                DoMaterielStocksIn(dtMaterielStocksIn, dtFactory);

                // 排序
                dtMaterielStocksIn.DefaultView.Sort = "Serial_Number,Stock_Date,Type ASC";
                dtMaterielStocksIn = dtMaterielStocksIn.DefaultView.ToTable();

                // 開始連接，并開始事務
                dao.OpenConnection();

                // 保存
                DataIn(dtMaterielStocksIn, dtFactory);

                dao.Commit();
                return "";

                //gvpbPersoStockIn.DataSource = null;
                //gvpbPersoStockIn.DataBind();

                // 匯入成功



            }
            catch (Exception ex)
            {

                dao.Rollback();
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("In報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                BatchBL Bbl = new BatchBL();
                Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
                return "erro";
            }
            finally
            {
                dao.CloseConnection();
            }

        }
        /// <summary>
        /// 按Perso廠商，從物料轉移檔中取物料的移入記錄
        /// </summary>
        /// <param name="dtMaterielStockIn">匯入DataTable</param>
        /// <param name="dtFactory">廠商DataTable</param>
        /// <returns>void</DataTable></returns>
        public List<object> getStocksMoveInOnDay(DataTable dtMaterielStockIn, DataTable dtFactory)
        {
            List<object> listStocksMoveInOnDay = new List<object>();
            string Serial_Number = "";
            try
            {
                #region 取物料移入記錄
                for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
                {
                    if (Serial_Number != Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]))
                    {
                        // 保存當前物品編號
                        Serial_Number = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]);
                        int Materiel_Type = 0;
                        // 信封
                        if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 2;
                        else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 1;
                        else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 3;

                        // 物料的上次結余日期
                        DateTime dtLastSurplusDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Date"]);
                        int lastNumber = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Num"]);
                        // 當前匯入文檔的物品結余日期
                        DateTime NowDateTime = DateTime.Now;
                        //DateTime NowDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Stock_Date"]);
                        //for (int i = 0; i < dtMaterielStockIn.Rows.Count; i++)
                        //{
                        //    if ("4" == dtMaterielStockIn.Rows[i]["type"].ToString())
                        //    {
                        //        NowDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[i]["Stock_Date"]);
                        //        break;
                        //    }
                        //}

                        // 取物料的移入訊息
                        DataTable dtblStoksMoveIn = StocksMoveIn(Convert.ToInt32(dtFactory.Rows[0]["Factory_RID"]),
                                                                Serial_Number,
                                                                dtLastSurplusDateTime,
                                                                NowDateTime);

                        // 如果物料移入訊息不為空
                        if (null != dtblStoksMoveIn)
                        {
                            foreach (DataRow drStoksMoveIn in dtblStoksMoveIn.Rows)
                            {
                                //DataRow drStocksMoveInOnDay = dtMaterielStockIn.NewRow();
                                //drStocksMoveInOnDay["Serial_Number"] = Serial_Number;
                                //drStocksMoveInOnDay["Materiel_Name"] = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Materiel_Name"]);
                                //drStocksMoveInOnDay["Stock_Date"] = Convert.ToDateTime(drStoksMoveIn["Move_Date"]);
                                //drStocksMoveInOnDay["Type"] = 7;  // 移入標識
                                //drStocksMoveInOnDay["Number"] = Convert.ToInt32(drStoksMoveIn["Move_Number"]);
                                //drStocksMoveInOnDay["Materiel_RID"] = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Materiel_RID"]);
                                //drStocksMoveInOnDay["System_Num"] = 0;
                                DataRow drStocksMoveInOnDay = dtMaterielStockIn.NewRow();
                                drStocksMoveInOnDay["Serial_Number"] = Serial_Number;
                                drStocksMoveInOnDay["Materiel_Name"] = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Materiel_Name"]);
                                drStocksMoveInOnDay["Last_Surplus_Date"] = dtLastSurplusDateTime;
                                drStocksMoveInOnDay["Last_Surplus_Num"] = lastNumber;
                                drStocksMoveInOnDay["Stock_Date"] = Convert.ToDateTime(drStoksMoveIn["Move_Date"]);
                                //update by Ian Huang start
                                drStocksMoveInOnDay["Type"] = 57;  // 移入標識
                                //update by Ian Huang end
                                drStocksMoveInOnDay["Number"] = Convert.ToInt32(drStoksMoveIn["Move_Number"]);
                                drStocksMoveInOnDay["Materiel_RID"] = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Materiel_RID"]);
                                drStocksMoveInOnDay["System_Num"] = 0;

                                // 零時存儲
                                listStocksMoveInOnDay.Add(drStocksMoveInOnDay);
                            }
                        }
                    }
                }
                #endregion 取物料移入記錄

                return listStocksMoveInOnDay;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("按Perso廠商，從物料轉移檔中取物料的移入記錄, getStocksMoveInOnDay報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
        }

        /// <summary>
        /// 取上次工作日后的所有工作日
        /// </summary>
        /// <param name="dtFrom"></param>
        /// <returns></returns>
        public ArrayList getAllWorkDate()
        {
            try
            {
                ArrayList arr = new ArrayList();
                DataSet dsWorkDate = dao.GetList(SEL_ALL_WORK_DATE, this.dirValues);
                if (null != dsWorkDate && dsWorkDate.Tables.Count > 0)
                {
                    for (int intLoop = 0; intLoop < dsWorkDate.Tables[0].Rows.Count - 1; intLoop++)
                    {
                        arr.Add(Convert.ToDateTime(dsWorkDate.Tables[0].Rows[intLoop][0]).ToString("yyyy/MM/dd"));
                    }
                }
                return arr;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取上次工作日后的所有工作日,getAllWorkDate報錯: " + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
                //ExceptionFactory.CreateCustomSaveException(BizMessage.BizMsg.ALT_DEPOSITORY_010_10, ex.Message, dao.LastCommands);
                //throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_010_10);
            }
        }

        /// <summary>
        /// 按Perso廠商，從物料轉移檔中取物料的移出記錄
        /// </summary>
        /// <param name="dtMaterielStockIn">匯入DataTable</param>
        /// <param name="dtFactory">廠商DataTable</param>
        /// <returns><DataTable>廠商的物料轉出記錄</DataTable></returns>
        public List<object> getStocksMoveOutOnDay(DataTable dtMaterielStockIn, DataTable dtFactory)
        {
            List<object> listStocksMoveOutOnDay = new List<object>();
            string Serial_Number = "";
            try
            {
                for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
                {
                    if (Serial_Number != Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]))
                    {
                        // 保存當前物品編號
                        Serial_Number = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]);
                        int Materiel_Type = 0;
                        // 信封
                        if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 2;
                        else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 1;
                        else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 3;


                        // 物料的上次結余日期
                        DateTime dtLastSurplusDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Date"]);
                        int lastNumber = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Num"]);


                        DateTime NowDateTime = DateTime.Now;
                        //// 當前匯入文檔的物品結余日期
                        //DateTime NowDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Stock_Date"]);
                        //for (int i = 0; i < dtMaterielStockIn.Rows.Count; i++)
                        //{
                        //    if ("4" == dtMaterielStockIn.Rows[i]["type"].ToString())
                        //    {
                        //        NowDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[i]["Stock_Date"]);
                        //        break;
                        //    }
                        //}

                        // 取物料的移出訊息
                        DataTable dtblStoksMoveOut = StocksMoveOut(Convert.ToInt32(dtFactory.Rows[0]["Factory_RID"]),
                                                               Serial_Number,
                                                                dtLastSurplusDateTime,
                                                                NowDateTime);

                        // 如果物料移出訊息不為空
                        if (null != dtblStoksMoveOut)
                        {
                            foreach (DataRow drStoksMoveOut in dtblStoksMoveOut.Rows)
                            {
                                DataRow drStocksMoveOutOnDay = dtMaterielStockIn.NewRow();
                                drStocksMoveOutOnDay["Serial_Number"] = Serial_Number;
                                drStocksMoveOutOnDay["Materiel_Name"] = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Materiel_Name"]);
                                drStocksMoveOutOnDay["Stock_Date"] = Convert.ToDateTime(drStoksMoveOut["Move_Date"]);
                                //update by Ian Huang start
                                drStocksMoveOutOnDay["Type"] = 58;  // 移出標識
                                //update by Ian Huang end

                                //MAX ADD 
                                drStocksMoveOutOnDay["Last_Surplus_Date"] = dtLastSurplusDateTime;
                                drStocksMoveOutOnDay["Last_Surplus_Num"] = lastNumber;

                                drStocksMoveOutOnDay["Number"] = Convert.ToInt32(drStoksMoveOut["Move_Number"]);
                                drStocksMoveOutOnDay["Materiel_RID"] = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Materiel_RID"]);
                                drStocksMoveOutOnDay["System_Num"] = 0;
                                // 保存
                                listStocksMoveOutOnDay.Add(drStocksMoveOutOnDay);
                            }
                        }
                    }
                }

                return listStocksMoveOutOnDay;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("按Perso廠商，從物料轉移檔中取物料的移出記錄, getStocksMoveOutOnDay報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
        }
        /// <summary>
        /// 從物料的上次結余日期至資料匯入日期，計算物料的每天耗用量
        /// 計算方法:物料當天的耗用量*（1+物料的損耗率）
        /// </summary>
        /// <param name="dtMaterielStockIn"></param>
        /// <returns></returns>
        public List<object> getMaterielUsedOnDay(DataTable dtMaterielStockIn, DataTable dtFactory)
        {
            // 復制Table;其中Number字段為實際使用量,System_Num字段為耗用量=實際使用量*(1+損耗率)
            List<object> listMaterielUsedOnDay = new List<object>();
            string Serial_Number = "";
            try
            {
                for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
                {
                    if (Serial_Number != Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]))
                    {
                        // 保存當前物品編號
                        Serial_Number = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]);

                        // 物料的上次結余日期
                        DateTime dtLastSurplusDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Date"]);
                        int lastNumber = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Num"]);
                        //DateTime dtStockDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Stock_Date"]);
                        //for (int i = 0; i < dtMaterielStockIn.Rows.Count; i++)
                        //{
                        //    if (dtMaterielStockIn.Rows[i]["type"].ToString() == "4")
                        //    {
                        //        dtStockDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[i]["Stock_Date"]);
                        //        break;
                        //    }
                        //}
                        DateTime dtStockDateTime = DateTime.Now;

                        // 取物料的庫存消耗
                        //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 ADD BY 楊昆 2009/09/17 start
                        //DataTable dtblSTOCKS_USED = MaterielUsedCount(Convert.ToInt32(dtFactory.Rows[0]["Factory_RID"]), Serial_Number, dtLastSurplusDateTime, dtStockDateTime);
                        InOut000BL BL000 = new InOut000BL();
                        DataTable dtblSTOCKS_USED = BL000.MaterielUsedCount(Convert.ToInt32(dtFactory.Rows[0]["Factory_RID"]), Serial_Number, dtLastSurplusDateTime, dtStockDateTime);
                        //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 ADD BY 楊昆 2009/09/17 end
                        // 如果物料庫存消耗檔不為空
                        if (null != dtblSTOCKS_USED)
                        {
                            foreach (DataRow drSTOCKS_USED in dtblSTOCKS_USED.Rows)
                            {
                                DataRow drMaterielCountUsedOnDay = dtMaterielStockIn.NewRow();
                                drMaterielCountUsedOnDay["Serial_Number"] = Serial_Number;
                                drMaterielCountUsedOnDay["Materiel_Name"] = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Materiel_Name"]);
                                drMaterielCountUsedOnDay["Stock_Date"] = Convert.ToDateTime(drSTOCKS_USED["Stock_Date"]);
                                drMaterielCountUsedOnDay["Last_Surplus_Date"] = dtLastSurplusDateTime;
                                drMaterielCountUsedOnDay["Last_Surplus_Num"] = lastNumber;
                                //update by Ian Huang start
                                drMaterielCountUsedOnDay["Type"] = 56;  // 耗用標識
                                //update by Ian Huang end
                                drMaterielCountUsedOnDay["Number"] = Convert.ToInt32(drSTOCKS_USED["Number"]);//沒有算損耗的
                                drMaterielCountUsedOnDay["Materiel_RID"] = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Materiel_RID"]);
                                drMaterielCountUsedOnDay["System_Num"] = Convert.ToInt32(drSTOCKS_USED["System_Num"]);//計算損耗的
                                // 保存
                                listMaterielUsedOnDay.Add(drMaterielCountUsedOnDay);
                            }
                        }
                    }
                }

                return listMaterielUsedOnDay;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("從物料的上次結余日期至資料匯入日期，計算物料的每天耗用量, getMaterielUsedOnDay報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 對匯入資料進行整理，并按物品生成每天的結余訊息
        /// </summary>
        /// <param name="dtMaterielStockIn">匯入DataTable</param>
        /// <param name="dtFactory">廠商DataTable</param>
        public void DoMaterielStocksIn(DataTable dtMaterielStockIn, DataTable dtFactory)
        {
            List<object> listDoMaterielStocksIn = new List<object>();
            string Serial_Number = "";// 物品編號
            DateTime SurplusDateTime = DateTime.Parse("1900-01-01");// 本次結余日期
            DateTime NowDateTime = DateTime.Parse(DateTime.Now.ToString("yyyy/MM/dd 12:00:00"));
            int SurplusSystemNum = 0;// 本次結余數量

            try
            {
                #region 取所有物品的庫存結余訊息
                for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
                {
                    if (Serial_Number != Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]))
                    {
                        // 保存當前物品編號
                        Serial_Number = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]);

                        bool SurplusNo1 = true;//標志---系統第一次作結余且是本次結余的第一天

                        // 計算結余開始日期(開始結余日期 = 最近一次的結余日期+1天)
                        SurplusDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Date"]).AddHours(24);
                        // 最近一次的結余數量
                        SurplusSystemNum = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Num"]);

                        // 當前匯入文檔的物品結余日期
                        for (int i = 0; i < dtMaterielStockIn.Rows.Count; i++)
                        {
                            if ("4" == dtMaterielStockIn.Rows[i]["type"].ToString() && Serial_Number == dtMaterielStockIn.Rows[i]["Serial_Number"].ToString())
                            {
                                NowDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[i]["Stock_Date"]);
                                break;
                            }
                        }

                        // 計算每天的結余數量
                        while (SurplusDateTime <= NowDateTime)
                        {
                            #region 計算當天的庫存結余
                            // 計算當天的庫存結余
                            for (int intRowSurplus = 0; intRowSurplus < dtMaterielStockIn.Rows.Count; intRowSurplus++)
                            {
                                if (Serial_Number == Convert.ToString(dtMaterielStockIn.Rows[intRowSurplus]["Serial_Number"])
                                    && (DateTime.Parse(SurplusDateTime.ToString("yyyy/MM/dd 00:00:00")) <=
                                        DateTime.Parse(Convert.ToDateTime(dtMaterielStockIn.Rows[intRowSurplus]["Stock_Date"]).ToString("yyyy/MM/dd 12:00:00"))
                                    && DateTime.Parse(SurplusDateTime.ToString("yyyy/MM/dd 23:59:59")) >=
                                        DateTime.Parse(Convert.ToDateTime(dtMaterielStockIn.Rows[intRowSurplus]["Stock_Date"]).ToString("yyyy/MM/dd 12:00:00"))
                                    || SurplusNo1 && DateTime.Parse(SurplusDateTime.ToString("yyyy/MM/dd 23:59:59")) >=
                                        Convert.ToDateTime(dtMaterielStockIn.Rows[intRowSurplus]["Stock_Date"])))
                                {
                                    int intType = Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Type"]);

                                    //update by Ian Huang start
                                    if (1 == intType)
                                        // 進貨
                                        SurplusSystemNum += Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (2 == intType)
                                        // 退貨
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (3 == intType)
                                        // 銷毀
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (56 == intType)
                                        // 耗用（需計算損耗）
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["System_Num"]);
                                    else if (57 == intType)
                                        // 移入
                                        SurplusSystemNum += Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (58 == intType)
                                        // 移出
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (6 == intType)
                                        // 抽驗
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (7 == intType)
                                        // 退件重寄
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    else if (8 == intType)
                                        // 電訊單異動
                                        SurplusSystemNum -= Convert.ToInt32(dtMaterielStockIn.Rows[intRowSurplus]["Number"]);
                                    //update by Ian Huang end
                                }
                            }
                            #endregion 計算當天的庫存結余

                            // 如果結余小余0，則為0
                            if (SurplusSystemNum < 0)
                            {
                                SurplusSystemNum = 0;
                            }

                            DataRow dr = dtMaterielStockIn.NewRow();
                            dr["Serial_Number"] = Serial_Number;
                            dr["Materiel_Name"] = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Materiel_Name"]);
                            dr["Stock_Date"] = SurplusDateTime;
                            // edit by Ian Huang start
                            dr["Type"] = 10;  // 系統結余
                            // edit by Ian Huang end
                            dr["Number"] = 0;
                            //dr["Materiel_RID"] = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Materiel_RID"]);
                            dr["System_Num"] = SurplusSystemNum;

                            // 將訊息添加到臨時List中
                            listDoMaterielStocksIn.Add(dr);
                            // 增加一天
                            SurplusDateTime = SurplusDateTime.AddHours(24);
                            // 
                            SurplusNo1 = false;//不是本次結余的第一天
                        }
                    }
                }
                #endregion 取所有物品的庫存結余訊息

                // 將物品的結余訊息添加到匯入資料中
                foreach (DataRow drSurplusDate in listDoMaterielStocksIn)
                {
                    dtMaterielStockIn.Rows.Add(drSurplusDate);
                }


            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("對匯入資料進行整理，并按物品生成每天的結余訊息, DoMaterielStocksIn報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// 判斷是否為用完為止
        /// </summary>
        /// <param name="strSerial_Number"></param>
        /// <returns></returns>
        public bool DmNotSafe_Type(string strSerial_Number)
        {
            bool n = true;
            try
            {
                if (strSerial_Number.Contains("C"))
                {
                    DataTable dtbl = dao.GetList("select * from dbo.DMTYPE_INFO where Serial_Number='" + strSerial_Number + "'").Tables[0];
                    if (dtbl.Rows.Count > 0)
                    {
                        if (dtbl.Rows[0]["Safe_Type"].ToString() == "3")
                            n = false;
                    }
                }

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("判斷是否為用完為止, DmNotSafe_Type報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return true;
            }
            return n;
        }


        /// <summary>
        /// 根據物料編號獲得對應名稱
        /// </summary>
        /// <param name="Serial_Number"></param>
        /// <returns></returns>
        public string getMateriel_Name(string Serial_Number)
        {
            string strRetrun = "";

            try
            {
                DataSet ds = new DataSet();
                dirValues.Clear();
                dirValues.Add("Serial_Number", Serial_Number);
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())// 信封
                {
                    ds = dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                }
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())// 卡單
                {
                    ds = dao.GetList(SEL_CARD_EXPONENT, dirValues);
                }
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())// DM
                {
                    ds = dao.GetList(SEL_DMTYPE_INFO, dirValues);
                }

                if (ds != null && ds.Tables[0].Rows.Count > 0)
                {
                    //獲得紙品物料名稱
                    strRetrun = ds.Tables[0].Rows[0]["name"].ToString();
                }
                else
                {
                    strRetrun = "";
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("根據物料編號獲得對應名稱, getMateriel_Name報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                strRetrun = "";
            }

            return strRetrun;
        }
        /// <summary>
        /// 保存匯入資料訊息
        /// </summary>
        /// <param name="dtMaterielStockIn">匯入資料表</param>
        /// <param name="dtblFactory">廠商訊息</param>
        //private void DataIn(DataTable dtMaterielStockIn, DataTable dtblFactory)
        //{
        //   try
        //   {            
        //       MATERIEL_STOCKS_MANAGE msmModel = new MATERIEL_STOCKS_MANAGE();
        //       string Serial_Number = "";
        //       string Materiel_Name = "";
        //       ArrayList arrWorkDate = getAllWorkDate();
        //       if (null == arrWorkDate)
        //       {
        //           return;
        //       }

        //       for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
        //       {
        //           DataRow drMaterielStockIn = dtMaterielStockIn.Rows[intRow];

        //           #region 刪除物料上次結餘和本次結餘之間的系統結餘，即Type = 4的記錄
        //           if (Convert.ToString(drMaterielStockIn["Serial_Number"]) != "")
        //           {
        //               Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);

        //               Materiel_Name = getMateriel_Name(Serial_Number);

        //               if (drMaterielStockIn["type"].ToString() == "4")
        //               {
        //                   DeleteSysSurplusData(Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]),
        //                                           Serial_Number,
        //                                           Convert.ToDateTime(drMaterielStockIn["Last_Surplus_Date"]),
        //                                           Convert.ToDateTime(drMaterielStockIn["Stock_Date"]));
        //               }
        //           }
        //           #endregion 刪除物料上次結餘和本次結餘之間的系統結餘，即Type = 4的記錄
        //           // 添加物料異動訊息
        //           if (1 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
        //               2 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
        //               3 == Convert.ToInt32(drMaterielStockIn["Type"]))
        //           {
        //               #region 物料異動                   
        //               msmModel.Number = Convert.ToInt32(drMaterielStockIn["Number"]);
        //               msmModel.Perso_Factory_RID = Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]);
        //               msmModel.Stock_Date = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]);
        //               msmModel.Type = Convert.ToString(drMaterielStockIn["Type"]);
        //               if (drMaterielStockIn["System_Num"].ToString() != "")
        //               {
        //                   msmModel.Invenroty_Remain = Convert.ToInt32(drMaterielStockIn["System_Num"]);
        //               }
        //               else
        //               {
        //                   msmModel.Invenroty_Remain = 0;
        //               }
        //               msmModel.Comment = Convert.ToString(drMaterielStockIn["Comment"]);
        //               msmModel.Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]); 
        //               Add(msmModel);
        //               #endregion 物料異動
        //           }
        //           // 添加物料盤整訊息
        //           else if (4 == Convert.ToInt32(drMaterielStockIn["Type"]))
        //           {
        //               int intWearRate = 0;
        //               #region 如果庫存結餘>廠商結餘
        //               if (dtMaterielStockIn.Rows[intRow + 1]["System_Num"].ToString() != "" && drMaterielStockIn["Number"].ToString() != "")
        //               {
        //                   if (Convert.ToInt32(dtMaterielStockIn.Rows[intRow + 1]["System_Num"]) >
        //                       Convert.ToInt32(drMaterielStockIn["Number"]))
        //                   {
        //                       string strMessage = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]).ToString("yyyy/MM/dd ") +
        //                                                       Convert.ToString(dtblFactory.Rows[0]["Factory_Name"]) + " " +
        //                                                       Materiel_Name +
        //                                                       "的庫存結餘大與廠商結餘";

        //                       string strMessageWearRate = "";

        //                       // 損耗率計算方法：最近一次結餘數量 + 進貨 - 退貨 - 銷毀 + 移入 - 移出 - 耗用（包括損耗）= 廠商結餘數量
        //                       // 耗用（包括損耗）= 最近一次結餘數量 - 廠商結餘數量 + 進貨 - 退貨 - 銷毀 + 移入 - 移出；
        //                       // 耗用（包括損耗） = 小計檔關聯的使用物品數量 * (1+實際損耗率)
        //                       // 實際損耗率 = 耗用（包括損耗）/ 小計檔關聯的使用物品數量 - 1;

        //                       // 小計檔關聯的使用物品數量
        //                       int intUsedCount = 0;
        //                       // 耗用（包括損耗） = 最近一次結餘數量 - 廠商結餘數量
        //                       int intUsedCountWear = 0;
        //                       if (drMaterielStockIn["Last_Surplus_Num"].ToString() != "" && drMaterielStockIn["Last_Surplus_Num"].ToString() != "0")
        //                       {
        //                           intUsedCountWear = Convert.ToInt32(drMaterielStockIn["Last_Surplus_Num"])
        //                                           - Convert.ToInt32(drMaterielStockIn["Number"]);
        //                       }
        //                       else
        //                       {
        //                           intUsedCountWear = Convert.ToInt32(drMaterielStockIn["Number"]);
        //                       }



        //                       DataRow[] drRate = dtMaterielStockIn.Select("Stock_Date <= #" + String.Format("{0:s}", Convert.ToDateTime(drMaterielStockIn["Stock_Date"])) + "#");
        //                       foreach (DataRow dr in drRate)
        //                       {
        //                           if (Convert.ToString(drMaterielStockIn["Serial_Number"]) == Convert.ToString(dr["Serial_Number"]))
        //                           // && Convert.ToDateTime (drMaterielStockIn["Stock_Date"]) > Convert.ToDateTime (dr["Stock_Date "]) )
        //                           {
        //                               if (6 == Convert.ToInt32(dr["Type"]))
        //                                   intUsedCount += Convert.ToInt32(dr["Number"]);
        //                               else if (1 == Convert.ToInt32(dr["Type"]) || 7 == Convert.ToInt32(dr["Type"]))
        //                                   intUsedCountWear += Convert.ToInt32(dr["Number"]);// + 進貨 + 移入
        //                               else if (2 == Convert.ToInt32(dr["Type"]) || 3 == Convert.ToInt32(dr["Type"]) ||
        //                                       8 == Convert.ToInt32(dr["Type"]))
        //                                   intUsedCountWear -= Convert.ToInt32(dr["Number"]);// - 退貨 - 銷毀 - 移出
        //                           }
        //                       }

        //                       if (intUsedCountWear > 0 && intUsedCount >= 0)
        //                       {
        //                           if (intUsedCount == 0)
        //                               intWearRate = 99;
        //                           else
        //                               // 損耗率 = 耗用（包括損耗）/ 小計檔關聯的使用物品數量 - 1;
        //                               intWearRate = Convert.ToInt32((Convert.ToDecimal(intUsedCountWear) / (Convert.ToDecimal(intUsedCount)) - 1) * 100);


        //                           strMessageWearRate = ";損耗率為：" + intWearRate.ToString() + "%;";

        //                           string[] arg = new string[3];
        //                           arg[0] = dtblFactory.Rows[0]["Factory_Name"].ToString();
        //                           arg[1] = Materiel_Name;
        //                           arg[2] = (intWearRate).ToString();
        //                           Warning.SetWarning(GlobalString.WarningType.MaterialDataIn, arg);
        //                           //}
        //                           //else if (intWearRate < 0)
        //                           //{
        //                           //    strMessageWearRate = ";損耗率為：" + (-intWearRate).ToString() + "%;";
        //                           //}
        //                       }
        //                       else
        //                       {
        //                           strMessageWearRate = ";";
        //                       }

        //                       strErr += strMessage + strMessageWearRate;
        //                   }
        //               }
        //               #endregion 如果庫存結餘>廠商結餘
        //               #region 庫存數量報警
        //               if (DmNotSafe_Type(drMaterielStockIn["Serial_Number"].ToString()))
        //               {
        //                   DataSet dtMateriel = this.GetMateriel(Convert.ToString(drMaterielStockIn["Serial_Number"]));
        //                   if (null != dtMateriel &&
        //                       dtMateriel.Tables.Count > 0 &&
        //                       dtMateriel.Tables[0].Rows.Count > 0)
        //                   {
        //                       // 最低安全庫存
        //                       if (GlobalString.SafeType.storage == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                       {

        //                           //廠商結余低于最低安全庫存數值時
        //                           if (Convert.ToInt32(drMaterielStockIn["Number"]) <
        //                               Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]))
        //                           {
        //                               strErr += Convert.ToString(dtblFactory.Rows[0]["Factory_Name"]) + "PERSO廠的" +
        //                                          dtMateriel.Tables[0].Rows[0]["Name"] + "紙品物料安全庫存量不足！";

        //                               string[] arg = new string[1];
        //                               arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                               Warning.SetWarning(GlobalString.WarningType.MaterialDataInSafe, arg);
        //                           }
        //                           // 安全天數
        //                       }
        //                       else if (GlobalString.SafeType.days == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                       {
        //                           // 檢查庫存是否充足
        //                           if (!this.CheckMaterielSafeDays(dtMateriel.Tables[0].Rows[0]["Serial_Number"].ToString(),
        //                                                   Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]),
        //                                                   Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]),
        //                                                   Convert.ToInt32(drMaterielStockIn["Number"])))
        //                           {
        //                               strErr += Convert.ToString(dtblFactory.Rows[0]["Factory_Name"]) + "PERSO廠的" +
        //                                           dtMateriel.Tables[0].Rows[0]["Name"].ToString() + "紙品物料安全庫存不足！";

        //                               string[] arg = new string[1];
        //                               arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                               Warning.SetWarning(GlobalString.WarningType.MaterialDataInSafe, arg);
        //                           }

        //                       }
        //                   }
        //               }
        //                #endregion 庫存數量報警
        //               #region 物料盤整
        //               msmModel.Number = Convert.ToInt32(drMaterielStockIn["Number"]);
        //               msmModel.Perso_Factory_RID = Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]);
        //               msmModel.Stock_Date = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]);
        //               msmModel.Type = "4";
        //               if (dtMaterielStockIn.Rows[intRow + 1]["System_Num"].ToString() != "")
        //               {
        //                   msmModel.Invenroty_Remain = Convert.ToInt32(dtMaterielStockIn.Rows[intRow + 1]["System_Num"].ToString());
        //               }
        //               else
        //               {
        //                   msmModel.Invenroty_Remain = 0;
        //               }
        //               msmModel.Comment = Convert.ToString(drMaterielStockIn["Comment"]);
        //               msmModel.Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);
        //               //if (intWearRate < 0)
        //               //{
        //               //msmModel.Real_Wear_Rate = -intWearRate;
        //               //}
        //               //else
        //               //{
        //               msmModel.Real_Wear_Rate = intWearRate;
        //               //}
        //               this.Add(msmModel);
        //               #endregion 物料盤整
        //           }
        //           // 添加物料每天結余訊息
        //           else if (5 == Convert.ToInt32(drMaterielStockIn["Type"]))
        //           {
        //               if (arrWorkDate.Contains(Convert.ToDateTime(drMaterielStockIn["Stock_Date"]).ToString("yyyy/MM/dd")))
        //               {
        //                   #region 物料每天結餘
        //                   MATERIEL_STOCKS msModel = new MATERIEL_STOCKS();
        //                   msModel.Number = Convert.ToInt32(drMaterielStockIn["System_Num"]);
        //                   msModel.Perso_Factory_RID = Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]);
        //                   msModel.Stock_Date = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]);
        //                   msModel.Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);
        //                   this.Addms(msModel);
        //                   #endregion 物料每天結餘

        //                   msmModel = this.getMSMModel(msModel);
        //                   //判斷當前是否為物品結餘時間(type=4)
        //                   if (msmModel.Stock_Date == msModel.Stock_Date)
        //                   {
        //                       //更新物料庫存管理信息的系統結餘
        //                       msmModel.Invenroty_Remain = msModel.Number;

        //                       this.dao.Update<MATERIEL_STOCKS_MANAGE>(msmModel, "RID");
        //                   }
        //               }
        //           }
        //       }
        //   }
        //   catch (Exception ex)
        //   {
        //       throw new Exception(ex.Message);
        //   }
        //}
        private void DataIn(DataTable dtMaterielStockIn, DataTable dtblFactory)
        {
            try
            {
                MATERIEL_STOCKS_MANAGE msmModel;
                ArrayList arrWorkDate = getAllWorkDate();

                string Serial_Number = "";
                string Materiel_Name = "";

                //先刪除掉DB中從上次結余以來的記錄！
                for (int intRowDel = 0; intRowDel < dtMaterielStockIn.Rows.Count; intRowDel++)
                {
                    DataRow drMaterielStockDel = dtMaterielStockIn.Rows[intRowDel];
                    if (Convert.ToString(drMaterielStockDel["Serial_Number"]) != "")
                    {
                        Serial_Number = Convert.ToString(drMaterielStockDel["Serial_Number"]);

                        if (drMaterielStockDel["type"].ToString() == "4")
                        {
                            DeleteSysSurplusData(Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]),
                                                    Serial_Number,
                                                    Convert.ToDateTime(drMaterielStockDel["Last_Surplus_Date"]),
                                                    Convert.ToDateTime(drMaterielStockDel["Stock_Date"]));
                        }
                    }
                }


                for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
                {
                    string strWearRate = "0";

                    DataRow drMaterielStockIn = dtMaterielStockIn.Rows[intRow];

                    Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);
                    Materiel_Name = getMateriel_Name(Serial_Number);

                    //#region 刪除物料上次結餘和本次結餘之間的系統結餘，即Type = 4的記錄
                    //if (Convert.ToString(drMaterielStockIn["Serial_Number"])!="")
                    //{
                    //    Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);

                    //    Materiel_Name = bl.getMateriel_Name(Serial_Number);

                    //    if (drMaterielStockIn["type"].ToString() == "4")
                    //    {
                    //        bl.DeleteSysSurplusData(Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]),
                    //                                Serial_Number,
                    //                                Convert.ToDateTime(drMaterielStockIn["Last_Surplus_Date"]),
                    //                                Convert.ToDateTime(drMaterielStockIn["Stock_Date"]));
                    //    }
                    //}
                    //#endregion 刪除物料上次結餘和本次結餘之間的系統結餘，即Type = 4的記錄

                    // 添加物料異動訊息
                    //edit by Ian Huang start
                    if (1 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                        2 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                        3 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                        6 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                        7 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                        8 == Convert.ToInt32(drMaterielStockIn["Type"]))
                    {
                        #region 物料異動
                        msmModel = new MATERIEL_STOCKS_MANAGE();
                        msmModel.Number = Convert.ToInt32(drMaterielStockIn["Number"]);
                        msmModel.Perso_Factory_RID = Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]);
                        msmModel.Stock_Date = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]);
                        msmModel.Type = Convert.ToString(drMaterielStockIn["Type"]);
                        if (drMaterielStockIn["System_Num"].ToString() != "")
                        {
                            msmModel.Invenroty_Remain = Convert.ToInt32(drMaterielStockIn["System_Num"]);
                        }
                        else
                        {
                            msmModel.Invenroty_Remain = 0;
                        }
                        msmModel.Comment = Convert.ToString(drMaterielStockIn["Comment"]);
                        msmModel.Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);

                        //add by Ian Huang start
                        //  先刪除日結寫入的 進貨、退貨、銷毀 信息
                        if (1 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                            2 == Convert.ToInt32(drMaterielStockIn["Type"]) ||
                            3 == Convert.ToInt32(drMaterielStockIn["Type"]))
                        {
                            DeleteSTOCKSMANAGE(Convert.ToDateTime(drMaterielStockIn["Stock_Date"]).ToString("yyyy/MM/dd 00:00:00"), CIMSClass.GlobalString.RCU.ACTIVED, Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]), Convert.ToString(drMaterielStockIn["Type"]), Convert.ToString(drMaterielStockIn["Serial_Number"]));
                        }
                        //add by Ian Huang end

                        Add(msmModel);
                        #endregion 物料異動
                    }
                    // 添加物料盤整訊息
                    else if (4 == Convert.ToInt32(drMaterielStockIn["Type"]))
                    {
                        int intWearRate = 0;

                        // add by Ian Huang start
                        decimal decWearRate = 0;
                        // add by Ian Huang end

                        #region 如果庫存結餘>廠商結餘
                        //應該這裡要調整 max 2011.03.18.
                        int TypeNo = 0;
                        String Serial_Number_temp = "";
                        for (int a = intRow; a < dtMaterielStockIn.Rows.Count; a++)
                        {
                            if (Convert.ToString(dtMaterielStockIn.Rows[a][3]).Equals("10"))
                            {
                                TypeNo = a - 1;
                                a = dtMaterielStockIn.Rows.Count;
                            }
                        }


                       // if (dtMaterielStockIn.Rows[intRow + 1]["System_Num"].ToString() != "" && drMaterielStockIn["Number"].ToString() != "")
                       // {
                        if (Convert.ToInt32(dtMaterielStockIn.Rows[TypeNo + 1]["System_Num"]) >
                                Convert.ToInt32(drMaterielStockIn["Number"]))
                            {
                                string strMessage = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]).ToString("yyyy/MM/dd ") +
                                                                Convert.ToString(dtblFactory.Rows[0]["Factory_Name"]) + " " +
                                                                Materiel_Name +
                                                                "的庫存結餘大與廠商結餘";

                                string strMessageWearRate = "";

                                // 損耗率計算方法：最近一次結餘數量 + 進貨 - 退貨 - 銷毀 + 移入 - 移出 - 耗用（包括損耗）= 廠商結餘數量
                                // 耗用（包括損耗）= 最近一次結餘數量 - 廠商結餘數量 + 進貨 - 退貨 - 銷毀 + 移入 - 移出；
                                // 耗用（包括損耗） = 小計檔關聯的使用物品數量 * (1+實際損耗率)
                                // 實際損耗率 = 耗用（包括損耗）/ 小計檔關聯的使用物品數量 - 1;
                                // 實際損耗率(2010/10/28日修改) = ( (上次廠商結餘-本次廠商結餘)/(耗用（包括損耗）+抽驗+退件重寄+電訊單異動)  )-1

                                // 小計檔關聯的使用物品數量
                                int intUsedCount = 0;
                                // 最近一次廠商結余+進貨-退貨-銷毀+移入-移出-物料消耗=本次廠商結余（匯入為4的庫存數量）
                                int intUsedCountWear = 0;
                                if (drMaterielStockIn["Last_Surplus_Num"].ToString() != "" && drMaterielStockIn["Last_Surplus_Num"].ToString() != "0")
                                {
                                    intUsedCountWear = Convert.ToInt32(drMaterielStockIn["Last_Surplus_Num"])
                                                    - Convert.ToInt32(drMaterielStockIn["Number"]);
                                }
                                else
                                {
                                    intUsedCountWear = Convert.ToInt32(drMaterielStockIn["Number"]);
                                }


                                DataRow[] drRate = dtMaterielStockIn.Select("Stock_Date <= #" + String.Format("{0:s}", Convert.ToDateTime(drMaterielStockIn["Stock_Date"])) + "#");

                                //edit by Ian Huang start
                                foreach (DataRow dr in drRate)
                                {
                                    if (Convert.ToString(drMaterielStockIn["Serial_Number"]) == Convert.ToString(dr["Serial_Number"]))
                                    // && Convert.ToDateTime (drMaterielStockIn["Stock_Date"]) > Convert.ToDateTime (dr["Stock_Date "]) )
                                    {
                                        if (56 == Convert.ToInt32(dr["Type"]))
                                            intUsedCount += Convert.ToInt32(dr["Number"]);
                                        else if (1 == Convert.ToInt32(dr["Type"]) || 57 == Convert.ToInt32(dr["Type"]))
                                            intUsedCountWear += Convert.ToInt32(dr["Number"]);// + 進貨 + 移入
                                        else if (2 == Convert.ToInt32(dr["Type"]) || 3 == Convert.ToInt32(dr["Type"]) || 58 == Convert.ToInt32(dr["Type"]))
                                            intUsedCountWear -= Convert.ToInt32(dr["Number"]);// - 退貨 - 銷毀 - 移出
                                        else if (6 == Convert.ToInt32(dr["Type"]) || 7 == Convert.ToInt32(dr["Type"]) || 8 == Convert.ToInt32(dr["Type"]))
                                            intUsedCount += Convert.ToInt32(dr["Number"]);// 耗用（包括損耗）+抽驗+退件重寄+電訊單異動
                                    }
                                }
                                //edit by Ian Huang end
                                if (intUsedCountWear > 0 && intUsedCount >= 0)
                                {
                                    // edit by Ian Huang start
                                    if (intUsedCount == 0)
                                    {
                                        intWearRate = 99;
                                        decWearRate = 99;
                                    }
                                    else
                                    {
                                        // 損耗率 = 耗用（包括損耗）/ 小計檔關聯的使用物品數量 - 1;
                                        intWearRate = Convert.ToInt32((Convert.ToDecimal(intUsedCountWear) / (Convert.ToDecimal(intUsedCount)) - 1) * 100);
                                        decWearRate = Math.Round((Convert.ToDecimal(intUsedCountWear) / (Convert.ToDecimal(intUsedCount)) - 1) * 100, 2, MidpointRounding.AwayFromZero);
                                    }
                                    // edit by Ian Huang start

                                    strMessageWearRate = ";損耗率為：" + intWearRate.ToString() + "%;";

                                    string[] arg = new string[3];
                                    arg[0] = dtblFactory.Rows[0]["Factory_Name"].ToString();
                                    arg[1] = getMateriel_Name(dtMaterielStockIn.Rows[intRow]["Serial_Number"].ToString());
                                    // edit by Ian Huang start
                                    //arg[2] = intWearRate.ToString();
                                    arg[2] = decWearRate.ToString();
                                    // edit by Ian Huang end
                                    //200911CR-廠商耗損率高於基本檔設定的耗損率，即發耗損率警訊 Edit By Yangkun 2009/11/25 start
                                    if (int.Parse(GetWearRate(dtMaterielStockIn.Rows[intRow]["Serial_Number"].ToString()).ToString()) < intWearRate)
                                    {
                                        Warning.SetWarning(GlobalString.WarningType.MaterialDataIn, arg);
                                    }
                                    //200911CR-廠商耗損率高於基本檔設定的耗損率，即發耗損率警訊 Edit By Yangkun 2009/11/25 end

                                    strWearRate = intWearRate.ToString();
                                }
                                else
                                {
                                    strMessageWearRate = ";";
                                }

                                // ShowMessage(strMessage + strMessageWearRate);
                            }
                       // }
                        #endregion 如果庫存結餘>廠商結餘

                        #region 庫存數量報警
                        if (DmNotSafe_Type(drMaterielStockIn["Serial_Number"].ToString()))
                        {
                            DataSet dtMateriel = GetMateriel(Convert.ToString(drMaterielStockIn["Serial_Number"]));
                            if (null != dtMateriel &&
                                dtMateriel.Tables.Count > 0 &&
                                dtMateriel.Tables[0].Rows.Count > 0)
                            {
                                // 最低安全庫存
                                if (GlobalString.SafeType.storage == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
                                {
                                    // 廠商結餘低於最低安全庫存數值時
                                    if (Convert.ToInt32(drMaterielStockIn["Number"]) <
                                        Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]))
                                    {
                                        string[] arg = new string[1];
                                        arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
                                        Warning.SetWarning(GlobalString.WarningType.MaterialDataInSafe, arg);

                                        //ShowMessage(Convert.ToString(dtblFactory.Rows[0]["Factory_Name"]) + "PERSO厰的" +
                                        //  Materiel_Name + "紙品物料安全庫存量不足！");
                                    }
                                    // 安全天數
                                }
                                else if (GlobalString.SafeType.days == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
                                {
                                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 ADD BY 楊昆 2009/09/30 start
                                    InOut000BL BL000 = new InOut000BL();
                                    // 檢查庫存是否充足
                                    //if (!CheckMaterielSafeDays(dtMateriel.Tables[0].Rows[0]["Serial_Number"].ToString(),
                                    if (!BL000.CheckMaterielSafeDays(dtMateriel.Tables[0].Rows[0]["Serial_Number"].ToString(),
                                        //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 ADD BY 楊昆 2009/09/30 end
                                                            Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]),
                                                            Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]),
                                                            Convert.ToInt32(drMaterielStockIn["Number"])))
                                    {
                                        string[] arg = new string[1];
                                        arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
                                        Warning.SetWarning(GlobalString.WarningType.MaterialDataInSafe, arg);

                                        //ShowMessage(Convert.ToString(dtblFactory.Rows[0]["Factory_Name"]) + "PERSO厰的" +
                                        //            Materiel_Name + "紙品物料安全庫存量不足！");
                                    }

                                }
                            }
                        }
                        #endregion 庫存數量報警

                        #region 物料盤整
                        msmModel = new MATERIEL_STOCKS_MANAGE();
                        msmModel.Number = Convert.ToInt32(drMaterielStockIn["Number"]);
                        msmModel.Perso_Factory_RID = Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]);
                        msmModel.Stock_Date = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]);
                        msmModel.Type = "4";
                        //if (drMaterielStockIn["System_Num"].ToString() != "")
                        //{
                        //    msmModel.Invenroty_Remain = Convert.ToInt32(drMaterielStockIn["System_Num"]);
                        //}
                        //else
                        //{
                        //    msmModel.Invenroty_Remain = Convert.ToInt32(drMaterielStockIn["Number"]);
                        //}

                        //將系統結余數保存到廠商結余的記錄的另一個欄位上！ max edit
                        //if (dtMaterielStockIn.Rows[intRow + 1]["System_Num"].ToString() != "")
                        //{
                        //    msmModel.Invenroty_Remain = Convert.ToInt32(dtMaterielStockIn.Rows[intRow + 1]["System_Num"].ToString());
                        //}
                        //else
                        //{
                        //    msmModel.Invenroty_Remain = 0;
                        //}
                        if (dtMaterielStockIn.Rows[TypeNo + 1]["System_Num"].ToString() != "")
                        {
                            msmModel.Invenroty_Remain = Convert.ToInt32(dtMaterielStockIn.Rows[TypeNo + 1]["System_Num"].ToString());
                        }
                        else
                        {
                            msmModel.Invenroty_Remain = 0;
                        }



                        msmModel.Comment = Convert.ToString(drMaterielStockIn["Comment"]);
                        msmModel.Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);
                        //if (intWearRate < 0)
                        //{
                        //    msmModel.Real_Wear_Rate = -intWearRate;
                        //}
                        //else
                        //{
                        //edti by Ian Huang start
                        //msmModel.Real_Wear_Rate = intWearRate;
                        msmModel.Real_Wear_Rate = decWearRate;
                        //edti by Ian Huang end
                        //}
                        Add(msmModel);
                        #endregion 物料盤整
                    }
                    // 添加物料每天結余訊息
                    else if (10 == Convert.ToInt32(drMaterielStockIn["Type"]))
                    {
                        if (arrWorkDate.Contains(Convert.ToDateTime(drMaterielStockIn["Stock_Date"]).ToString("yyyy/MM/dd")))
                        {
                            #region 物料每天結餘
                            MATERIEL_STOCKS msModel = new MATERIEL_STOCKS();
                            msModel.Number = Convert.ToInt32(drMaterielStockIn["System_Num"]);
                            msModel.Perso_Factory_RID = Convert.ToInt32(dtblFactory.Rows[0]["Factory_RID"]);
                            msModel.Stock_Date = Convert.ToDateTime(drMaterielStockIn["Stock_Date"]);
                            msModel.Serial_Number = Convert.ToString(drMaterielStockIn["Serial_Number"]);
                            Addms(msModel);
                            #endregion 物料每天結餘

                            msmModel = getMSMModel(msModel);
                            if (msmModel != null)
                            {
                                //判斷當前是否為物品結餘時間(type=4)
                                //if (msmModel.Stock_Date == msModel.Stock_Date)
                                //{
                                //更新物料庫存管理信息的系統結餘
                                msmModel.Invenroty_Remain = msModel.Number;
                                dao.Update<MATERIEL_STOCKS_MANAGE>(msmModel, "RID");
                                //}
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("DataIn報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(ex.Message);
            }
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="Factory_RID">Perso廠商RID</param>
        /// <param name="Materiel_Type">物料類型</param>
        /// <param name="Materiel_RID">物料RID</param>
        /// <param name="lastSurplusDateTime">最近一次的結餘日期</param>
        /// <param name="thisSurplusDateTime">本次結餘日期</param>
        /// <returns>DataTable<廠商轉入記錄></returns>
        public void DeleteSysSurplusData(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                //this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                dao.ExecuteNonQuery(DEL_MATERIEL_STOCKS_SYS, this.dirValues);
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("DeleteSysSurplusData報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                //ExceptionFactory.CreateCustomSaveException(BizMessage.BizMsg.ALT_DEPOSITORY_010_10, ex.Message, dao.LastCommands);
                throw ex;
            }
        }

        /// <summary>
        /// 從物料移動檔中，按廠商、物品、時間段取物品的轉入記錄
        /// </summary>
        /// <param name="Factory_RID">Perso廠商RID</param>
        /// <param name="Materiel_Type">物料類型</param>
        /// <param name="Materiel_RID">物料RID</param>
        /// <param name="lastSurplusDateTime">最近一次的結餘日期</param>
        /// <param name="thisSurplusDateTime">本次結餘日期</param>
        /// <returns>DataTable<廠商轉入記錄></returns>
        public DataTable StocksMoveIn(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            DataTable dtStocksMoveIn = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                DataSet dstMaterielStocksMove = dao.GetList(SEL_MATERIEL_STOCKS_MOVE_IN, this.dirValues);
                if (null != dstMaterielStocksMove
                            && dstMaterielStocksMove.Tables.Count > 0
                            && dstMaterielStocksMove.Tables[0].Rows.Count > 0)
                {
                    dtStocksMoveIn = dstMaterielStocksMove.Tables[0];
                }

                return dtStocksMoveIn;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("從物料移動檔中，按廠商、物品、時間段取物品的轉入記錄, StocksMoveIn報錯: " + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 從物料移動檔中，按廠商、物品、時間段取物品的轉出記錄
        /// </summary>
        /// <param name="Factory_RID">Perso廠商RID</param>
        /// <param name="Materiel_Type">物料類型</param>
        /// <param name="Materiel_RID">物料RID</param>
        /// <param name="lastSurplusDateTime">最近一次的結餘日期</param>
        /// <param name="thisSurplusDateTime">本次結餘日期</param>
        /// <returns>DataTable<廠商轉出記錄></returns>
        public DataTable StocksMoveOut(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            DataTable dtblStocksMoveOut = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                DataSet dstMaterielStocksMove = dao.GetList(SEL_MATERIEL_STOCKS_MOVE_OUT, this.dirValues);
                if (null != dstMaterielStocksMove
                            && dstMaterielStocksMove.Tables.Count > 0
                            && dstMaterielStocksMove.Tables[0].Rows.Count > 0)
                {
                    dtblStocksMoveOut = dstMaterielStocksMove.Tables[0];
                }

                return dtblStocksMoveOut;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("從物料移動檔中，按廠商、物品、時間段取物品的轉出記錄, StocksMoveOut報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// 計算物料庫存消耗檔
        /// </summary>
        /// <param name="Factory_RID">Perso廠商RID</param>
        /// <param name="Serial_Number">物料編號</param>    
        /// <param name="lastSurplusDateTime">最近一次的結餘日期</param>
        /// <param name="thisSurplusDateTime">本次結餘日期</param>
        /// <returns>DataTable<物料使用記錄></returns>
        public DataTable MaterielUsedCount(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            DataTable dtSubtotal_Import = null;
            try
            {
                string strFirst = Serial_Number.Substring(0, 1).ToUpper();
                string strSQL = "";
                if (strFirst.Equals("A"))
                {
                    // strSQL = SEL_MATERIAL_USED_ENVELOPE;                 
                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 start
                    strSQL = SEL_MATERIAL_USED_ENVELOPE_REPLACE;
                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 end
                }
                else if (strFirst.Equals("B"))
                {
                    //strSQL = SEL_MATERIAL_USED_CARD_EXPONENT;
                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 start
                    strSQL = SEL_MATERIAL_USED_CARD_EXPONENT_REPLACE;
                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 end
                }
                else if (strFirst.Equals("C"))
                {
                    //strSQL = SEL_MATERIAL_USED_DM;
                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 start
                    strSQL = SEL_MATERIAL_USED_DM_REPLACE;
                    //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 end
                }
                else
                {
                    return null;
                }

                dirValues.Clear();
                dirValues.Add("Perso_Factory_RID", Factory_RID);
                dirValues.Add("Serial_Number", Serial_Number);
                dirValues.Add("From_Date_Time", lastSurplusDateTime);
                dirValues.Add("End_Date_Time", thisSurplusDateTime);
                DataSet dstSTOCKS_USED = dao.GetList(strSQL, dirValues);
                if (null != dstSTOCKS_USED && dstSTOCKS_USED.Tables.Count > 0 &&
                                dstSTOCKS_USED.Tables[0].Rows.Count > 0)
                {
                    dtSubtotal_Import = dstSTOCKS_USED.Tables[0];
                    dtSubtotal_Import.Columns.Add(new DataColumn("System_Num", Type.GetType("System.Int32")));
                    for (int intRow = 0; intRow < dtSubtotal_Import.Rows.Count; intRow++)
                    {
                        // 取物品的損耗率(關聯到物品表，取物品表的損耗率）
                        Decimal dWear_Rate = GetWearRate(Serial_Number);
                        //Decimal dWear_Rate = 0.0M;
                        // 系統耗用量
                        dtSubtotal_Import.Rows[intRow]["System_Num"] = Convert.ToInt32(dtSubtotal_Import.Rows[intRow]["Number"]) * (dWear_Rate / 100 + 1);
                    }
                }
                return dtSubtotal_Import;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("計算物料庫存消耗檔, MaterielUsedCount報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
                //ExceptionFactory.CreateCustomSaveException(BizMessage.BizMsg.ALT_DEPOSITORY_010_10, ex.Message, dao.LastCommands);
                //throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_010_10);
            }
        }
        /// <summary>
        /// 計算物料庫存消耗檔并保存入資料庫
        /// </summary>         
        /// <param name="DateTime">本次結餘日期</param>
        /// 
        public List<string> SaveMaterielUsedCount(DateTime Date)
        {
            DataTable dtSubtotal_Import = null;
            List<string> lstSerielNumber = new List<string>();
            MATERIEL_STOCKS_USED msuModel = new MATERIEL_STOCKS_USED();
            try
            {

                //string strSQL = SEL_MATERIAL_USED_ENVELOPE_S + SEL_MATERIAL_USED_CARD_EXPONENT_S + SEL_MATERIAL_USED_DM_S;
                //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 start
                string strSQL = SEL_MATERIAL_USED_ENVELOPE_S_REPLACE + SEL_MATERIAL_USED_CARD_EXPONENT_S_REPLACE + SEL_MATERIAL_USED_DM_S_REPLACE;
                //200908CR物料的消耗計算改為用小計檔的「替換前」版面計算 add by 楊昆 2009/09/03 end
                string delSQL = "Delete MATERIEL_STOCKS_USED where  Stock_Date =@End_Date_Time ";

                dirValues.Clear();
                //dirValues.Add("From_Date_Time", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("End_Date_Time", Date);
                DataSet dstSTOCKS_USED = dao.GetList(strSQL, dirValues);
                dao.ExecuteNonQuery(delSQL, dirValues);
                for (int i = 0; i < dstSTOCKS_USED.Tables.Count; i++)
                {
                    if (null != dstSTOCKS_USED && dstSTOCKS_USED.Tables.Count > 0 &&
                                    dstSTOCKS_USED.Tables[i].Rows.Count > 0)
                    {
                        dtSubtotal_Import = dstSTOCKS_USED.Tables[i];
                        dtSubtotal_Import.Columns.Add(new DataColumn("System_Num", Type.GetType("System.Int32")));
                        for (int intRow = 0; intRow < dtSubtotal_Import.Rows.Count; intRow++)
                        {
                            // 取物品的損耗率(關聯到物品表，取物品表的損耗率）                    
                            Decimal dWear_Rate = GetWearRate(dtSubtotal_Import.Rows[intRow]["Serial_Number"].ToString());
                            // 系統耗用量
                            dtSubtotal_Import.Rows[intRow]["System_Num"] = Convert.ToInt32(dtSubtotal_Import.Rows[intRow]["Number"]) * (dWear_Rate / 100 + 1);

                            if (dtSubtotal_Import.Rows[intRow]["Serial_Number"].ToString() != null && dtSubtotal_Import.Rows[intRow]["Serial_Number"].ToString() != "")
                            {
                                // 保存物料品名編號，為判斷物料的庫存和安全水位作準備
                                if (-1 == lstSerielNumber.IndexOf(dtSubtotal_Import.Rows[intRow]["Serial_Number"].ToString()))
                                {
                                    lstSerielNumber.Add(dtSubtotal_Import.Rows[intRow]["Serial_Number"].ToString());
                                }
                                if (dtSubtotal_Import.Rows[intRow]["Stock_Date"].ToString() != null && dtSubtotal_Import.Rows[intRow]["Stock_Date"].ToString() != "")
                                {
                                    msuModel.Stock_Date = DateTime.Parse(dtSubtotal_Import.Rows[intRow]["Stock_Date"].ToString());
                                }
                                msuModel.Number = long.Parse(dtSubtotal_Import.Rows[intRow]["System_Num"].ToString());
                                msuModel.Serial_Number = dtSubtotal_Import.Rows[intRow]["Serial_Number"].ToString();
                                msuModel.Perso_Factory_RID = Convert.ToInt32(dtSubtotal_Import.Rows[intRow]["Perso_Factory_RID"].ToString());
                                dao.Add<MATERIEL_STOCKS_USED>(msuModel, "RID");
                            }
                        }
                    }
                }
                return lstSerielNumber;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("計算物料庫存消耗檔并保存入資料庫, SaveMaterielUsedCount報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
                //ExceptionFactory.CreateCustomSaveException(BizMessage.BizMsg.ALT_DEPOSITORY_010_10, ex.Message, dao.LastCommands);
                //throw new Exception(BizMessage.BizMsg.ALT_DEPOSITORY_010_10);
            }
        }

        /// <summary>
        /// 新增加記錄
        /// </summary>
        /// <param name="msmModel"></param>
        public void Add(MATERIEL_STOCKS_MANAGE msmModel)
        {
            try
            {
                dao.Add<MATERIEL_STOCKS_MANAGE>(msmModel, "RID");
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("新增加記錄, Add報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        public void Addms(MATERIEL_STOCKS msModel)
        {
            try
            {
                dao.Add<MATERIEL_STOCKS>(msModel, "RID");
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("Addms報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 取物料的物品、RID等訊息
        /// </summary>
        /// <param name="Serial_Number">品名編號</param>
        /// <returns><DataTable>物料DataTable</returns>
        public DataSet GetMateriel(string Serial_Number)
        {
            DataSet dtsMateriel = null;
            try
            {
                // 取物料的品名
                this.dirValues.Clear();
                this.dirValues.Add("serial_number", Serial_Number);

                // 信封
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_ENVELOPE_INFO, this.dirValues);
                // 寄卡單
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_CARD_EXPONENT, this.dirValues);
                // DM
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_DMTYPE_INFO, this.dirValues);
                return dtsMateriel;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取物料的物品、RID等訊息, GetMateriel報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 取物品的損耗率
        /// </summary>    
        /// <param name="Serial_Number">物品編號 1：信封；2：寄卡單；3：DM</param>
        /// <returns>Decimal<物品的耗用率></returns>
        public Decimal GetWearRate(string Serial_Number)
        {
            Decimal dWearRate = 0;
            DataSet dstWearRate = null;

            try
            {
                dirValues.Clear();
                dirValues.Add("Serial_Number", Serial_Number);
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())// 信封
                {
                    dstWearRate = dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                }
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())// 卡單
                {
                    dstWearRate = dao.GetList(SEL_CARD_EXPONENT, dirValues);
                }
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())// DM
                {
                    dstWearRate = dao.GetList(SEL_DMTYPE_INFO, dirValues);
                }

                if (null != dstWearRate &&
                        dstWearRate.Tables.Count > 0 &&
                        dstWearRate.Tables[0].Rows.Count > 0)
                {
                    // 取損耗率
                    dWearRate = Convert.ToDecimal(dstWearRate.Tables[0].Rows[0]["Wear_Rate"]);
                }

                return dWearRate;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("取物品的損耗率, GetWearRate報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// 獲得物料庫存管理信息
        /// </summary>
        /// <param name="msModel"></param>
        /// <returns></returns>
        public MATERIEL_STOCKS_MANAGE getMSMModel(MATERIEL_STOCKS msModel)
        {
            dirValues.Clear();
            dirValues.Add("perso_factory_rid", msModel.Perso_Factory_RID);
            dirValues.Add("serial_number", msModel.Serial_Number);
            MATERIEL_STOCKS_MANAGE msmModel = dao.GetModel<MATERIEL_STOCKS_MANAGE>("select top 1 * from MATERIEL_STOCKS_MANAGE where rst='A' and perso_factory_rid=@perso_factory_rid and serial_number=@serial_number and type='4' order by stock_date desc", dirValues);

            return msmModel;
        }
        /// <summary>
        /// 檢查物料的庫存是否安全（安全天數）
        /// </summary>
        /// <param name="Materiel_RID"></param>
        /// <param name="Materiel_Type"></param>
        /// <param name="Factory_RID"></param>
        /// <param name="Days"></param>
        /// <returns></returns>
        //public bool CheckMaterielSafeDays(string Serial_Number,
        //                                int Factory_RID,
        //                                int Days,
        //                                int Stock_Number)
        //{
        //    bool blCheckMaterielSafeDays = true;
        //    Days = Days + 1;   // 為了適應匯入時的函數需要，需要多減一天
        //    DateTime dtStartTime = DateTime.Now.AddDays(-Days);
        //    DataTable dtblSubtotal_Import = MaterielUsedCount(Factory_RID,
        //                                        Serial_Number,
        //                                        dtStartTime,
        //                                        DateTime.Now);

        //    int intMaterielWear = 0;
        //    if (null != dtblSubtotal_Import &&
        //        dtblSubtotal_Import.Rows.Count > 0)
        //    {
        //        // 前N天的耗用量
        //        for (int intRow = 0; intRow < dtblSubtotal_Import.Rows.Count; intRow++)
        //        {
        //            intMaterielWear += Convert.ToInt32(dtblSubtotal_Import.Rows[intRow]["System_Num"]);
        //        }

        //        // 如果庫存小于前N天的耗用量
        //        if (Stock_Number < intMaterielWear)
        //        {
        //            blCheckMaterielSafeDays = false;
        //        }
        //    }

        //    return blCheckMaterielSafeDays;
        //}
        /// <summary>
        /// 根據物料和數量計算實際的數量！
        /// </summary>
        /// <param name="MNumber"></param>
        /// <param name="MCount"></param>
        /// <returns></returns>
        public int ComputeMaterialNumber(string MNumber, long MCount)
        {
            int iReturn = 0;

            try
            {
                decimal dWear_Rate = this.GetWearRate(MNumber);
                iReturn = Convert.ToInt32(MCount * (dWear_Rate / 100 + 1));
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("根據物料和數量計算實際的數量！, ComputeMaterialNumber報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }

            return iReturn;
        }

        //add by Ian Huang start
        /// <summary>
        /// 按Perso廠商，從物料庫存異動檔中取進貨、退貨、銷毀記錄
        /// </summary>
        /// <param name="dtMaterielStockIn">匯入DataTable</param>
        /// <param name="dtFactory">廠商DataTable</param>
        /// <returns>void</DataTable></returns>
        public List<object> getStocksTransactionOnDay(DataTable dtMaterielStockIn, DataTable dtFactory)
        {
            List<object> listStocksTransactionOnDay = new List<object>();
            string Serial_Number = "";
            try
            {
                #region 取進貨、退貨、銷毀記錄
                for (int intRow = 0; intRow < dtMaterielStockIn.Rows.Count; intRow++)
                {
                    if (Serial_Number != Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]))
                    {
                        // 保存當前物品編號
                        Serial_Number = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Serial_Number"]);
                        int Materiel_Type = 0;
                        // 信封
                        if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 2;
                        else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 1;
                        else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                            Materiel_Type = 3;

                        // 物料的上次結餘日期
                        DateTime dtLastSurplusDateTime = Convert.ToDateTime(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Date"]);
                        int lastNumber = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Last_Surplus_Num"]);

                        // 計算物料系統結餘需要算到系統當前日期。
                        DateTime NowDateTime = DateTime.Now;

                        // 取物料的移入訊息
                        DataTable dtblStoksTransaction = StocksTransaction(Convert.ToInt32(dtFactory.Rows[0]["Factory_RID"]),
                                                                    Serial_Number,
                                                                    dtLastSurplusDateTime,
                                                                    NowDateTime);

                        // 如果物料移入訊息不為空
                        if (null != dtblStoksTransaction)
                        {
                            foreach (DataRow drStoksTransaction in dtblStoksTransaction.Rows)
                            {
                                DataRow drStocksTransactionOnDay = dtMaterielStockIn.NewRow();
                                drStocksTransactionOnDay["Serial_Number"] = Serial_Number;
                                drStocksTransactionOnDay["Materiel_Name"] = Convert.ToString(dtMaterielStockIn.Rows[intRow]["Materiel_Name"]);
                                drStocksTransactionOnDay["Last_Surplus_Date"] = dtLastSurplusDateTime;
                                drStocksTransactionOnDay["Last_Surplus_Num"] = lastNumber;
                                drStocksTransactionOnDay["Stock_Date"] = Convert.ToDateTime(drStoksTransaction["Transaction_Date"]);
                                drStocksTransactionOnDay["Type"] = Convert.ToInt32(drStoksTransaction["Param_Code"]);  // Param_Code 對應 異動類型
                                drStocksTransactionOnDay["Number"] = Convert.ToInt32(drStoksTransaction["Transaction_Amount"]);
                                drStocksTransactionOnDay["Materiel_RID"] = Convert.ToInt32(dtMaterielStockIn.Rows[intRow]["Materiel_RID"]);
                                drStocksTransactionOnDay["System_Num"] = 0;

                                // 零時存儲
                                listStocksTransactionOnDay.Add(drStocksTransactionOnDay);
                            }
                        }
                    }
                }
                #endregion 取物料移入記錄

                return listStocksTransactionOnDay;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("按Perso廠商，從物料庫存異動檔中取進貨、退貨、銷毀記錄, getStocksTransactionOnDay報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// 從物料異動檔中，按廠商、物品、時間段物料的進貨、退貨、銷毀記錄
        /// </summary>
        /// <param name="Factory_RID">Perso廠商RID</param>
        /// <param name="Serial_Number">物料類型</param>
        /// <param name="lastSurplusDateTime">最近一次的結餘日期</param>
        /// <param name="thisSurplusDateTime">本次結餘日期</param>
        /// <returns>DataTable<物料的進貨、退貨、銷毀記錄></returns>
        public DataTable StocksTransaction(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            DataTable dtStocksTransaction = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("perso_factory_rid", Factory_RID);
                this.dirValues.Add("Serial_Number", Serial_Number);
                this.dirValues.Add("lastSurplusDateTime", DateTime.Parse(lastSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                this.dirValues.Add("thisSurplusDateTime", DateTime.Parse(thisSurplusDateTime.ToString("yyyy/MM/dd 23:59:59")));
                DataSet dstMaterielStocksMove = dao.GetList(SEL_MATERIEL_STOCKS_TRANSACTION, this.dirValues);
                if (null != dstMaterielStocksMove
                            && dstMaterielStocksMove.Tables.Count > 0
                            && dstMaterielStocksMove.Tables[0].Rows.Count > 0)
                {
                    dtStocksTransaction = dstMaterielStocksMove.Tables[0];
                }

                return dtStocksTransaction;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("從物料異動檔中，按廠商、物品、時間段物料的進貨、退貨、銷毀記錄, StocksTransaction報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// 刪除日結寫入MATERIEL_STOCKS_MANAGE表的 進貨、退貨、銷毀 信息，以避免出現重複數據
        /// </summary>
        /// <param name="strStockDate">異動日期</param>
        /// <param name="strRCU">異動人員</param>
        /// <param name="iPersoFactoryRID">廠商id</param>
        /// <param name="strType">類型</param>
        /// <param name="strSerialNumber">物料類型</param>
        public void DeleteSTOCKSMANAGE(string strStockDate, string strRCU, int iPersoFactoryRID, string strType, string strSerialNumber)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("Stock_Date", strStockDate);
                dirValues.Add("RCU", strRCU);
                dirValues.Add("Perso_Factory_RID ", iPersoFactoryRID);
                dirValues.Add("Type", strType);
                dirValues.Add("Serial_Number", strSerialNumber);

                dao.ExecuteNonQuery(DELETE_MATERIEL_STOCKS_MANAGE, dirValues);

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("刪除日結寫入MATERIEL_STOCKS_MANAGE表的 進貨、退貨、銷毀 信息，以避免出現重複數據, DeleteSTOCKSMANAGE報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// 獲得系統結餘
        /// </summary>
        /// <param name="strStockdate">結餘日期</param>
        /// <param name="strSerialNum">Serial Number</param>
        /// <param name="strFactoryRID">person廠ID</param>
        /// <returns></returns>
        public int selectSTOCKS(string strStockdate, string strSerialNum, int iFactoryRID)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("Stock_date", strStockdate);
                dirValues.Add("Serial_Number", strSerialNum);
                dirValues.Add("Perso_Factory_RID ", iFactoryRID);

                DataSet ds = dao.GetList(SEL_MATERIEL_STOCKS, dirValues);

                if (null != ds
                            && ds.Tables.Count > 0
                            && ds.Tables[0].Rows.Count > 0)
                {
                    return int.Parse(ds.Tables[0].Rows[0]["Number"].ToString());
                }

                return 0;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("獲得系統結餘, selectSTOCKS報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        public MATERIEL_STOCKS_MOVE getexistModel(string strMoveDate, string strSerialNumber, int iMoveNumber, int iType, string strFactoryRID)
        {
            MATERIEL_STOCKS_MOVE model = null;
            string strSQL = SEL_MATERIEL_STOCKS_MOVE;

            if (5 == iType)
            {
                strSQL = strSQL + " and From_Factory_RID = 0 and To_Factory_RID <> ";
            }
            else
            {
                strSQL = strSQL + " and To_Factory_RID = 0  and From_Factory_RID <> ";
            }

            strSQL = strSQL + strFactoryRID;

            dirValues.Clear();
            dirValues.Add("Move_Date", strMoveDate);
            dirValues.Add("Serial_Number", strSerialNumber);
            dirValues.Add("Move_Number", iMoveNumber);

            try
            {
                model = dao.GetModel<MATERIEL_STOCKS_MOVE>(strSQL, dirValues);
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("getexistModel報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            return model;
        }

        public bool isexistModel(string strMoveDate, string strSerialNumber, int iMoveNumber, int iType, string strFactoryRID)
        {
            DataSet ds = null;
            string strSQL = SEL_MATERIEL_STOCKS_MOVE;

            if (5 == iType)
            {
                strSQL = strSQL + " and From_Factory_RID = ";
            }
            else
            {
                strSQL = strSQL + " and To_Factory_RID = ";
            }

            strSQL = strSQL + strFactoryRID;

            dirValues.Clear();
            dirValues.Add("Move_Date", strMoveDate);
            dirValues.Add("Serial_Number", strSerialNumber);
            dirValues.Add("Move_Number", iMoveNumber);

            try
            {
                ds = dao.GetList(strSQL, dirValues);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("isexistModel報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            return false;
        }

        public void delMove(string strFactoryRID, string strMoveDate, string strRCT)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("Factory_RID", strFactoryRID);

                string strU1 = UPDATE_MATERIEL_STOCKS_MOVE_1;
                string strU2 = UPDATE_MATERIEL_STOCKS_MOVE_2;
                string strD1 = DEL_MATERIEL_STOCKS_MOVE_1;
                string strD2 = DEL_MATERIEL_STOCKS_MOVE_2;

                string strCon = "";

                if ("" == strRCT.Trim())
                {
                    strCon = " and Move_Date = '" + strMoveDate + "'";
                }

                if ("" == strMoveDate.Trim())
                {
                    strCon = " and convert(varchar(10),RCT,111) = '" + strRCT + "'";
                }


                dao.ExecuteNonQuery(strU1 + strCon, dirValues);
                dao.ExecuteNonQuery(strU2 + strCon, dirValues);
                dao.ExecuteNonQuery(strD1 + strCon, dirValues);
                dao.ExecuteNonQuery(strD2 + strCon, dirValues);

            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("delMove報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        public void AddM(MATERIEL_STOCKS_MOVE model)
        {
            try
            {
                //dao.Add<MATERIEL_STOCKS_MOVE>(model);
                // rid = Convert.ToInt32(dao.AddAndGetID<MATERIEL_STOCKS_MOVE>(model, "Rid"));
                dirValues.Clear();
                dirValues.Add("Move_Date", model.Move_Date);
                dirValues.Add("RCU", model.RCU);
                dirValues.Add("RUU", "JOB");
                dirValues.Add("RCT", DateTime.Now);
                dirValues.Add("RUT", DateTime.Now);
                dirValues.Add("RST", GlobalString.RST.ACTIVED);
                dirValues.Add("Move_Number", model.Move_Number);
                dirValues.Add("From_Factory_RID", model.From_Factory_RID);
                dirValues.Add("To_Factory_RID", model.To_Factory_RID);
                dirValues.Add("Move_ID", model.Move_ID);
                dirValues.Add("Serial_Number", model.Serial_Number);

                dao.ExecuteNonQuery(INSERT_MATERIEL_STOCKS_MOVE, dirValues);


            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("AddM報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// 獲取資料表中相同YYYYMMDD的轉移單號的最大ID
        /// </summary>
        /// <param name="Move_Date">移動日期</param>
        /// <returns></returns>
        public string GetMove_ID(String Move_Date)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("move_date1", Move_Date + " 00:00:00");
                dirValues.Add("move_date2", Move_Date + " 23:59:59");
                DateTime dtMove_Date = Convert.ToDateTime(Move_Date);

                // 取轉移日期當天的最大轉移單號
                DataSet dtsMaxMoveID = dao.GetList(SEL_MAX_MOVE_ID, dirValues);
                if (dtsMaxMoveID.Tables[0].Rows.Count > 0)
                {
                    int intMaxID = Convert.ToInt32(dtsMaxMoveID.Tables[0].Rows[0]["Move_ID"].ToString().Substring(8, 2));
                    intMaxID++;
                    if (intMaxID > 9)
                    {
                        return dtMove_Date.ToString("yyyyMMdd") + intMaxID.ToString();
                    }
                    else
                    {
                        return dtMove_Date.ToString("yyyyMMdd") + "0" + intMaxID.ToString();
                    }
                }
                else
                {
                    return dtMove_Date.ToString("yyyyMMdd") + "01";
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("獲取資料表中相同YYYYMMDD的轉移單號的最大ID, GetMove_ID報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        public void Update(MATERIEL_STOCKS_MOVE model)
        {
            try
            {
                dao.Update<MATERIEL_STOCKS_MOVE>(model, "RID");
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 補充Log內容
                LogFactory.Write("Update(MATERIEL_STOCKS_MOVE)報錯:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        //add by Ian Huang end

    }
}
