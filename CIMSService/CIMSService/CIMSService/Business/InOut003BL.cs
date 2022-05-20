//*****************************************
//*  �@    �̡GGaoAi
//*  �\�໡���G�妸�鵲�@�~
//*  �Ыؤ���G2008-12-3
//*  �ק����G
//*  �ק�O���G
//*            ��2009-09-01
//*              �ק� ����
//*                      1.�������e�P�����Z���t�Ӳ��ʫH��
//*                      2.�鵲�ɹ�����e�������p�p�ɩM�����e�������t�Ӳ����ɧ@�B�z
//*                      3.���Ʈ��ӻP�N��O�Υδ����e�������p�p�ɩM�����e�������t�Ӳ����ɭp�� 
//*****************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Data;
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
    class InOut003BL : BaseLogic
    {
        #region SQL�y�y
        public const string SEL_ALL_SHOULD_SURPLUS_CARDTYPE = "SELECT DISTINCT B.Perso_Factory_RID,F.Factory_ShortName_CN,B.TYPE,B.AFFINITY,B.PHOTO,CT.NAME FROM (SELECT FCI.Perso_Factory_RID,FCI.TYPE,FCI.AFFINITY,FCI.PHOTO " +
                        "FROM FACTORY_CHANGE_IMPORT FCI " +
                        "WHERE FCI.RST = 'A' AND FCI.Date_Time >= @date_time_start AND FCI.Date_Time <= @date_time_end " +
                        "UNION " +
                        " SELECT DS.Perso_Factory_RID,CT.TYPE,CT.AFFINITY,CT.PHOTO " +
                        "FROM DEPOSITORY_STOCK DS INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DS.Space_Short_RID = CT.RID " +
                        "WHERE DS.RST = 'A' AND DS.Income_Date >= @date_time_start AND DS.Income_Date <= @date_time_end " +
                        "UNION " +
                        " SELECT DC.Perso_Factory_RID,CT.TYPE,CT.AFFINITY,CT.PHOTO " +
                        "FROM DEPOSITORY_CANCEL DC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DC.Space_Short_RID = CT.RID " +
                        "WHERE DC.RST = 'A' AND DC.Cancel_Date >= @date_time_start AND DC.Cancel_Date <= @date_time_end " +
                        "UNION " +
                        " SELECT DR.Perso_Factory_RID,CT.TYPE,CT.AFFINITY,CT.PHOTO " +
                        "FROM DEPOSITORY_RESTOCK DR INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DR.Space_Short_RID = CT.RID " +
                        "WHERE DR.RST = 'A' AND DR.Reincome_Date >= @date_time_start AND DR.Reincome_Date <= @date_time_end " +
                        "UNION " +
                        " SELECT STI.Perso_Factory_RID,STI.TYPE,STI.AFFINITY,STI.PHOTO " +
                        "FROM SUBTOTAL_IMPORT STI " +
                        "WHERE STI.RST = 'A' AND STI.Date_Time >= @date_time_start AND STI.Date_Time <= @date_time_end " +
                        "UNION " +
                        " SELECT CTSM.From_Factory_RID AS Perso_Factory_RID,CT.TYPE,CT.AFFINITY,CT.PHOTO " +
                        "FROM CARDTYPE_STOCKS_MOVE CTSM INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CTSM.CardType_RID = CT.RID " +
                        "WHERE CTSM.RST = 'A' AND CTSM.Move_Date >= @date_time_start AND CTSM.Move_Date<=@date_time_end " +
                        "UNION " +
                        " SELECT CTSM.To_Factory_RID AS Perso_Factory_RID,CT.TYPE,CT.AFFINITY,CT.PHOTO " +
                        "FROM CARDTYPE_STOCKS_MOVE CTSM INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CTSM.CardType_RID = CT.RID " +
                        "WHERE CTSM.RST = 'A' AND CTSM.Move_Date >= @date_time_start AND CTSM.Move_Date<=@date_time_end " +
                        "UNION " +
                        " SELECT CS.Perso_Factory_RID,CT.TYPE,CT.AFFINITY,CT.PHOTO " +
                        "FROM CARDTYPE_STOCKS CS INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CS.CardType_RID = CT.RID " +
                        "WHERE CS.RST = 'A' AND CS.Stock_Date >= @stock_date_start AND CS.Stock_Date<= @stock_date_end AND CS.Stocks_Number >0) B " +
                        "INNER JOIN Factory F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND B.Perso_Factory_RID = F.RID " +
                        "INNER JOIN Card_Type CT ON CT.RST = 'A' AND B.TYPE = CT.TYPE AND B.AFFINITY = CT.AFFINITY AND B.PHOTO = CT.PHOTO ";
       
        public const string SEL_CARD_TYPE = "SELECT RID,Name FROM CARD_TYPE WHERE RST='A' AND TYPE = @type AND AFFINITY = @affinity AND PHOTO = @photo";

        public const string SEL_CARDTYPE_STATUS = "select Status_Name from dbo.CARDTYPE_STATUS where RID = @rid";

        public const string SEL_LAST_SURPLUS_DAY = "SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS WHERE RST = 'A' ORDER BY  Stock_Date DESC";

        public const string SEL_LAST_SURPLUS_DAY_NEXT = "SELECT TOP 1 DATE_TIME FROM WORK_DATE WHERE RST='A' AND Is_WorkDay='Y' AND DATE_TIME > @LastSurplusDate ORDER BY DATE_TIME";

        public const string CON_CHECK_WORKDATE = " SELECT COUNT(*) FROM WORK_DATE WHERE RST = 'A' AND Is_WorkDay='Y' AND Date_Time >= @date_start AND Date_Time <= @date_end ";

        public const string CON_DATE_SURPLUS_CHECK = " SELECT COUNT(*) " +
                            "FROM CARDTYPE_STOCKS " +
                            "WHERE RST = 'A' AND Stock_Date >= @date_start AND Stock_Date <= @date_end ";

        public const string CON_CHECK_DATE_SURPLUS_BEFORE = " SELECT COUNT(*) FROM CARDTYPE_STOCKS WHERE RST = 'A' AND CONVERT(char(10), Stock_Date, 111) = ( SELECT TOP 1 CONVERT(char(10), Date_Time, 111) " +
                            "FROM WORK_DATE WHERE RST = 'A' AND  Is_WorkDay='Y' AND Date_Time < @date ORDER BY Date_Time DESC) ";

        #region �p��t�Ϊ��t�ӵ��l,�}����t�Ӳ��ʸ�T�M�t�β��ʸ�T�O�_�۲šC
        //(�t�ӶפJ��T���`)
        public const string SEL_FACTORY_IMPORT_STOCKS = " SELECT FCI.Perso_Factory_RID,F.Factory_ShortName_CN,FCI.TYPE,FCI.AFFINITY,FCI.PHOTO,CT.NAME,CS.Status_Name,SUM(Number) as Number " +
                                "FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND FCI.Perso_Factory_RID = F.RID " +
                                "INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO " +
                                "WHERE FCI.RST = 'A' AND FCI.Date_Time >= @date_time_start AND FCI.Date_Time <= @date_time_end " +
                                "GROUP BY FCI.Perso_Factory_RID,F.Factory_ShortName_CN,FCI.TYPE,FCI.AFFINITY,FCI.PHOTO,CT.NAME,CS.Status_Name ";
        //(�t�ΤJ�w�έp)
        public const string SEL_SYS_IN_STOCKS = " SELECT F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME,SUM(Income_Number) as Number " +
                                "FROM DEPOSITORY_STOCK DS INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DS.Space_Short_RID = CT.RID " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND DS.Perso_Factory_RID = F.RID " +
                                "WHERE DS.RST = 'A' AND DS.Income_Date >= @date_time_start AND DS.Income_Date <= @date_time_end " +
                                "GROUP BY F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME ";
        //(�t�ΰh�f�έp)
        public const string SEL_SYS_RETURN_STOCKS = " SELECT F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME,SUM(Cancel_Number) as Number " +
                                "FROM DEPOSITORY_CANCEL DC INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DC.Space_Short_RID = CT.RID " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND DC.Perso_Factory_RID = F.RID " +
                                "WHERE DC.RST = 'A' AND DC.Cancel_Date >= @date_time_start AND DC.Cancel_Date <= @date_time_end " +
                                "GROUP BY F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME ";
        //(�t�ΦA�J�w�έp)
        public const string SEL_SYS_DEPOSITORY_RESTOCK = " SELECT F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME,SUM(Reincome_Number) as Number " +
                                "FROM DEPOSITORY_RESTOCK DR INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND DR.Space_Short_RID = CT.RID " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND DR.Perso_Factory_RID = F.RID " +
                                "WHERE DR.RST = 'A' AND DR.Reincome_Date >= @date_time_start AND DR.Reincome_Date <= @date_time_end " +
                                "GROUP BY F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME ";
        //(�t��3D�BDA�BPM�BRN)
        public const string SEL_SYS_SUBTOTAL_TYPE = " SELECT F.RID,F.Factory_ShortName_CN,STI.TYPE,STI.AFFINITY,STI.PHOTO,MCT.Type_Name,CT.NAME,SUM(Number) as Number " +
                                ", CTO.Type as OLDTYPE , CTO.AFFINITY as OLDAFFINITY , CTO.PHOTO as OLDPHOTO , CTO.NAME as OLDNAME " +
                                "FROM SUBTOTAL_IMPORT STI INNER JOIN MAKE_CARD_TYPE MCT ON MCT.RST = 'A' AND STI.MakeCardType_RID = MCT.RID AND (MCT.Type_Name = '3D' OR MCT.Type_Name = 'DA' OR MCT.Type_Name = 'PM' OR MCT.Type_Name = 'RN') " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND STI.Perso_Factory_RID = F.RID " +
                                " inner join Card_Type CTO on STI.Old_CardType_rid = CTO.rid " +
                                "INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND STI.TYPE = CT.TYPE AND STI.AFFINITY = CT.AFFINITY AND STI.PHOTO = CT.PHOTO " +
                                "WHERE STI.RST = 'A' AND STI.Date_Time >= @date_time_start AND STI.Date_Time <= @date_time_end " +
                                "GROUP BY F.RID,F.Factory_ShortName_CN,STI.TYPE,STI.AFFINITY,STI.PHOTO,MCT.Type_Name,CT.NAME,CTO.TYPE,CTO.AFFINITY,CTO.PHOTO,CTO.NAME ";
        //(���X)
        public const string SEL_SYS_MOVEOUT_STOCKS = " SELECT F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME,SUM(CTSM.Move_Number) as Number " +
                                "FROM CARDTYPE_STOCKS_MOVE CTSM INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CTSM.CardType_RID = CT.RID " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND CTSM.From_Factory_RID = F.RID " +
                                "WHERE CTSM.RST = 'A' AND CTSM.Move_Date >= @date_time_start AND CTSM.Move_Date<=@date_time_end " +
                                "GROUP BY F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME ";
        //(���J)
        public const string SEL_SYS_MOVEIN_STOCKS = " SELECT F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME,SUM(CTSM.Move_Number) as Number " +
                                "FROM CARDTYPE_STOCKS_MOVE CTSM INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CTSM.CardType_RID = CT.RID " +
                                "INNER JOIN FACTORY F ON F.RST = 'A' AND F.Is_Perso = 'Y' AND CTSM.To_Factory_RID = F.RID " +
                                "WHERE CTSM.RST = 'A' AND CTSM.Move_Date >= @date_time_start AND CTSM.Move_Date<=@date_time_end " +
                                "GROUP BY F.RID,F.Factory_ShortName_CN,CT.TYPE,CT.AFFINITY,CT.PHOTO,CT.NAME ";
        //(Perso�t�d�ثe�ѵ��l)
        //public const string SEL_PERSO_CARDTYPE_BEFORE_DATE_SURPLUS = " SELECT TOP 1 Stocks_Number " +
        //                        "FROM CARDTYPE_STOCKS CS INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CS.CardType_RID = CT.RID " +
        //                        "WHERE CS.RST = 'A' AND CS.Perso_Factory_RID = @perso_factory_rid AND CT.TYPE = @type AND CT.AFFINITY = @affinity AND CT.Photo = @photo " +
        //                        "ORDER BY Stock_Date DESC ";
        public const string SEL_PERSO_CARDTYPE_BEFORE_DATE_SURPLUS = " SELECT TOP 1 Stocks_Number " +
                         "FROM CARDTYPE_STOCKS CS INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND CS.CardType_RID = CT.RID " +
                         "WHERE CS.RST = 'A' AND CS.Perso_Factory_RID = @perso_factory_rid AND CT.TYPE = @type AND CT.AFFINITY = @affinity AND CT.Photo = @photo " +
                         " AND Stock_Date =@Stock_Date";
        // �����ӥd����
        public const string SEL_EXPRESSIONS_DEFINE = " SELECT Type_RID,CS.Status_Code,CS.Status_Name,ED.Operate " +
                            "FROM EXPRESSIONS_DEFINE ED INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND ED.Type_RID = CS.RID " +
                            "WHERE ED.RST = 'A' AND ED.Expressions_RID = 2 ";

        //// �d�جO�_�������d�ˬd
        //public const string CON_CARD_TYPE_GROUP = "SELECT COUNT(*) as Num " +
        //                        "FROM CARD_TYPE CT INNER JOIN GROUP_CARD_TYPE GCT ON GCT.RST = 'A' AND CT.RID = GCT.CardType_RID " +
        //                        "INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND GCT.Group_RID = CG.RID AND CG.Param_Code = '" + GlobalString.Parameter.Type + "'" +
        //                        "WHERE CT.RST = 'A' AND CG.Group_Name = '" + GlobalString.Virtual_Card_Group.virtual_card_group + "' " +
        //                        " AND CT.Type = @type AND CT.Affinity = @affinity AND CT.Photo = @photo ";

        #endregion �p��t�Ϊ��t�ӵ��l

        #region �i��鵲
        public const string SEL_SYS_IN_STOCKS_SURPLUS = " SELECT DS.RID,DS.Perso_Factory_RID,DS.Space_Short_RID,DS.Wafer_RID,DS.Income_Number " +
                            "FROM DEPOSITORY_STOCK DS " +
                            "WHERE DS.RST = 'A' AND DS.Income_Date >= @date_time_start AND DS.Income_Date<= @date_time_end ";

        public const string SEL_SYS_DEPOSITORY_RESTOCK_SURPLUS = " SELECT RID,Perso_Factory_RID,Space_Short_RID,Wafer_RID,Reincome_Number " +
                            "FROM DEPOSITORY_RESTOCK " +
                            "WHERE RST = 'A' AND Reincome_Date >= @date_time_start AND Reincome_Date<= @date_time_end";

        public const string SEL_SYS_RETURN_STOCKS_SURPLUS = " SELECT Stock_RID,Cancel_Number " +
                            "FROM DEPOSITORY_CANCEL " +
                            "WHERE RST = 'A' AND Cancel_Date >= @date_time_start AND Cancel_Date<= @date_time_end";

        public const string SEL_WAFER_CARDTYPE_USELOG_RID = " SELECT RID " +
                            "FROM WAFER_CARDTYPE_USELOG " +
                            "WHERE Operate_Type = '1' AND Operate_RID IN ( SELECT RID FROM DEPOSITORY_STOCK WHERE RST = 'A' AND Stock_RID = @stock_rid) ";

        public const string CON_WAFER_USELOG_ROLLBACK = " SELECT COUNT(*) " +
                            "FROM WAFER_USELOG_ROLLBACK " +
                            "WHERE RST = 'A' AND UseLog_RID = @uselog_rid AND Check_Date >= @check_date_start AND Check_Date <= @check_date_end ";

        public const string INSERT_WAFER_USELOG_ROLLBACK = " INSERT INTO WAFER_USELOG_ROLLBACK (RCU,RUU,RCT,RUT,RST,Income_Date,Usable_Number,Factory_RID,CardType_RID,Begin_Date,End_Date,Wafer_RID,Operate_RID,Operate_Type,UseLog_RID,Check_Date,CardType_Move_RID,Number,BackUp_Date,unit_Price) " +
                    "SELECT '1','1',getdate(),getdate(),'A',Income_Date,Usable_Number,Factory_RID,CardType_RID,Begin_Date,End_Date,Wafer_RID,Operate_RID,Operate_Type,RID,@check_date,CardType_Move_RID,Number,@check_date,unit_Price " +
                        "FROM WAFER_CARDTYPE_USELOG " +
                        "WHERE RID = @uselog_rid ";

        public const string UPDATE_WAFER_CARDTYPE_USELOG = " UPDATE WAFER_CARDTYPE_USELOG " +
                    "SET Usable_Number = Usable_Number - @cancel_number,Number = Number - @cancel_number , begin_date = case when year(begin_date) = 1900 then @check_date else begin_date end " +
                    "WHERE RID = @rid ";

        public const string SEL_WAFER_CARDTYPE_USELOG = " SELECT Factory_RID,CardType_RID,RID,Usable_Number,Operate_RID,Operate_Type,Wafer_RID,Begin_Date,Income_Date,Unit_Price " +
                        "FROM WAFER_CARDTYPE_USELOG " +
                        "WHERE RST = 'A' AND Usable_Number>0 " +
                        "ORDER BY Factory_RID,CARDTYPE_RID,Income_Date , Rid ";

        //���USELOG��ROLLBACK�ɡA����²��P�_�O�_�i�μƶq�j��0�A�Ĥ@����0�ɡA���ӭn���ROLLBACK��
        public const string SEL_WAFER_CARDTYPE_USELOG_FIRST_ZERO = " SELECT Factory_RID,CardType_RID,RID,Usable_Number,Operate_RID,Operate_Type,Wafer_RID,Begin_Date,Income_Date,Unit_Price     " +
                        "FROM WAFER_CARDTYPE_USELOG U " +
                        "WHERE RST = 'A' AND  not exists ( select * from WAFER_USELOG_ROLLBACK W where w.uselog_rid = U.rid and W.usable_number = 0 )  " +
                        "ORDER BY Factory_RID,CARDTYPE_RID,Income_Date ";

        //public const string SEL_USE_CARDTYPE = " SELECT FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID,SUM(Number) as Number " +
        //                    "FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO " +
        //                    "WHERE FCI.RST = 'A' AND Date_Time >= @date_time_start AND Date_Time<= @date_time_end " +
        //                    "GROUP BY FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID ";
        public const string SEL_USE_CARDTYPE = " SELECT A.Perso_Factory_RID,A.RID,A.Status_RID,SUM(A.Number) as Number FROM ( " +
                        " SELECT FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID,SUM(Number) as Number FROM FACTORY_CHANGE_IMPORT FCI " +
                        "INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO " +
                       "WHERE FCI.RST = 'A' AND Date_Time >= @date_time_start AND Date_Time<= @date_time_end  and FCI.Status_RID not in ('1','2','3','4')" +
                       "GROUP BY FCI.Perso_Factory_RID,CT.RID,FCI.Status_RID " +
                       "union " +
                         "select SI.Perso_Factory_Rid,CT.RID,CS1.RID as Status_RID,SUM(Number) as Number  from SUBTOTAL_IMPORT  SI " +
                         "INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO " +
                         "INNER JOIN MAKE_CARD_TYPE M ON M.RST='A' AND M.RID=SI.MakeCardType_RID AND M.Is_Import='Y' " +
                         "INNER JOIN CARDTYPE_STATUS CS1 ON CS1.RST='A' AND CS1.Status_Name=M.Type_Name  " +
                         "where SI.RST = 'A' AND SI.Date_Time >= @date_time_start AND SI.Date_Time<= @date_time_end and M.Type_Name IN ('3D','DA','PM','RN') " +
                         "GROUP BY SI.Perso_Factory_RID,CT.RID,CS1.RID  " +
                         ") A  " +
                         "GROUP BY A.Perso_Factory_RID,A.RID,A.Status_RID " +
                        " order by A.Perso_Factory_RID,A.RID,A.Status_RID ";


        public const string UPDATE_WAFER_CARDTYPE_USELOG_1 = " UPDATE WAFER_CARDTYPE_USELOG SET Usable_Number = Usable_Number - @number , begin_date = case when year(begin_date) = 1900 then @check_date else begin_date end WHERE RID = @rid ";

        public const string UPDATE_WAFER_CARDTYPE_USELOG_2 = " UPDATE WAFER_CARDTYPE_USELOG SET Usable_Number = 0,End_Date = @check_date , begin_date = case when year(begin_date) = 1900 then @check_date else begin_date end WHERE RID = @rid ";

        public const string SEL_CARD_TYPE_MOVE_SURPLUS = " SELECT * " +
                            "FROM CARDTYPE_STOCKS_MOVE " +
                            "WHERE RST = 'A' AND Move_Date >= @check_date_start AND Move_Date<=@check_date_end ";

        public const string SEL_MATERIAL_BY_SUBTOTAL = " SELECT EI.Serial_Number AS EI_Number,CE.Serial_Number as CE_Number,A.Perso_Factory_RID,A.Number1 " +
                        "FROM (SELECT SI.Perso_Factory_RID,CT.RID,SUM(Number) AS Number1 " +
                        "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO " +
                            "WHERE SI.RST = 'A' AND SI.Date_Time >= @check_date_start AND SI.Date_Time <= @check_date_end " +
                            "GROUP BY SI.Perso_Factory_RID,CT.RID ) A " +
                        "INNER JOIN CARD_TYPE CT1 ON CT1.RST = 'A' AND A.RID = CT1.RID " +
                        "left JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND CT1.Envelope_RID = EI.RID " +
                        "left JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND CT1.Exponent_RID = CE.RID ";
        public const string SEL_MATERIAL_BY_SUBTOTAL_DM = " SELECT DI.Serial_Number DI_Number,A.Perso_Factory_RID,A.Number1,DI.Card_Type_Link_Type,DCT.CardType_RID"
                        + " FROM (SELECT SI.Perso_Factory_RID,CT.RID,SUM(Number) AS Number1,si.MakeCardType_rid"
                        + " FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO "
                        + " WHERE SI.RST = 'A' AND SI.Date_Time >=  @check_date_start "
                        + " AND SI.Date_Time <= @check_date_end "
                        + " GROUP BY SI.Perso_Factory_RID,CT.RID,si.MakeCardType_rid ) A "
                        + " inner join DM_MAKECARDTYPE DMM on DMM.MakeCardType_RID=a.MakeCardType_rid"
                        + " inner join DMTYPE_INFO DI on DI.rid=DMM.DM_RID"
                        + " left join  DM_CARDTYPE DCT ON DCT.RST = 'A' AND A.RID = DCT.CardType_RID and DCT.DM_RID=DI.rid"; 


        public const string SEL_SUBTOTAL_PROJECT_COST = " SELECT SI.Perso_Factory_RID,CT.RID,CG.RID AS CARDGROUPRID,SI.Number " +
                            "FROM SUBTOTAL_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO " +
                            "INNER JOIN GROUP_CARD_TYPE GCT ON GCT.RST = 'A' AND CT.RID = GCT.CardType_RID " +
                            "INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND GCT.Group_RID = CG.RID AND Param_Code = '" + GlobalString.Parameter.Finance + "' " +
                            "WHERE SI.RST = 'A' AND SI.Date_Time >= @check_date_start AND SI.Date_Time <= @check_date_end ";
      
        //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 start
        public const string SEL_SUBTOTAL_REPLACE_PROJECT_COST = "SELECT Perso_Factory_RID,RID,CARDGROUPRID,sum(Number) as  Number from ( " +
             " SELECT SI.Perso_Factory_RID,CT.RID,CG.RID AS CARDGROUPRID,SI.Number " +
             " FROM SUBTOTAL_REPLACE_IMPORT SI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND SI.TYPE = CT.TYPE AND SI.AFFINITY = CT.AFFINITY AND SI.PHOTO = CT.PHOTO " +
             " INNER JOIN GROUP_CARD_TYPE GCT ON GCT.RST = 'A' AND CT.RID = GCT.CardType_RID " +
             " INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND GCT.Group_RID = CG.RID AND Param_Code = '" + GlobalString.Parameter.Finance + "' " +
             " WHERE SI.RST = 'A' AND SI.Date_Time >= @check_date_start AND SI.Date_Time <= @check_date_end " +
             " union all " +
             " SELECT FCRI.Perso_Factory_RID,CT.RID,CG.RID AS CARDGROUPRID,Case FCRI.Status_RID when '5' then 0-FCRI.Number when '6' then 0-FCRI.Number when '7' then FCRI.Number end as Number " +
             " FROM FACTORY_CHANGE_REPLACE_IMPORT FCRI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCRI.TYPE = CT.TYPE AND FCRI.AFFINITY = CT.AFFINITY AND FCRI.PHOTO = CT.PHOTO " +
             " INNER JOIN GROUP_CARD_TYPE GCT ON GCT.RST = 'A' AND CT.RID = GCT.CardType_RID " +
             " INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND GCT.Group_RID = CG.RID AND Param_Code = '" + GlobalString.Parameter.Finance + "' " +
             " WHERE FCRI.RST = 'A'AND FCRI.Status_RID in ('5','6','7')  AND FCRI.Date_Time >= @check_date_start AND FCRI.Date_Time <= @check_date_end " +
             " ) A " +
             " Group by Perso_Factory_RID,RID,CARDGROUPRID";
        //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 end
        public const string SEL_PROJECT_STEP_SURPLUS = "SELECT PP.RID,PPP.Price " +
                     "FROM CARDTYPE_PERSO_PROJECT CPP " +
                     "INNER JOIN CARDTYPE_PROJECT_TIME CPT ON CPT.RST  = 'A' AND CPP.ProjectTime_RID = CPT.RID " +
                     "INNER JOIN PERSO_PROJECT PP ON PP.RST = 'A' AND CPT.PersoProject_RID = PP.RID AND PP.Normal_Special = '1' " +
                     "INNER JOIN PERSO_PROJECT_PRICE PPP ON PPP.RST = 'A' AND CPT.PersoProject_RID = PPP.Perso_Project_RID " +
                 "WHERE CPP.RST = 'A' AND CPP.CardType_RID = @CTRID " +
                     "AND CPT.Use_Date_Begin<=@Date_Time AND CPT.Use_Date_End>=@Date_Time " +
                     "AND PPP.Use_Date_Begin<=@Date_Time AND PPP.Use_Date_End>=@Date_Time " +
                     "AND PP.Factory_RID = @perso_factory_rid ";

        public const string SEL_SPECIAL_PROJECT_COST = " SELECT SUM(PP.Unit_Price*SPPI.Number) " +
                            "FROM SPECIAL_PERSO_PROJECT_IMPORT SPPI INNER JOIN PERSO_PROJECT PP ON PP.RST = 'A' AND SPPI.PersoProject_RID = PP.RID " +
                            "WHERE SPPI.RST = 'A' AND YEAR(SPPI.Project_Date) = @year ";

        public const string SEL_EXCEPTION_PROJECT_COST = " SELECT CardGroup_RID,SUM(Unit_Price*Number),Group_Name " +
                            "FROM EXCEPTION_PERSO_PROJECT EPP INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND EPP.CardGroup_RID = CG.RID " +
                            "WHERE EPP.RST = 'A' AND YEAR(Project_Date) = @year " +
                            "GROUP BY CardGroup_RID,Group_Name ";

        public const string SEL_PERSO_PROJECT_CHANGE_DETAIL = " SELECT CardGroup_RID,SUM(Price),Group_Name " +
                            "FROM PERSO_PROJECT_CHANGE_DETAIL PPC INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND PPC.CardGroup_RID = CG.RID " +
                            "WHERE PPC.RST = 'A' AND YEAR(Project_Date) = @year " +
                            "GROUP BY CardGroup_RID,Group_Name ";

        public const string SEL_PERSO_PROJECT_NORMAL = " SELECT Card_Group_RID,SUM(Sum),Group_Name " +
                            "FROM PERSO_PROJECT_DETAIL PPD INNER JOIN CARD_GROUP CG ON CG.RST = 'A' AND PPD.Card_Group_RID = CG.RID " +
                            "WHERE YEAR(Use_Date) = @year " +
                            "GROUP BY Card_Group_RID,Group_Name ";
        // �鵲�ɼаO
        public const string SEL_MATERIEL_BUDGET_SUM_CARD = " SELECT Budget FROM MATERIEL_BUDGET WHERE RST = 'A' AND Budget_Year = @year AND Materiel_Type = '9' ";

        public const string DEL_MAKE_COST_FROM_SUBTOTAL_IMPORT = "DELETE FROM PERSO_PROJECT_DETAIL " +
                            "WHERE RST = 'A' AND Use_Date>=@Begin_Date AND Use_Date <= @Finish_Date ";

        public const string SEL_MATERIEL_BUDGET_SUM_BANK = " SELECT Budget FROM MATERIEL_BUDGET WHERE RST = 'A' AND Budget_Year = @year AND Materiel_Type = '10' ";

        public const string UPDATE_DEPOSITORY_STOCK = " UPDATE DEPOSITORY_STOCK SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Income_Date >= @check_date_start AND Income_Date <= @check_date_end";

        public const string UPDATE_DEPOSITORY_RESTOCK = " UPDATE DEPOSITORY_RESTOCK SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Reincome_Date >= @check_date_start AND Reincome_Date <= @check_date_end";

        public const string UPDATE_DEPOSITORY_CANCEL = " UPDATE DEPOSITORY_CANCEL SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Cancel_Date >= @check_date_start AND Cancel_Date <= @check_date_end ";

        public const string UPDATE_SUBTOTAL_IMPORT = " UPDATE SUBTOTAL_IMPORT SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";

        public const string UPDATE_FACTORY_CHANGE_IMPORT = " UPDATE FACTORY_CHANGE_IMPORT SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";
        //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 start
        public const string UPDATE_SUBTOTAL_REPLACE_IMPORT = " UPDATE SUBTOTAL_REPLACE_IMPORT SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";
        
        public const string UPDATE_FACTORY_CHANGE_REPLACE_IMPORT = " UPDATE FACTORY_CHANGE_REPLACE_IMPORT SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Date_Time >= @check_date_start AND Date_Time <= @check_date_end ";
        //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 end
        public const string UPDATE_CARDTYPE_STOCKS_MOVE = " UPDATE CARDTYPE_STOCKS_MOVE SET Is_Check = 'Y',Check_Date = @check_date WHERE RST = 'A' AND Move_Date >= @check_date_start AND Move_Date <= @check_date_end ";

        public const string INSERT_CARDTYPE_STOCKS = " INSERT INTO CARDTYPE_STOCKS(Stock_Date,Stocks_Number,Perso_Factory_RID,CardType_RID) " +
                            "SELECT FCI.Date_Time,FCI.Number,FCI.Perso_Factory_RID,CT.RID " +
                            "FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO " +
                                "INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID " +
                            "WHERE FCI.RST  = 'A' AND FCI.Date_Time >= @check_date_start AND FCI.Date_Time <= @check_date_end AND CS.Status_Name = '�t�ӵ��l' ";

        public const string SEL_CARDTYPE_STOCKS = " SELECT FCI.Date_Time,FCI.Number,FCI.Perso_Factory_RID,CT.RID " +
                        "FROM FACTORY_CHANGE_IMPORT FCI INNER JOIN CARD_TYPE CT ON CT.RST = 'A' AND FCI.TYPE = CT.TYPE AND FCI.AFFINITY = CT.AFFINITY AND FCI.PHOTO = CT.PHOTO " +
                        "INNER JOIN CARDTYPE_STATUS CS ON CS.RST = 'A' AND FCI.Status_RID = CS.RID " +
                            "WHERE FCI.RST  = 'A' AND FCI.Date_Time >= @check_date_start AND FCI.Date_Time <= @check_date_end AND CS.Status_Name = '�t�ӵ��l'";


        #endregion
        public const string SEL_SURPLUS_CHECK = "SELECT TOP 1 Stock_Date FROM CARDTYPE_STOCKS "
                         + " ORDER BY Stock_Date DESC";

        public const string SEL_WORKDATE_NOT_SURPLUS = "SELECT Date_Time FROM WORK_DATE "
                         + " WHERE RST = 'A' AND Date_Time > @lasttime AND  Date_Time <= @now and is_workday='Y' Order by Date_Time ";

        public const string SEL_LAST_WORK_DATE = "SELECT TOP 1 Date_Time " +
                    "FROM WORK_DATE " +
                    "WHERE Date_Time < @date_time AND Is_WorkDay='Y' " +
                    "ORDER BY Date_Time DESC";

        public const string SEL_MATERIEL_STOCKS_MANAGER = "SELECT MSM.Perso_Factory_RID,MSM.Serial_Number,MSM.Number," +
                            "CASE SUBSTRING(MSM.Serial_Number,1,1) WHEN 'A' THEN EI.NAME WHEN 'B' THEN CE.NAME WHEN 'C' THEN DI.NAME END AS NAME " +
                        "FROM MATERIEL_STOCKS_MANAGE MSM LEFT JOIN ENVELOPE_INFO EI ON EI.RST = 'A' AND MSM.Serial_Number = EI.Serial_Number " +
                            "LEFT JOIN CARD_EXPONENT CE ON CE.RST = 'A' AND MSM.Serial_Number = CE.Serial_Number " +
                            "LEFT JOIN DMTYPE_INFO DI ON DI.RST = 'A' AND MSM.Serial_Number = DI.Serial_Number " +
                        "WHERE Convert(varchar(10),Stock_Date,111) = @stock_date AND (Type = 4 or Type = 5) " +
                            "AND MSM.Serial_Number IN (";

        public const string SEL_MATERIEL_USED = "SELECT sum(Number) " +
                        "FROM MATERIEL_STOCKS_USED " +
                        "WHERE RST = 'A' AND Perso_Factory_RID = @perso_factory_rid AND " +
                            "Serial_Number = @serial_number AND " +
                            "Convert(varchar(10),Stock_Date,111) = @stock_date ";
        public const string DEL_MATERIEL_STOCKS_USED = " Delete From MATERIEL_STOCKS_USED where Stock_Date >= @check_date_start AND Stock_Date <= @check_date_end";

        public const string SEL_ENVELOPE_INFO = "SELECT * "
                                    + "FROM ENVELOPE_INFO "
                                    + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_CARD_EXPONENT = "SELECT * "
                                        + "FROM CARD_EXPONENT "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_DMTYPE_INFO = "SELECT * "
                                        + "FROM DMTYPE_INFO "
                                        + "WHERE RST = 'A' AND Serial_Number = @serial_number";
        public const string SEL_MATERIEL_STOCKS_USED = "select * from MATERIEL_STOCKS_USED where rst='A' "
                                        + "AND Serial_Number=@Serial_Number "
                                        + "AND Perso_Factory_RID=@Perso_Factory_RID "
                                        + "AND Stock_Date > @lastSurplusDateTime "
                                        + "AND Stock_Date <= @thisSurplusDateTime";

        //�J�w�������Ѫ��A�����ө�JROLLBACK�ɡA�ҥH�ҥH�s�W������A�[�@�B�R���ާ@�I
        public const string DEL_WAFER_USELOG_ROLLBACK_TODAY = " DELETE FROM WAFER_USELOG_ROLLBACK WHERE income_date = @check_date_start";


        Dictionary<string, object> dirValues = new Dictionary<string, object>();

        public string strErr;

        #endregion
        /// <summary>
        /// �鵲�e�A��t�Ӯw�s���ʸ�T�M�t�ήw�s���ʸ�T����C
        /// </summary>
        /// <param name="dtSurplus">�鵲���</param>
        /// <param name="dtStockDiff">��T���۲Ŷ�</param>
        public bool Compare(DateTime dtSurplus)
        {
            DataSet dsFci = new DataSet();
            DataSet dsSys = new DataSet();
            bool flag = true;
            strErr = "";
            try
            {
                // ���t�Ӯw�s���ʸ�T�M�t�ήw�s��T
                getFactorySysStockNumber(dtSurplus, ref dsFci, ref dsSys);               

                // ����t�Ӯw�s���ʸ�T�M�t�ήw�s��T
                object[] arg = new object[3];
                arg[0] = dtSurplus.ToString("yyyy/MM/dd");
                flag = CompareFactorySys(dsFci, dsSys, arg);

                // {0} {1}perso�t {2}����²�ٮֹ靈�~�A����鵲
                if (!flag)
                    Warning.SetWarning(GlobalString.WarningType.BatchCompareNotPass, arg);

               
                ////200908CR�����e�Z�t�Ӳ����ɤ�� add by ���� 2009/09/03 start  //20090922CR�s�W�ݨD�������B�޿�
                //CIMSClass.Business.InOut007BL BL007 = new CIMSClass.Business.InOut007BL();
                //DataTable dtFactoryReplace = new DataTable();
                //object[] arg1 = new object[1];
                //arg1[0] = dtSurplus.ToString("yyyy/MM/dd");
                //BL007.GetCompareFactoryReplace(dtSurplus, ref dtFactoryReplace);
                ////�p�s�b�t���ơA����鵲
                //if (dtFactoryReplace.Rows.Count > 0)
                //{
                //    Warning.SetWarning(GlobalString.WarningType.BatchCompareFactoryReplace, arg1);
                //    flag = false;
                //}
                ////200908CR�����e�Z�t�Ӳ����ɤ�� add by ���� 2009/09/03 start
                return flag;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�鵲�e�A��t�Ӯw�s���ʸ�T�M�t�ήw�s���ʸ�T����C, Compare����: " + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="dtSurplus">�鵲���</param>
        /// <param name="dtLastSurplus">�W���鵲���</param>
        /// <returns></returns>
        public DataTable getAllShouldSurplusCardType(DateTime dtSurplus,
                                    DateTime dtLastSurplus)
        {
            DataTable dtRet = null;
            try
            {
                this.dirValues.Clear();
                this.dirValues.Add("date_time_start", dtSurplus.ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("date_time_end", dtSurplus.ToString("yyyy/MM/dd 23:59:59"));
                this.dirValues.Add("stock_date_start", dtLastSurplus.ToString("yyyy/MM/dd 00:00:00"));
                this.dirValues.Add("stock_date_end", dtLastSurplus.ToString("yyyy/MM/dd 23:59:59"));
                DataSet dsAllShouldSurplusCardType = dao.GetList(SEL_ALL_SHOULD_SURPLUS_CARDTYPE,
                                            this.dirValues);
                if (dsAllShouldSurplusCardType != null
                    && dsAllShouldSurplusCardType.Tables.Count > 0
                    && dsAllShouldSurplusCardType.Tables[0].Rows.Count > 0)
                {
                    dtRet = dsAllShouldSurplusCardType.Tables[0];
                }

                return dtRet;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("getAllShouldSurplusCardType����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// ���̦Z�@���鵲���
        /// </summary>
        /// <returns></returns>
        public DateTime getLastSurplusDate()
        {
            DateTime dtLastSurplusDate = Convert.ToDateTime("1900-01-01");
            try
            {
                DataSet dsLAST_SURPLUS_DAY = dao.GetList(SEL_LAST_SURPLUS_DAY);
                if (dsLAST_SURPLUS_DAY != null
                        && dsLAST_SURPLUS_DAY.Tables.Count > 0
                        && dsLAST_SURPLUS_DAY.Tables[0].Rows.Count > 0)
                {
                    dtLastSurplusDate = Convert.ToDateTime(dsLAST_SURPLUS_DAY.Tables[0].Rows[0]["Stock_Date"].ToString());
                }
                return dtLastSurplusDate;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���̦Z�@���鵲���, getLastSurplusDate����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// ���t�Ӯw�s���ʸ�T�M�t�ήw�s��T
        /// </summary>
        /// <param name="Date">�鵲���</param>
        /// <param name="dsFACTORY_CHANGE_IMPORT">�t�Ӯw�s���ʸ�T</param>
        /// <param name="dsSys_Stock">�t�ήw�s���ʸ�T</param>
        public void getFactorySysStockNumber(DateTime Date,
                ref DataSet dsFACTORY_CHANGE_IMPORT,
                ref DataSet dsSys_Stock)
        {
            try
            {
                //���t�����Ӥ鵲��Perso�t�B�d�ذT��
                string sLastDate = GetLastStock_Date();
                DataTable dtShouldSurplusCardType;
                if (sLastDate != null && sLastDate != "")
                {
                    DateTime dtLastDate = Convert.ToDateTime(sLastDate);

                    dtShouldSurplusCardType = getAllShouldSurplusCardType(Date, dtLastDate);
                    if (null == dtShouldSurplusCardType)
                    {
                        throw new Exception("�S���ݭn�鵲��Perso�t�d�ذT���C");
                    }

                }
                else
                {
                    throw new Exception("�����e�@�u�@��w�s�H���C");
                }


                #region �����ӥd����
                this.dirValues.Clear();
                DataSet dsEXPRESSIONS = dao.GetList(SEL_EXPRESSIONS_DEFINE);
                if (!(null != dsEXPRESSIONS &&
                        dsEXPRESSIONS.Tables.Count > 0 &&
                        dsEXPRESSIONS.Tables[0].Rows.Count > 0))
                {
                    strErr += "�t�Ϊ����ӥd���������T�A�и�޲z���T�{�C";
                }
                #endregion

                #region �t�Ӯw�s���ʶפJ��Perso�t�B�d�ءB�d�ت��p���`�C
                dirValues.Clear();
                dirValues.Add("date_time_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("date_time_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                dsFACTORY_CHANGE_IMPORT = dao.GetList(SEL_FACTORY_IMPORT_STOCKS, dirValues);
                #endregion �t�Ӯw�s����

                if (dsFACTORY_CHANGE_IMPORT.Tables[0].Rows.Count == 0)
                    throw new Exception(Date.ToShortDateString()+"�d���w�s�����ɥ��פJ���\�A�t�Τ�����鵲�@�~");

                // ���t�μt�ӵ��lDataTable
                DataTable dtSys_Stock_Surplus = getDTSys_Stock();

                #region ���t�ήw�s���ʸ�T
                // ���t�ΤJ�w�B�h�f�B�A�J�w�B(3D�BDA�BPM�BRN)�B���X�B���J��T
                DataSet dsSYSTEM_FACTORY_STOCK = dao.GetList(SEL_SYS_IN_STOCKS +
                                    SEL_SYS_RETURN_STOCKS +
                                    SEL_SYS_DEPOSITORY_RESTOCK +
                                    SEL_SYS_SUBTOTAL_TYPE +
                                    SEL_SYS_MOVEOUT_STOCKS +
                                    SEL_SYS_MOVEIN_STOCKS, dirValues);
                // �p��t�Ϊ��t�ӵ��l
                // �p���k�G�t�Ϊ��e�ѵ��l + �J�w - �h�f + �A�J�w - ���� + ���J - ���X
                //�W�@�鵲���
                DateTime LastSurplusDate = getLastSurplusDate();
                //            foreach (DataRow drFactoryChangeImport in dsFACTORY_CHANGE_IMPORT.Tables[0].Rows)
                foreach (DataRow drFactoryChangeImport in dtShouldSurplusCardType.Rows)
                {
                    //if (drFactoryChangeImport["Status_Name"].ToString() == "�t�ӵ��l")
                    //{
                    int intPersoCardTypeBeforeDateSurplus = 0;
                    #region ���t�θӥd�ءBPerso�t�e�@�Ѫ����l
                    this.dirValues.Clear();
                    dirValues.Add("perso_factory_rid", drFactoryChangeImport["Perso_Factory_RID"].ToString());
                    dirValues.Add("type", drFactoryChangeImport["TYPE"].ToString());
                    dirValues.Add("affinity", drFactoryChangeImport["AFFINITY"].ToString());
                    dirValues.Add("photo", drFactoryChangeImport["PHOTO"].ToString());
                    dirValues.Add("Stock_Date", LastSurplusDate);
                    DataSet dsPersoCardTypeLastSurplus = dao.GetList(SEL_PERSO_CARDTYPE_BEFORE_DATE_SURPLUS, dirValues);
                    if (null != dsPersoCardTypeLastSurplus &&
                        dsPersoCardTypeLastSurplus.Tables.Count > 0 &&
                        dsPersoCardTypeLastSurplus.Tables[0].Rows.Count > 0)
                    {
                        intPersoCardTypeBeforeDateSurplus = Convert.ToInt32(dsPersoCardTypeLastSurplus.Tables[0].Rows[0]["Stocks_Number"].ToString());

                        if (intPersoCardTypeBeforeDateSurplus < 0)
                        {
                            intPersoCardTypeBeforeDateSurplus = 0;
                        }

                    }
                    #endregion ���t�θӥd�ءBPerso�t�e�@�Ѫ����l

                    int intPersoCardTypeUsedNumber = 0;
                    #region �̮��ӥd�����A�p���Perso�t�B�d�ت����ӥd��
                    for (int intLoop = 0; intLoop < dsEXPRESSIONS.Tables[0].Rows.Count; intLoop++)
                    {
                        if (dsEXPRESSIONS.Tables[0].Rows[intLoop]["Operate"].ToString() != "��")
                        {
                            DataRow[] drUsed = null;
                            DataRow[] drUsedIn = null;
                            string strStatus_Name = dsEXPRESSIONS.Tables[0].Rows[intLoop]["Status_Name"].ToString();
                            switch (strStatus_Name.ToUpper())
                            {
                                case "3D":
                                    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[3].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Type_Name = '3D'");
                                    break;
                                case "DA":
                                    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[3].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Type_Name = 'DA'");
                                    break;
                                case "PM":
                                    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[3].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Type_Name = 'PM'");
                                    break;
                                case "RN":
                                    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[3].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Type_Name = 'RN'");
                                    break;
                                case "�˥d":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�˥d'");
                                    break;
                                case "���s�d":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '���s�d'");
                                    break;
                                case "�ɻs�d":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�ɻs�d'");
                                    break;
                                case "�s�l�d":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�s�l�d'");
                                    break;
                                case "�ƥd":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�ƥd'");
                                    break;
                                case "�P�����}":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�P�����}'");
                                    break;
                                case "�ʥd":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�ʥd'");
                                    break;
                                case "�P��":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�P��'");
                                    break;
                                case "�վ�":
                                    drUsed = dsFACTORY_CHANGE_IMPORT.Tables[0].Select("Perso_Factory_RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString() +
                                                               " AND Status_Name = '�վ�'");
                                    break;
                                //case "����":
                                //    // ���X
                                //    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[4].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                //                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                //                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                //                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                                //    // ���J
                                //    drUsedIn = dsSYSTEM_FACTORY_STOCK.Tables[5].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                //                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                //                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                //                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                                //    break;
                                //case "�J�w":
                                //    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[0].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                //                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                //                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                //                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                                //    break;
                                //case "�h�f":
                                //    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[1].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                //                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                //                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                //                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                                //    break;
                                //case "�A�J�w":
                                //    drUsed = dsSYSTEM_FACTORY_STOCK.Tables[2].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                //                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                //                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                //                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                                //    break;
                            }

                            if (drUsed == null && drUsedIn == null)
                                continue;

                            // ���઺�p��
                            if (strStatus_Name == "����")
                            {
                                //if (dsEXPRESSIONS.Tables[0].Rows[intLoop]["Operate"].ToString() == "+")
                                //{
                                //    if (drUsedIn.Length > 0)
                                //    {
                                //        intPersoCardTypeUsedNumber -= Convert.ToInt32(drUsedIn[0]["Number"].ToString());
                                //    }
                                //    if (drUsed.Length > 0)
                                //    {
                                //        intPersoCardTypeUsedNumber += Convert.ToInt32(drUsed[0]["Number"].ToString());
                                //    }
                                //}
                                //else
                                //{
                                //    if (drUsedIn.Length > 0)
                                //    {
                                //        intPersoCardTypeUsedNumber += Convert.ToInt32(drUsedIn[0]["Number"].ToString());
                                //    }
                                //    if (drUsed.Length > 0)
                                //    {
                                //        intPersoCardTypeUsedNumber -= Convert.ToInt32(drUsed[0]["Number"].ToString());
                                //    }
                                //}
                            }
                            // ��L�D���઺�p��
                            else
                            {
                                if (dsEXPRESSIONS.Tables[0].Rows[intLoop]["Operate"].ToString() == "+")
                                {
                                    //  Legend 2017/05/24 �K�[����null�P�_
                                    if (drUsed != null)
                                    {
                                        if (drUsed.Length > 0)
                                        {
                                            for (int i = 0; i < drUsed.Length; i++)
                                            {
                                                intPersoCardTypeUsedNumber += Convert.ToInt32(drUsed[i]["Number"].ToString());
                                            }

                                        }
                                    }
                                }
                                else
                                {
                                    //  Legend 2017/05/24 �K�[����null�P�_
                                    if (drUsed != null)
                                    {
                                        if (drUsed.Length > 0)
                                        {
                                            for (int i = 0; i < drUsed.Length; i++)
                                            {
                                                intPersoCardTypeUsedNumber -= Convert.ToInt32(drUsed[i]["Number"].ToString());
                                            }

                                        }
                                    }
                                }
                                //}
                            }
                        }
                    }
                    #endregion �̮��ӥd�����A�p���Perso�t�B�d�ت����ӥd��

                    int intDepositoryInNumber = 0;
                    #region ��Perso�t�B�ӥd�ت��J�w�ƶq
                    DataRow[] drDepIn = dsSYSTEM_FACTORY_STOCK.Tables[0].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                    if (drDepIn.Length > 0)
                    {
                        intDepositoryInNumber = Convert.ToInt32(drDepIn[0]["Number"].ToString());
                    }
                    #endregion ��Perso�t�B�ӥd�ت��J�w�ƶq

                    int intDepositoryCancelNumber = 0;
                    #region ��Perso�t�B�ӥd�ت��h�f�ƶq
                    DataRow[] drDepCancel = dsSYSTEM_FACTORY_STOCK.Tables[1].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                    if (drDepCancel.Length > 0)
                    {
                        intDepositoryCancelNumber = Convert.ToInt32(drDepCancel[0]["Number"].ToString());
                    }
                    #endregion ��Perso�t�B�ӥd�ت��h�f�ƶq

                    int intDepositoryReInNumber = 0;
                    #region ��Perso�t�B�ӥd�ت��A�J�w�ƶq
                    DataRow[] drDepReIn = dsSYSTEM_FACTORY_STOCK.Tables[2].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                    if (drDepReIn.Length > 0)
                    {
                        intDepositoryReInNumber = Convert.ToInt32(drDepReIn[0]["Number"].ToString());
                    }
                    #endregion ��Perso�t�B�ӥd�ت��A�J�w�ƶq

                    int intDepositoryMoveNumber = 0;
                    #region ��Perso�t�B�ӥd�ت�����ƶq
                    DataRow[] drDepMoveOut = dsSYSTEM_FACTORY_STOCK.Tables[4].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                    if (drDepMoveOut.Length > 0)
                    {
                        // ��X�ƶq
                        intDepositoryMoveNumber -= Convert.ToInt32(drDepMoveOut[0]["Number"].ToString());
                    }

                    DataRow[] drDepMoveIn = dsSYSTEM_FACTORY_STOCK.Tables[5].Select("RID = " + drFactoryChangeImport["Perso_Factory_RID"].ToString() +
                                                               " AND Type = " + drFactoryChangeImport["Type"].ToString() +
                                                               " AND Affinity = " + drFactoryChangeImport["Affinity"].ToString() +
                                                               " AND Photo = " + drFactoryChangeImport["Photo"].ToString());
                    if (drDepMoveIn.Length > 0)
                    {
                        // ��J�ƶq
                        intDepositoryMoveNumber += Convert.ToInt32(drDepMoveIn[0]["Number"].ToString());
                    }
                    #endregion ��Perso�t�B�ӥd�ت�����ƶq

                    int intPersoCardTypeSurplus = 0;//��ѵ��l
                    // ��ѵ��l�ƶq = �e�鵲�l + �J�w - �h�f + �A�J�w - ���ӥd + 4.11�w�s����
                    intPersoCardTypeSurplus = intPersoCardTypeBeforeDateSurplus
                                            + intDepositoryInNumber - intDepositoryCancelNumber
                                            + intDepositoryReInNumber - intPersoCardTypeUsedNumber
                                            + intDepositoryMoveNumber;
                    // �K�[���l
                    DataRow drNewSurplus = dtSys_Stock_Surplus.NewRow();
                    drNewSurplus["Perso_Factory_RID"] = drFactoryChangeImport["Perso_Factory_RID"];
                    drNewSurplus["Factory_ShortName_CN"] = drFactoryChangeImport["Factory_ShortName_CN"];
                    drNewSurplus["TYPE"] = drFactoryChangeImport["TYPE"];
                    drNewSurplus["AFFINITY"] = drFactoryChangeImport["AFFINITY"];
                    drNewSurplus["PHOTO"] = drFactoryChangeImport["PHOTO"];
                    drNewSurplus["Name"] = drFactoryChangeImport["Name"];
                    drNewSurplus["Status_Name"] = "�t�ӵ��l";
                    drNewSurplus["Number"] = intPersoCardTypeSurplus;
                    // �N���l�T���K�[��t�μt�Ӯw�s���ʰT�����C
                    dtSys_Stock_Surplus.Rows.Add(drNewSurplus);
                }
                #endregion �t�ήw�s����

                if (null != dsSYSTEM_FACTORY_STOCK)
                {
                    dsSYSTEM_FACTORY_STOCK.Tables.Add(dtSys_Stock_Surplus);
                    dsSys_Stock = dsSYSTEM_FACTORY_STOCK;
                }
            }

            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���t�Ӯw�s���ʸ�T�M�t�ήw�s��T, getFactorySysStockNumber����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// ����t�Ӯw�s���ʩM�t�ήw�s���ʸ�T
        /// </summary>
        /// <param name="dsFactory_Stock_Number">�t�Ӯw�s���ʸ�T</param>
        /// <param name="dsSystem_Stock_Number">�t�ήw�s���ʸ�T</param>
        /// <param name="dtStockDiff">���ʸ�T���ŰO��</param>
        /// <param name="dtStockSurplusDiff">���ʸ�T���`���ŰO��</param>
        public bool CompareFactorySys(DataSet dsFactory_Stock_Number,
                                    DataSet dsSystem_Stock_Number,
                                    object[] argFun)
        {

            DataRow[] drUsed = null;
            DataRow[] drUsedMoveIn = null;
            //dtStockDiff = getDTStockDiff();
            
            #region �����ʸ�T���ŰO��
            foreach (DataRow drFactoryStockNumber in dsFactory_Stock_Number.Tables[0].Rows)
            {
                string strStatus_Name = drFactoryStockNumber["Status_Name"].ToString();
                if (strStatus_Name == "�J�w")
                {
                    drUsed = dsSystem_Stock_Number.Tables[0].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND Type = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
                }
                else if (strStatus_Name == "�h�f")
                {
                    drUsed = dsSystem_Stock_Number.Tables[1].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND Type = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
                }
                else if (strStatus_Name == "�A�J�w")
                {
                    drUsed = dsSystem_Stock_Number.Tables[2].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND Type = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
                }
                else if (strStatus_Name == "3D")
                {
                    drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND OLDType = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND OLDAffinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND OLDPhoto = " + drFactoryStockNumber["Photo"].ToString() +
                                                            " AND Type_Name = '3D'");
                }
                else if (strStatus_Name == "DA")
                {
                    drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND OLDType = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND OLDAffinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND OLDPhoto = " + drFactoryStockNumber["Photo"].ToString() +
                                                            " AND Type_Name = 'DA'");
                }
                else if (strStatus_Name == "PM")
                {
                    drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND OLDType = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND OLDAffinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND OLDPhoto = " + drFactoryStockNumber["Photo"].ToString() +
                                                            " AND Type_Name = 'PM'");
                }
                else if (strStatus_Name == "RN")
                {
                    drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND OLDType = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND OLDAffinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND OLDPhoto = " + drFactoryStockNumber["Photo"].ToString() +
                                                            " AND Type_Name = 'RN'");
                }
                else if (strStatus_Name == "����")
                {
                    drUsed = dsSystem_Stock_Number.Tables[4].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND Type = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
                    drUsedMoveIn = dsSystem_Stock_Number.Tables[5].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND Type = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
                }
                else if (strStatus_Name == "�t�ӵ��l")
                {
                    drUsed = dsSystem_Stock_Number.Tables[6].Select("Perso_Factory_RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
                                                            " AND Type = " + drFactoryStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
                }

                // �ˬd�ƶq�O�_�۵�
                if (strStatus_Name == "����")
                {
                    int intMoveNumber = 0;
                    if (drUsed != null)
                    {
                        if (drUsed.Length > 0)
                        {
                            intMoveNumber -= Convert.ToInt32(drUsed[0]["Number"].ToString());
                        }

                        // Legend 2017/05/24 �K�[�P�_��  drUsedMoveIn ����null��, �A�P�_����
                        if (drUsedMoveIn != null)
                        {
                            if (drUsedMoveIn.Length > 0)
                            {
                                intMoveNumber += Convert.ToInt32(drUsedMoveIn[0]["Number"].ToString());
                            }
                        }

                        // �ˬd����ƶq�O�_�۵�
                        if (intMoveNumber != Convert.ToInt32(drFactoryStockNumber["Number"].ToString()))
                        {
                            if (intMoveNumber != 0)
                            {
                                argFun[1] = drFactoryStockNumber["Factory_ShortName_CN"].ToString();
                                argFun[2] = drFactoryStockNumber["NAME"].ToString();
                                return false;
                            }
                        }
                    }
                    // �ˬd�J�w�B�h�f�B�A�J�w�B3D�BDA�BPM�BRN�B����B���鵲�l
                }
                else if (strStatus_Name == "�J�w" ||
                    strStatus_Name == "�h�f" ||
                    strStatus_Name == "�A�J�w" ||
                    strStatus_Name == "3D" ||
                    strStatus_Name == "DA" ||
                    strStatus_Name == "PM" ||
                    strStatus_Name == "RN" ||
                    strStatus_Name == "�t�ӵ��l")
                {
                    if (drUsed != null && drUsed.Length != 0)
                    {
                        int iNumber = 0;
                        for (int i = 0; i < drUsed.Length; i++)
                        {
                            iNumber += Convert.ToInt32(drUsed[i]["Number"].ToString());
                        }

                        if (iNumber != Convert.ToInt32(drFactoryStockNumber["Number"].ToString())) 
                        {
                            // �t�Ӯw�s���ʩM�t�ήw�s���ʤ��۲ŮɲK�[�O�s�ӰO��
                            //DataRow drStockDiff = dtStockDiff.NewRow();
                            //drStockDiff["Factory_ShortName_CN"] = drFactoryStockNumber["Factory_ShortName_CN"];
                            //drStockDiff["Name"] = drFactoryStockNumber["Name"];
                            //drStockDiff["Type"] = strStatus_Name;
                            //drStockDiff["Factory_Number"] = drFactoryStockNumber["Number"];
                            //drStockDiff["System_Number"] = Convert.ToInt32(drUsed[0]["Number"].ToString());
                            //dtStockDiff.Rows.Add(drStockDiff);
                            if (iNumber != 0)
                            {
                                argFun[1] = drFactoryStockNumber["Factory_ShortName_CN"].ToString();
                                argFun[2] = drFactoryStockNumber["NAME"].ToString();
                                return false;
                            }
                        }
                    }
                    //else
                    //{
                    //    // �p�G���l�q���t�ơC�������d�~�]�n��ܡC
                    //    if (Convert.ToInt32(drUsed[0]["Number"].ToString()) < 0 && strStatus_Name == "�t�ӵ��l")
                    //    {
                    //        this.dirValues.Clear();
                    //        this.dirValues.Add("type", drFactoryStockNumber["Type"].ToString());
                    //        this.dirValues.Add("affinity", drFactoryStockNumber["Affinity"].ToString());
                    //        this.dirValues.Add("photo", drFactoryStockNumber["Photo"].ToString());
                    //        DataSet dstVirtualCardGroup = dao.GetList(CON_CARD_TYPE_GROUP, this.dirValues);
                    //        if (null != dstVirtualCardGroup && dstVirtualCardGroup.Tables.Count > 0 &&
                    //            dstVirtualCardGroup.Tables[0].Rows.Count > 0)
                    //        {
                    //            if (Convert.ToInt32(dstVirtualCardGroup.Tables[0].Rows[0]["Num"].ToString()) == 0)
                    //            {   // ���������d
                    //                // �t�Ӯw�s���ʩM�t�ήw�s���ʤ��۲ŮɲK�[�O�s�ӰO��
                    //                DataRow drStockDiff = dtStockDiff.NewRow();
                    //                drStockDiff["Factory_ShortName_CN"] = drFactoryStockNumber["Factory_ShortName_CN"];
                    //                drStockDiff["Name"] = drFactoryStockNumber["Name"];
                    //                drStockDiff["Type"] = strStatus_Name;
                    //                drStockDiff["Factory_Number"] = drFactoryStockNumber["Number"];
                    //                drStockDiff["System_Number"] = Convert.ToInt32(drUsed[0]["Number"].ToString());
                    //                dtStockDiff.Rows.Add(drStockDiff);
                    //            }
                    //        }
                    //    }
                    //}
                }
            }

            #endregion �����ʸ�T���ŰO��


            #region ���t�ΰT�����s�b�A�t�Ӳ��ʤ��s�b�T��
            // �J�w
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[0].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["RID"].ToString() +
                                                            " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drSysStockNumber["Photo"].ToString() +
                                                            " AND Status_Name = '�J�w'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    if (drSysStockNumber["Number"].ToString() != "" && drSysStockNumber["Number"].ToString() != "0")
                    {
                        argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                        argFun[2] = drSysStockNumber["NAME"].ToString();
                        return false;
                    }

                }
            }

            // �h�f
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[1].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["RID"].ToString() +
                                                            " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drSysStockNumber["Photo"].ToString() +
                                                            " AND Status_Name = '�h�f'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    if (drSysStockNumber["Number"].ToString() != "" && drSysStockNumber["Number"].ToString() != "0")
                    {
                        argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                        argFun[2] = drSysStockNumber["NAME"].ToString();
                        return false;
                    }

                }
            }

            // �A�J�w
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[2].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["RID"].ToString() +
                                                            " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drSysStockNumber["Photo"].ToString() +
                                                            " AND Status_Name = '�A�J�w'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    if (drSysStockNumber["Number"].ToString() != "" && drSysStockNumber["Number"].ToString() != "0")
                    {
                        argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                        argFun[2] = drSysStockNumber["NAME"].ToString();
                        return false;
                    }
                }
            }

            // �p�p��
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[3].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["RID"].ToString() +
                                                " AND Type = " + drSysStockNumber["OLDType"].ToString() +
                                                " AND Affinity=" + drSysStockNumber["OLDAffinity"].ToString() +
                                                " AND Photo = " + drSysStockNumber["OLDPhoto"].ToString() +
                                                " AND Status_Name = '" + drSysStockNumber["Type_Name"].ToString() + "'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    if (drSysStockNumber["Number"].ToString() != "" && drSysStockNumber["Number"].ToString() != "0")
                    {
                        argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                        argFun[2] = drSysStockNumber["OLDNAME"].ToString();
                        return false;
                    }

                }
            }

            // ����]��X�^
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[4].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["RID"].ToString() +
                                                            " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drSysStockNumber["Photo"].ToString() +
                                                            " AND Status_Name = '����'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    // �p��t�β���ƶq
                    int intMove = 0;
                    DataRow[] drMoveIns = dsSystem_Stock_Number.Tables[5].Select("RID = " + drSysStockNumber["RID"].ToString() +
                                                           " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                           " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                           " AND Photo = " + drSysStockNumber["Photo"].ToString());

                    if (null != drMoveIns && drMoveIns.Length > 0)
                    {
                        intMove = Convert.ToInt32(drMoveIns[0]["Number"]) - Convert.ToInt32(drSysStockNumber["Number"]);
                    }
                    else
                    {
                        intMove -= Convert.ToInt32(drSysStockNumber["Number"]);
                    }

                    if (intMove != 0)
                    {
                        argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                        argFun[2] = drSysStockNumber["NAME"].ToString();
                        return false;
                    }
                }
            }

            // ����]��J�^
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[5].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["RID"].ToString() +
                                                            " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                            " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                            " AND Photo = " + drSysStockNumber["Photo"].ToString() +
                                                            " AND Status_Name = '����'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    // �p��t�β���ƶq
                    int intMove = 0;
                    DataRow[] drMoveOuts = dsSystem_Stock_Number.Tables[4].Select("RID = " + drSysStockNumber["RID"].ToString() +
                                                           " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                           " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                           " AND Photo = " + drSysStockNumber["Photo"].ToString());
                    if (null != drMoveOuts && drMoveOuts.Length > 0)
                    {
                        // ���X�ɤw�g�ˬd
                    }
                    else
                    {
                        intMove = Convert.ToInt32(drSysStockNumber["Number"]);
                        if (intMove != 0)
                        {
                            argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                            argFun[2] = drSysStockNumber["NAME"].ToString();
                            return false;
                        }
                    }








                }
            }

            // �t�ӵ��E
            foreach (DataRow drSysStockNumber in dsSystem_Stock_Number.Tables[6].Rows)
            {
                drUsed = dsFactory_Stock_Number.Tables[0].Select("Perso_Factory_RID = " + drSysStockNumber["Perso_Factory_RID"].ToString() +
                                                        " AND Type = " + drSysStockNumber["Type"].ToString() +
                                                        " AND Affinity = " + drSysStockNumber["Affinity"].ToString() +
                                                        " AND Photo = " + drSysStockNumber["Photo"].ToString() +
                                                        " AND Status_Name = '�t�ӵ��l'");
                if (drUsed != null && drUsed.Length == 0)
                {
                    if (drSysStockNumber["Number"].ToString() != "" && drSysStockNumber["Number"].ToString() != "0")
                    {
                        argFun[1] = drSysStockNumber["Factory_ShortName_CN"].ToString();
                        argFun[2] = drSysStockNumber["NAME"].ToString();
                        return false;
                    }
                }
            }

            return true;

            #endregion ���t�ΰT�����s�b�A�t�Ӳ��ʤ��s�b�T��


        }

        /// <summary>
        /// ����t�Ӯw�s���ʩM�t�ήw�s���ʸ�T
        /// </summary>
        /// <param name="dsFactory_Stock_Number">�t�Ӯw�s���ʸ�T</param>
        /// <param name="dsSystem_Stock_Number">�t�ήw�s���ʸ�T</param>
        /// <param name="dtStockDiff">���ʸ�T���ŰO��</param>
        /// <param name="dtStockSurplusDiff">���ʸ�T���`���ŰO��</param>
        //public void CompareFactorySys(DataSet dsFactory_Stock_Number,
        //                            DataSet dsSystem_Stock_Number,
        //                            ref DataTable dtStockDiff)
        //{

        //    DataRow[] drUsed = null;
        //    DataRow[] drUsedMoveIn = null;

        //    #region �����ʸ�T���ŰO��
        //    foreach (DataRow drFactoryStockNumber in dsFactory_Stock_Number.Tables[0].Rows)
        //    {
        //        string strStatus_Name = drFactoryStockNumber["Status_Name"].ToString();
        //        if (strStatus_Name == "�J�w")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[0].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
        //        }
        //        else if (strStatus_Name == "�h�f")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[1].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
        //        }
        //        else if (strStatus_Name == "�A�J�w")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[2].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
        //        }
        //        else if (strStatus_Name == "3D")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString() +
        //                                                    " AND Type_Name = '3D'");
        //        }
        //        else if (strStatus_Name == "DA")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString() +
        //                                                    " AND Type_Name = 'DA'");
        //        }
        //        else if (strStatus_Name == "PM")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString() +
        //                                                    " AND Type_Name = 'PM'");
        //        }
        //        else if (strStatus_Name == "RN")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[3].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString() +
        //                                                    " AND Type_Name = 'RN'");
        //        }
        //        else if (strStatus_Name == "����")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[4].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
        //            drUsedMoveIn = dsSystem_Stock_Number.Tables[5].Select("RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
        //        }
        //        else if (strStatus_Name == "�t�ӵ��l")
        //        {
        //            drUsed = dsSystem_Stock_Number.Tables[6].Select("Perso_Factory_RID = " + drFactoryStockNumber["Perso_Factory_RID"].ToString() +
        //                                                    " AND Type = " + drFactoryStockNumber["Type"].ToString() +
        //                                                    " AND Affinity = " + drFactoryStockNumber["Affinity"].ToString() +
        //                                                    " AND Photo = " + drFactoryStockNumber["Photo"].ToString());
        //        }
        //        // �ˬd�ƶq�O�_�۵�
        //        if (strStatus_Name == "����")
        //        {
        //            int intMoveNumber = 0;
        //            if (drUsed != null)
        //            {
        //                if (drUsed.Length > 0)
        //                {
        //                    intMoveNumber -= Convert.ToInt32(drUsed[0]["Number"].ToString());
        //                }
        //                if (drUsedMoveIn.Length > 0)
        //                {
        //                    intMoveNumber += Convert.ToInt32(drUsedMoveIn[0]["Number"].ToString());
        //                }

        //                // �ˬd����ƶq�O�_�۵�
        //                if (intMoveNumber != Convert.ToInt32(drFactoryStockNumber["Number"].ToString()))
        //                {

        //                }
        //            }
        //            // �ˬd�J�w�B�h�f�B�A�J�w�B3D�BDA�BPM�BRN�B����B���鵲�E
        //        }
        //        else if (strStatus_Name == "�J�w" ||
        //            strStatus_Name == "�h�f" ||
        //            strStatus_Name == "�A�J�w" ||
        //            strStatus_Name == "3D" ||
        //            strStatus_Name == "DA" ||
        //            strStatus_Name == "PM" ||
        //            strStatus_Name == "RN" ||
        //            strStatus_Name == "�t�ӵ��l")
        //        {
        //            int iSystem_Number = 0;
        //            if (drUsed != null)
        //            {
        //                if (drUsed.Length > 0)
        //                {
        //                    iSystem_Number = Convert.ToInt32(drUsed[0]["Number"].ToString());
        //                }
        //            }
        //            if (iSystem_Number !=
        //                Convert.ToInt32(drFactoryStockNumber["Number"].ToString()))
        //            {
        //                // �t�Ӯw�s���ʩM�t�ήw�s���ʤ��۲ŮɲK�[�O�s�ӰO��
        //                argFun[1] = drFactoryStockNumber["Factory_ShortName_CN"].ToString();
        //                argFun[2] = drFactoryStockNumber["NAME"].ToString();
        //                return false;
        //            }
        //            else
        //            {
        //                // �p�G���E�q���t�ơC�������d�~�]�n��ܡC
        //                if (iSystem_Number < 0 && strStatus_Name == "�t�ӵ��l")
        //                {
        //                    this.dirValues.Clear();
        //                    this.dirValues.Add("type", drFactoryStockNumber["Type"].ToString());
        //                    this.dirValues.Add("affinity", drFactoryStockNumber["Affinity"].ToString());
        //                    this.dirValues.Add("photo", drFactoryStockNumber["Photo"].ToString());
        //                    DataSet dstVirtualCardGroup = dao.GetList(CON_CARD_TYPE_GROUP, this.dirValues);
        //                    if (null != dstVirtualCardGroup && dstVirtualCardGroup.Tables.Count > 0 &&
        //                        dstVirtualCardGroup.Tables[0].Rows.Count > 0)
        //                    {
        //                        if (Convert.ToInt32(dstVirtualCardGroup.Tables[0].Rows[0]["Num"].ToString()) == 0)
        //                        {   // ���������d
        //                            // �t�Ӯw�s���ʩM�t�ήw�s���ʤ��۲ŮɲK�[�O�s�ӰO��
        //                            argFun[1] = drFactoryStockNumber["Factory_ShortName_CN"].ToString();
        //                            argFun[2] = drFactoryStockNumber["NAME"].ToString();

        //                            // XXXPerso�tXXX����²�ٮw�s�����ɡAĵ��
        //                            object[] arg = new object[2];
        //                            arg[0] = drFactoryStockNumber["Factory_ShortName_CN"].ToString();
        //                            arg[1] = drFactoryStockNumber["Name"].ToString();
        //                            Warning.SetWarning(GlobalString.WarningType.CardTypeNotEnough, arg);

        //                            return false;

        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    #endregion �����ʸ�T���ŰO��


        //    return true;
        //}

        /// <summary>
        /// ���t�μt�ӵ��lDataTable���c
        /// </summary>
        private DataTable getDTSys_Stock()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Perso_Factory_RID", Type.GetType("System.Int32")));
            dt.Columns.Add(new DataColumn("Factory_ShortName_CN", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("TYPE", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("AFFINITY", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("PHOTO", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Name", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Status_Name", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Number", Type.GetType("System.Int32")));
            return dt;
        }
        /// <summary>
        /// �i��鵲
        /// </summary>
        /// <returns></returns>
        public string DaySurplus(DateTime Date)
        {
            string strRet = "";
            try
            {
                dao.OpenConnection();
                // �d�ش����W���ܤƪ��B�z
                getWaferUsedLog(Date);
                // �Τp�p�ɥͦ��d�����������ƯӥΰO��
                List<string> lstMaterielUsed = (List<string>)getMaterialUsed(Date);
                // �p�⪫�ƳѧE�ƶq�}ĵ��
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/25 start
                InOut000BL BL000 = new InOut000BL();
                BL000.getDayMaterielStocks(Date, lstMaterielUsed);
                //getMaterielStocks(Date, lstMaterielUsed);
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/25 end
                //200911CR �鵲�ɭp��C�骫�Ʈw�s���E add by ���� 2009/11/26 start
                BL000.SaveSurplusSystemNum(Date);
                //200911CR �鵲�ɭp��C�骫�Ʈw�s���E add by ���� 2009/11/26 end
                // �p��N�s�O��
                getProjectCost(Date);
                // �N�����O���аO�]�m���鵲...����
                setDaySurplus(Date);
                // ����ư�
                dao.Commit();

                return strRet;
            }

            catch (Exception ex)
            {
                //�ưȦ^�u
                dao.Rollback();
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�i��鵲, DaySurplus����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                BatchBL Bbl = new BatchBL();
                Bbl.SendMail(ConfigurationManager.AppSettings["ManagerMail"], ConfigurationManager.AppSettings["MailSubject"], ConfigurationManager.AppSettings["MailBody"]);
                throw ex;
            }

            finally
            {
                //�����s��
                dao.CloseConnection();

            }
            return strRet;
        }

        ///// <summary>
        ///// �p�⪫�ƳѧE�ƶq�}ĵ��
        ///// </summary>
        ///// <param name="Date"></param>
        //public void getMaterielStocks(DateTime Surplus_Date,
        //                List<string> lstSerielNumber)
        //{
        //    try
        //    {
        //        Depository010BL bl010 = new Depository010BL();

        //        // �����~�s���ɡA�B�z
        //        if (lstSerielNumber.Count > 0)
        //        {
        //            string strSerielNumbers = "'";
        //            foreach (string strSerielNumberFor in lstSerielNumber)
        //            {
        //                strSerielNumbers += strSerielNumberFor + "','";
        //            }
        //            strSerielNumbers = strSerielNumbers.Substring(0, strSerielNumbers.Length - 2);

        //            // ����e�Ѫ��W�@�u�@��
        //            DateTime dtLastWorkDate = DateTime.Parse("1900-01-01");
        //            dirValues.Clear();
        //            dirValues.Add("date_time", Surplus_Date.ToString("yyyy/MM/dd"));
        //            DataSet dsWorkDate = dao.GetList(SEL_LAST_WORK_DATE, this.dirValues);
        //            if (null != dsWorkDate &&
        //                dsWorkDate.Tables.Count > 0 &&
        //                dsWorkDate.Tables[0].Rows.Count > 0)
        //            {
        //                dtLastWorkDate = Convert.ToDateTime(dsWorkDate.Tables[0].Rows[0]["Date_Time"].ToString());
        //            }

        //            // ����e�u�@�骺�W�@�u�@�骺�Ҧ��w�s
        //            dirValues.Clear();
        //            dirValues.Add("stock_date", dtLastWorkDate.ToString("yyyy/MM/dd"));
        //            //dirValues.Add("seriel_numbers", strSerielNumbers);
        //            DataSet dsMaterielStocksManager = dao.GetList(SEL_MATERIEL_STOCKS_MANAGER + strSerielNumbers+")", this.dirValues);
        //            if (null != dsMaterielStocksManager &&
        //                dsMaterielStocksManager.Tables.Count > 0 &&
        //                dsMaterielStocksManager.Tables[0].Rows.Count > 0)
        //            {
        //                foreach (DataRow drMSM in dsMaterielStocksManager.Tables[0].Rows)
        //                {
        //                    dirValues.Clear();
        //                    dirValues.Add("stock_date", Surplus_Date.ToString("yyyy/MM/dd"));
        //                    dirValues.Add("serial_number", drMSM["Serial_Number"].ToString());
        //                    dirValues.Add("perso_factory_rid", drMSM["Perso_Factory_RID"].ToString());
        //                    DataSet dsMaterielStocksUsed = dao.GetList(SEL_MATERIEL_USED, this.dirValues);
        //                    if (null != dsMaterielStocksUsed &&
        //                        dsMaterielStocksUsed.Tables.Count > 0 &&
        //                        dsMaterielStocksUsed.Tables[0].Rows.Count > 0)
        //                    {
        //                        // �e�@�Ѫ��w�s
        //                        int intLastStockNumber = 0;
        //                        int intTheDayUsedNumber = 0;

        //                        if (!StringUtil.IsEmpty(drMSM["Number"].ToString()))
        //                            intLastStockNumber = Convert.ToInt32(drMSM["Number"].ToString());
        //                        // ���Ѯ��Ӫ�
        //                        if (!StringUtil.IsEmpty(dsMaterielStocksUsed.Tables[0].Rows[0][0].ToString()))
        //                            intTheDayUsedNumber = Convert.ToInt32(dsMaterielStocksUsed.Tables[0].Rows[0][0].ToString());

        //                        // �w�s��0�ɡA��ܮw�s����
        //                        if (intLastStockNumber <= 0)
        //                        {
        //                            if (bl010.DmNotSafe_Type(drMSM["Serial_Number"].ToString()))
        //                            {
        //                                // �w�s����
        //                                string[] arg = new string[1];
        //                                arg[0] = drMSM["Name"].ToString();
        //                                Warning.SetWarning(GlobalString.WarningType.MaterialDataInMiss, arg);
        //                            }
        //                        }
        //                        // �p�G�e�@�Ѫ��w�s�p�E���Ѫ�����
        //                        else if (intLastStockNumber < intTheDayUsedNumber)
        //                        {
        //                            if (bl010.DmNotSafe_Type(drMSM["Serial_Number"].ToString()))
        //                            {
        //                                // �w�s����
        //                                string[] arg = new string[1];
        //                                arg[0] = drMSM["Name"].ToString();
        //                                Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInMiss, arg);
        //                                Warning.SetWarning(GlobalString.WarningType.SubtotalMaterialInMiss, arg);
        //                            }
        //                        }
        //                        else
        //                        {
        //                            // �����ƪ��w���w�s�T��
        //                            DataSet dtMateriel = GetMateriel(drMSM["Serial_Number"].ToString());
        //                            if (null != dtMateriel &&
        //                                dtMateriel.Tables.Count > 0 &&
        //                                dtMateriel.Tables[0].Rows.Count > 0)
        //                            {
        //                                // �̧C�w���w�s
        //                                if (GlobalString.SafeType.storage == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                                {
        //                                    // �t�ӵ��l�C��̧C�w���w�s�ƭȮ�
        //                                    if (intLastStockNumber - intTheDayUsedNumber <
        //                                        Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]))
        //                                    {
        //                                        string[] arg = new string[1];
        //                                        arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                                        Warning.SetWarning(GlobalString.WarningType.SubtoalMaterialInSafe, arg);
        //                                        Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInSafe, arg);
        //                                    }
        //                                    // �w���Ѽ�
        //                                }
        //                                else if (GlobalString.SafeType.days == Convert.ToString(dtMateriel.Tables[0].Rows[0]["Safe_Type"]))
        //                                {
        //                                    // �ˬd�w�s�O�_�R��
        //                                    if (!CheckMaterielSafeDays(drMSM["Serial_Number"].ToString(),
        //                                                            Convert.ToInt32(drMSM["Perso_Factory_RID"].ToString()),
        //                                                            Convert.ToInt32(dtMateriel.Tables[0].Rows[0]["Safe_Number"]),
        //                                                            intLastStockNumber - intTheDayUsedNumber)) 
        //                                    {
        //                                        string[] arg = new string[1];
        //                                        arg[0] = dtMateriel.Tables[0].Rows[0]["Name"].ToString();
        //                                        Warning.SetWarning(GlobalString.WarningType.SubtoalMaterialInSafe, arg);
        //                                        Warning.SetWarning(GlobalString.WarningType.PersoChangeMaterialInSafe, arg);
        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogFactory.Write(ex.ToString(), GlobalString.LogType.ErrorCategory);
        //        throw ex;
        //    }
        //}

        /// <summary>
        /// �����~���l�Ӳv
        /// </summary>    
        /// <param name="Serial_Number">���~�s�� 1�G�H�ʡF2�G�H�d��F3�GDM</param>
        /// <returns>Decimal<���~���ӥβv></returns>
        public Decimal GetWearRate(string Serial_Number)
        {
            Decimal dWearRate = 0;
            DataSet dstWearRate = null;

            try
            {
                dirValues.Clear();
                dirValues.Add("Serial_Number", Serial_Number);
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())// �H��
                {
                    dstWearRate = dao.GetList(SEL_ENVELOPE_INFO, dirValues);
                }
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())// �d��
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
                    // ���l�Ӳv
                    dWearRate = Convert.ToDecimal(dstWearRate.Tables[0].Rows[0]["Wear_Rate"]);
                }

                return dWearRate;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�����~���l�Ӳv, GetWearRate����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// �p�⪫�Ʈw�s������
        /// </summary>
        /// <param name="Factory_RID">Perso�t��RID</param>
        /// <param name="Serial_Number">���ƽs��</param>    
        /// <param name="lastSurplusDateTime">�̪�@�������l���</param>
        /// <param name="thisSurplusDateTime">�������l���</param>
        /// <returns>DataTable<���ƨϥΰO��></returns>
        public DataTable MaterielUsedCount(int Factory_RID,
                            string Serial_Number,
                            DateTime lastSurplusDateTime,
                            DateTime thisSurplusDateTime)
        {
            DataTable dtSubtotal_Import = null;
            try
            {
                dirValues.Clear();
                dirValues.Add("Perso_Factory_RID", Factory_RID);
                dirValues.Add("Serial_Number", Serial_Number);
                dirValues.Add("lastSurplusDateTime", lastSurplusDateTime);
                dirValues.Add("thisSurplusDateTime", thisSurplusDateTime);
                DataSet dstSTOCKS_USED = dao.GetList(SEL_MATERIEL_STOCKS_USED, dirValues);
                if (null != dstSTOCKS_USED && dstSTOCKS_USED.Tables.Count > 0 &&
                                dstSTOCKS_USED.Tables[0].Rows.Count > 0)
                {
                    dtSubtotal_Import = dstSTOCKS_USED.Tables[0];
                    dtSubtotal_Import.Columns.Add(new DataColumn("System_Num", Type.GetType("System.Int32")));
                    for (int intRow = 0; intRow < dtSubtotal_Import.Rows.Count; intRow++)
                    {
                        // �����~���l�Ӳv(���p�쪫�~��A�����~���l�Ӳv�^
                        //Decimal dWear_Rate = GetWearRate(Serial_Number);
                        // �t�ίӥζq                        
                        //�gUSED��ɤw�p����Ӳv�A���A�p��I
                        Decimal dWear_Rate = 0;
                        dtSubtotal_Import.Rows[intRow]["System_Num"] = Convert.ToInt32(dtSubtotal_Import.Rows[intRow]["Number"]) * (dWear_Rate / 100 + 1);
                    }
                }
                return dtSubtotal_Import;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�p�⪫�Ʈw�s������, MaterielUsedCount����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// �ˬd���ƪ��w�s�O�_�w���]�w���Ѽơ^
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
        //    Days = Days + 1;   // ���F�A���פJ�ɪ���ƻݭn�A�ݭn�h��@��
        //    DateTime dtStartTime = DateTime.Now.AddDays(-Days);
        //    DataTable dtblSubtotal_Import = MaterielUsedCount(Factory_RID,
        //                                        Serial_Number,
        //                                        dtStartTime,
        //                                        DateTime.Now);

        //    int intMaterielWear = 0;
        //    if (null != dtblSubtotal_Import &&
        //        dtblSubtotal_Import.Rows.Count > 0)
        //    {
        //        // �eN�Ѫ��ӥζq
        //        for (int intRow = 0; intRow < dtblSubtotal_Import.Rows.Count; intRow++)
        //        {
        //            intMaterielWear += Convert.ToInt32(dtblSubtotal_Import.Rows[intRow]["System_Num"]);
        //        }

        //        // �p�G�w�s�p�_�eN�Ѫ��ӥζq
        //        if (Stock_Number < intMaterielWear)
        //        {
        //            blCheckMaterielSafeDays = false;
        //        }
        //    }

        //    return blCheckMaterielSafeDays;
        //}

        /// <summary>
        /// �d�ش����W���ܤƪ��B�z
        /// </summary>
        /// <returns></returns>
        public bool getWaferUsedLog(DateTime Date)
        {
            WAFER_CARDTYPE_USELOG wcuMModel = new WAFER_CARDTYPE_USELOG();
            //�d�خ��Ӫ�
            DataTable dtUSE_CARDTYPE = new DataTable();
            dtUSE_CARDTYPE.Columns.Add("Perso_Factory_RID");
            dtUSE_CARDTYPE.Columns.Add("CardType_RID");
            dtUSE_CARDTYPE.Columns.Add("Number");

            try
            {
                dirValues.Clear();
                dirValues.Add("date_time_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("date_time_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                #region Perso�t���J�w�B�z
                //���Ҧ��鵲�����Ѫ��J�w�O��
                DataSet dsIN_STOCKS = dao.GetList(SEL_SYS_IN_STOCKS_SURPLUS, dirValues);

                //�N�J�w�O���K�[�ܴ����W���ܤƪ�
                for (int i = 0; i < dsIN_STOCKS.Tables[0].Rows.Count; i++)
                {
                    wcuMModel.Income_Date = Date;
                    wcuMModel.Usable_Number = Convert.ToInt32(dsIN_STOCKS.Tables[0].Rows[i]["Income_Number"]);
                    wcuMModel.Factory_RID = Convert.ToInt32(dsIN_STOCKS.Tables[0].Rows[i]["Perso_Factory_RID"]);
                    wcuMModel.CardType_RID = Convert.ToInt32(dsIN_STOCKS.Tables[0].Rows[i]["Space_Short_RID"]);

                    //�s�W�����W���ܤƪ�ɡA�g�J������̤p���1900/01/01
                    //wcuMModel.Begin_Date = Date;
                    wcuMModel.Begin_Date = DateTime.MinValue.AddYears(1899);
                    
                    wcuMModel.Wafer_RID = Convert.ToInt32(dsIN_STOCKS.Tables[0].Rows[i]["Wafer_RID"]);
                    wcuMModel.Operate_RID = Convert.ToInt32(dsIN_STOCKS.Tables[0].Rows[i]["RID"]);
                    wcuMModel.Operate_Type = "1";
                    wcuMModel.CardType_Move_RID = 0;
                    wcuMModel.Number = wcuMModel.Usable_Number;
                    dao.Add<WAFER_CARDTYPE_USELOG>(wcuMModel, "RID");
                }
                #endregion

                #region Perso�t���A�J�w�B�z
                //���Ҧ��鵲�����Ѫ��A�J�w�O��
                DataSet dsRESTOCK = dao.GetList(SEL_SYS_DEPOSITORY_RESTOCK_SURPLUS, dirValues);

                //�N�A�J�w�O���K�[�촹���W���ܤƪ�
                for (int i = 0; i < dsRESTOCK.Tables[0].Rows.Count; i++)
                {
                    wcuMModel.Income_Date = Date;
                    wcuMModel.Usable_Number = Convert.ToInt32(dsRESTOCK.Tables[0].Rows[i]["Reincome_Number"]);
                    wcuMModel.Factory_RID = Convert.ToInt32(dsRESTOCK.Tables[0].Rows[i]["Perso_Factory_RID"]);
                    wcuMModel.CardType_RID = Convert.ToInt32(dsRESTOCK.Tables[0].Rows[i]["Space_Short_RID"]);

                    //�s�W�����W���ܤƪ�ɡA�g�J������̤p���1900/01/01
                    //wcuMModel.Begin_Date = Date;
                    wcuMModel.Begin_Date = DateTime.MinValue.AddYears(1899);

                    wcuMModel.Wafer_RID = Convert.ToInt32(dsRESTOCK.Tables[0].Rows[i]["Wafer_RID"]);
                    wcuMModel.Operate_RID = Convert.ToInt32(dsRESTOCK.Tables[0].Rows[i]["RID"]);
                    wcuMModel.Operate_Type = "2";
                    wcuMModel.CardType_Move_RID = 0;
                    wcuMModel.Number = wcuMModel.Usable_Number;
                    dao.Add<WAFER_CARDTYPE_USELOG>(wcuMModel, "RID");
                }
                #endregion

                #region Perso�t���h�f�B�z
                //���Ҧ��鵲�����Ѫ��h�f�O��
                DataSet dsRETURN_STOCKS = dao.GetList(SEL_SYS_RETURN_STOCKS_SURPLUS, dirValues);
                DataSet dsWAFER_CARDTYPE_USELOG_RID = null;

                //�ΰh�f�O�������h�f�q�A���������W���ܤƪ����ѧE�ƶq
                for (int i = 0; i < dsRETURN_STOCKS.Tables[0].Rows.Count; i++)
                {
                    dirValues.Clear();
                    dirValues.Add("stock_rid", dsRETURN_STOCKS.Tables[0].Rows[i]["Stock_RID"].ToString());
                    dsWAFER_CARDTYPE_USELOG_RID = dao.GetList(SEL_WAFER_CARDTYPE_USELOG_RID, dirValues);

                    //�ˬd�ӰO���O�_�w�g�O�s
                    saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG_RID.Tables[0].Rows[0]["RID"]), Date);

                    dirValues.Clear();
                    dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG_RID.Tables[0].Rows[0]["RID"].ToString());
                    dirValues.Add("cancel_number", dsRETURN_STOCKS.Tables[0].Rows[i]["Cancel_Number"].ToString());
                    dirValues.Add("check_date", Date);

                    //���������W���ܤƪ����ѧE�ƶq
                    dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG, dirValues);
                }
                #endregion

                #region �d������B�z
                // �������W���ܤƪ��O��,DataTable<�����W���ܤƪ�>
                DataSet dsWAFER_CARDTYPE_USELOG = dao.GetList(SEL_WAFER_CARDTYPE_USELOG);
                // ���Ҧ��d���w�s����O���ADataTable<�d������>
                dirValues.Clear();
                dirValues.Add("check_date_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("check_date_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                DataSet dsCARD_TYPE_MOVE_SURPLUS = dao.GetList(SEL_CARD_TYPE_MOVE_SURPLUS, dirValues);
                foreach (DataRow dr in dsCARD_TYPE_MOVE_SURPLUS.Tables[0].Rows)
                {
                    int intWAFER_CARDTYPE_USELOGRows = 0;
                    intWAFER_CARDTYPE_USELOGRows = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows.Count;
                    for (int i = 0; i < intWAFER_CARDTYPE_USELOGRows; i++)
                    {
                        if (dr["From_Factory_RID"].ToString() == dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Factory_RID"].ToString() &&
                            dr["CardType_RID"].ToString() == dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["CardType_RID"].ToString())
                        {
                            if (Convert.ToInt32(dr["Move_Number"]) < Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]))
                            {
                                //�ˬd�ӰO���O�_�w�g�O�s
                                saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"]), Date);

                                //������W���ܤƪ�i�μƶq 
                                dirValues.Clear();
                                dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"].ToString());
                                dirValues.Add("number", Convert.ToInt32(dr["Move_Number"]));
                                dirValues.Add("check_date", Date);
                                dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG_1, dirValues);

                                //�����W���ܤƪ��K�[�A��JPerso�t�B�d�ءB���������B���J�ƶq�O��
                                //�s�W�����W���ܤƪ�ɡA�g�J������̤p���1900/01/01
                                //wcuMModel.Begin_Date = Convert.ToDateTime(dr["Move_Date"]);
                                wcuMModel.Begin_Date = DateTime.MinValue.AddYears(1899);

                                wcuMModel.Income_Date = Convert.ToDateTime(dr["Move_Date"]);
                                wcuMModel.CardType_RID = Convert.ToInt32(dr["CardType_RID"]);
                                wcuMModel.Factory_RID = Convert.ToInt32(dr["To_Factory_RID"]);
                                wcuMModel.Operate_RID = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_RID"]);
                                wcuMModel.Operate_Type = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_Type"].ToString();
                                wcuMModel.Usable_Number = Convert.ToInt32(dr["Move_Number"]);
                                wcuMModel.Wafer_RID = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Wafer_RID"]);
                                wcuMModel.CardType_Move_RID = Convert.ToInt32(dr["RID"]);
                                wcuMModel.Number = Convert.ToInt32(dr["Move_Number"]);
                                wcuMModel.Unit_Price = Convert.ToDecimal(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Unit_Price"]);
                                dao.Add<WAFER_CARDTYPE_USELOG>(wcuMModel, "RID");
                               

                                //DataTable<�����W���ܤƪ�>���K�[Perso�t�B�d�ءB���J�ƶq�]���Z��������@�ǳơ^
                                DataRow drow = dsWAFER_CARDTYPE_USELOG.Tables[0].NewRow();
                                drow["Begin_Date"] = dr["Move_Date"];
                                drow["Income_Date"] = dr["Move_Date"];
                                drow["CardType_RID"] = dr["CardType_RID"];
                                drow["Factory_RID"] = dr["To_Factory_RID"];
                                drow["Operate_RID"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_RID"];
                                drow["Operate_Type"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_Type"].ToString();
                                drow["Usable_Number"] = dr["Move_Number"];
                                drow["Wafer_RID"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Wafer_RID"];
                                dsWAFER_CARDTYPE_USELOG.Tables[0].Rows.Add(drow);
                                break;
                            }

                            if (Convert.ToInt32(dr["Move_Number"]) == Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]))
                            {
                                //�ˬd�ӰO���O�_�w�g�O�s
                                saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"]), Date);

                                //�N�i�μƶq�]�m��0�B�NEnd_Date�]�m���鵲���
                                dirValues.Clear();
                                dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"].ToString());
                                dirValues.Add("check_date", Date);
                                dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG_2, dirValues);

                                //�����W���ܤƪ��K�[�A��JPerso�t�B�d�ءB���������B���J�ƶq�O��
                                //�s�W�����W���ܤƪ�ɡA�g�J������̤p���1900/01/01
                                //wcuMModel.Begin_Date = Convert.ToDateTime(dr["Move_Date"]);
                                wcuMModel.Begin_Date = DateTime.MinValue.AddYears(1899);
                                wcuMModel.Income_Date = Convert.ToDateTime(dr["Move_Date"]);
                                wcuMModel.CardType_RID = Convert.ToInt32(dr["CardType_RID"]);
                                wcuMModel.Factory_RID = Convert.ToInt32(dr["To_Factory_RID"]);
                                wcuMModel.Operate_RID = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_RID"]);
                                wcuMModel.Operate_Type = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_Type"].ToString();
                                wcuMModel.Usable_Number = Convert.ToInt32(dr["Move_Number"]);
                                wcuMModel.Wafer_RID = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Wafer_RID"]);
                                wcuMModel.CardType_Move_RID = Convert.ToInt32(dr["RID"]);
                                wcuMModel.Number = Convert.ToInt32(dr["Move_Number"]);
                                wcuMModel.Unit_Price = Convert.ToDecimal(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Unit_Price"]);
                                dao.Add<WAFER_CARDTYPE_USELOG>(wcuMModel, "RID");
                                

                                //DataTable<�����W���ܤƪ�>���K�[Perso�t�B�d�ءB���J�ƶq�]���Z��������@�ǳơ^
                                DataRow drow = dsWAFER_CARDTYPE_USELOG.Tables[0].NewRow();
                                drow["Begin_Date"] = dr["Move_Date"];
                                drow["Income_Date"] = dr["Move_Date"];
                                drow["CardType_RID"] = dr["CardType_RID"];
                                drow["Factory_RID"] = dr["To_Factory_RID"];
                                drow["Operate_RID"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_RID"];
                                drow["Operate_Type"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_Type"];
                                drow["Usable_Number"] = dr["Move_Number"];
                                drow["Wafer_RID"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Wafer_RID"];
                                dsWAFER_CARDTYPE_USELOG.Tables[0].Rows.Add(drow);
                                break;
                            }

                            if (Convert.ToInt32(dr["Move_Number"]) > Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]))
                            {
                                //�ˬd�ӰO���O�_�w�g�O�s
                                saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"]), Date);

                                //�N�i�μƶq�]�m��0�B�NEnd_Date�]�m���鵲��� todo
                                dirValues.Clear();
                                dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"].ToString());
                                dirValues.Add("check_date", Date);
                                dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG_2, dirValues);
                                dr["Move_Number"] = Convert.ToInt32(dr["Move_Number"]) - Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]);

                                //�����W���ܤƪ��K�[�A��JPerso�t�B�d�ءB���������B���J�ƶq�O��
                                //�s�W�����W���ܤƪ�ɡA�g�J������̤p���1900/01/01
                                //wcuMModel.Begin_Date = Convert.ToDateTime(dr["Move_Date"]);
                                wcuMModel.Begin_Date = DateTime.MinValue.AddYears(1899);

                                wcuMModel.Income_Date = Convert.ToDateTime(dr["Move_Date"]);
                                wcuMModel.CardType_RID = Convert.ToInt32(dr["CardType_RID"]);
                                wcuMModel.Factory_RID = Convert.ToInt32(dr["To_Factory_RID"]);
                                wcuMModel.Operate_RID = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_RID"]);
                                wcuMModel.Operate_Type = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_Type"].ToString();
                                wcuMModel.Usable_Number = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]); //Convert.ToInt32(dr["Move_Number"]);
                                wcuMModel.Wafer_RID = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Wafer_RID"]);
                                wcuMModel.CardType_Move_RID = Convert.ToInt32(dr["RID"]);
                                wcuMModel.Number = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]);//Convert.ToInt32(dr["Move_Number"]);
                                wcuMModel.Unit_Price = Convert.ToDecimal(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Unit_Price"]);
                                dao.Add<WAFER_CARDTYPE_USELOG>(wcuMModel, "RID");
                                

                                //DataTable<�����W���ܤƪ�>���K�[Perso�t�B�d�ءB���J�ƶq�]���Z��������@�ǳơ^
                                DataRow drow = dsWAFER_CARDTYPE_USELOG.Tables[0].NewRow();
                                drow["Begin_Date"] = dr["Move_Date"];
                                drow["Income_Date"] = dr["Move_Date"];
                                drow["CardType_RID"] = dr["CardType_RID"];
                                drow["Factory_RID"] = dr["To_Factory_RID"];
                                drow["Operate_RID"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_RID"];
                                drow["Operate_Type"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Operate_Type"];
                                drow["Usable_Number"] = Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]);//dr["Move_Number"];
                                drow["Wafer_RID"] = dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Wafer_RID"];
                                dsWAFER_CARDTYPE_USELOG.Tables[0].Rows.Add(drow);
                            }
                        }
                    }
                }
                #endregion

                #region �t�Ӳ��ʸ�ƳB�z(�B�z�S��]�A�G3D�BDA�BPM�BRN�p�p�ɮ��ӥH�μ˥d +���s�d + �ɻs�d + �s�l�d + �ƥd +�P�����} + �ʥd + �P��+ �վ��L����)
                //���s�������W���ܤƪ��O��,DataTable<�����W���ܤƪ�>
                 dsWAFER_CARDTYPE_USELOG = dao.GetList(SEL_WAFER_CARDTYPE_USELOG);
                //���鵲�����Ѫ��Ҧ����ʰO��,DataSet(�w�s����)
                dirValues.Clear();
                dirValues.Add("date_time_start", Date.ToString("yyyy-MM-dd 00:00:00"));
                dirValues.Add("date_time_end", Date.ToString("yyyy-MM-dd 23:59:59"));
                DataSet dsUSE_CARDTYPE = dao.GetList(SEL_USE_CARDTYPE, dirValues);
                //���d�����Ӥ���
                DataSet dsEXPRESSIONS = dao.GetList(SEL_EXPRESSIONS_DEFINE);

                //��Perso�t�B�d�ت��p����Ӷq�]�`���[�`�U�ت��p�����Ӽƶq�^
                int Card_Type_Rid = 0;
                int Perso_Factory_RID = 0;
                int Number = 0;
                //todo ���`���i�H��i���s�x�L�{
                foreach (DataRow dr in dsUSE_CARDTYPE.Tables[0].Rows)
                {
                    if ((Convert.ToInt32(dr["RID"]) != Card_Type_Rid) ||
                        (Convert.ToInt32(dr["Perso_Factory_RID"]) != Perso_Factory_RID))
                    {
                        if (Card_Type_Rid != 0 && Perso_Factory_RID != 0 && Number != 0)
                        {
                            DataRow drow = dtUSE_CARDTYPE.NewRow();
                            drow["Number"] = Number.ToString();
                            drow["Perso_Factory_RID"] = Perso_Factory_RID.ToString();
                            drow["CardType_RID"] = Card_Type_Rid.ToString();
                            dtUSE_CARDTYPE.Rows.Add(drow);
                        }

                        #region �����ӥd����,�p����ӥd��
                        Number = 0;
                        DataRow[] drEXPRESSIONS = dsEXPRESSIONS.Tables[0].Select("Type_RID = " + dr["Status_RID"].ToString());
                        if (drEXPRESSIONS.Length > 0)
                        {
                            if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Add_RID)
                            {
                                Number += Convert.ToInt32(dr["Number"]);
                                Card_Type_Rid = Convert.ToInt32(dr["RID"]);
                                Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
                            }
                            else if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Del_RID)
                            {
                                Number -= Convert.ToInt32(dr["Number"]);
                                Card_Type_Rid = Convert.ToInt32(dr["RID"]);
                                Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region �����ӥd����,�p����ӥd��
                        DataRow[] drEXPRESSIONS = dsEXPRESSIONS.Tables[0].Select("Type_RID = " + dr["Status_RID"].ToString());
                        if (drEXPRESSIONS.Length > 0)
                        {
                            if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Add_RID)
                            {
                                Number += Convert.ToInt32(dr["Number"]);
                            }
                            else if (drEXPRESSIONS[0]["Operate"].ToString() == GlobalString.Operation.Del_RID)
                            {
                                Number -= Convert.ToInt32(dr["Number"]);
                            }
                        }
                        #endregion
                    }
                }
                if (Card_Type_Rid != 0 && Perso_Factory_RID != 0 && Number != 0)
                {
                    DataRow drow = dtUSE_CARDTYPE.NewRow();
                    drow["Number"] = Number.ToString();
                    drow["Perso_Factory_RID"] = Perso_Factory_RID.ToString();
                    drow["CardType_RID"] = Card_Type_Rid.ToString();
                    dtUSE_CARDTYPE.Rows.Add(drow);
                }

                //�ھ�Perso�t�B�d�ءB���Ӷq�A�������W���ܤƪ��d�سѧE�ƶq�CDataTable<�d�خ��Ӫ�>
                foreach (DataRow dr in dtUSE_CARDTYPE.Rows)
                {
                    for (int i = 0; i < dsWAFER_CARDTYPE_USELOG.Tables[0].Rows.Count; i++)
                    {
                        if (dr["Perso_Factory_RID"].ToString() == dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Factory_RID"].ToString() &&
                            dr["CardType_RID"].ToString() == dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["CardType_RID"].ToString())
                        {
                            if (Convert.ToInt32(dr["Number"]) < Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]))
                            {
                                //�O�s�������e���O���A�H�ƨ����鵲�ɫ�_�w�s�q�C
                                saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"]), Date);

                                //������W���ܤƪ�i�μƶq 
                                dirValues.Clear();
                                dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"].ToString());
                                dirValues.Add("number", Convert.ToInt32(dr["Number"]));
                                dirValues.Add("check_date", Date);
                                dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG_1, dirValues);
                                break;
                            }

                            if (Convert.ToInt32(dr["Number"]) == Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]))
                            {
                                //�O�s�������e���O���A�H�ƨ����鵲�ɫ�_�w�s�q�C
                                saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"]), Date);

                                //�N�i�μƶq�]�m��0�B�NEnd_Date�]�m���鵲���
                                dirValues.Clear();
                                dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"].ToString());
                                dirValues.Add("check_date", Date);
                                dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG_2, dirValues);
                                break;
                            }

                            if (Convert.ToInt32(dr["Number"]) > Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]))
                            {
                                //�O�s�������e���O���A�H�ƨ����鵲�ɫ�_�w�s�q�C
                                saveWafer_Uselog_Rollback(Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"]), Date);

                                //�N�i�μƶq�]�m��0�B�NEnd_Date�]�m���鵲��� todo
                                dirValues.Clear();
                                dirValues.Add("rid", dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["RID"].ToString());
                                dirValues.Add("check_date", Date);
                                dao.ExecuteNonQuery(UPDATE_WAFER_CARDTYPE_USELOG_2, dirValues);
                                dr["Number"] = Convert.ToInt32(dr["Number"]) - Convert.ToInt32(dsWAFER_CARDTYPE_USELOG.Tables[0].Rows[i]["Usable_Number"]);
                            }
                        }
                    }
                }

                #endregion

              

                #region 2009/01/11 �����������R,�N�鵲��ѨS���ܤƪ������W���ɫO�s��^�u�ɤ�
                dirValues.Clear();
                dirValues.Add("date_time_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("date_time_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                // ���s�������W���ܤƪ��O��,DataTable<�����W���ܤƪ�>
                dsWAFER_CARDTYPE_USELOG = dao.GetList(SEL_WAFER_CARDTYPE_USELOG_FIRST_ZERO);
                foreach (DataRow drChengben in dsWAFER_CARDTYPE_USELOG.Tables[0].Rows)
                {
                    //�ˬd�ӰO���O�_�w�g�O�s
                    saveWafer_Uselog_Rollback(Convert.ToInt32(drChengben["RID"]), Date);
                }
                #endregion 2009/01/11 �����������R,�N�鵲��ѨS���ܤƪ������W���ɫO�s��^�u�ɤ�

                //��ѤJ�w���O���A������COPY��ROLLBACK�ɡA�ҥH�b�Ҧ��s�W������A�R���@�B�R���ާ@�I
                dirValues.Clear();
                dirValues.Add("check_date_start", Date.ToString("yyyy-MM-dd 00:00:00"));
                dao.ExecuteNonQuery(DEL_WAFER_USELOG_ROLLBACK_TODAY, dirValues);

                return true;
            }

            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�d�ش����W���ܤƪ��B�z, getWaferUsedLog����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
            return true;
        }

        /// <summary>
        /// �ھڪ��ƩM�ƶq�p���ڪ��ƶq�I
        /// </summary>
        /// <param name="MNumber"></param>
        /// <param name="MCount"></param>
        /// <returns></returns>
        private int ComputeMaterialNumber(string MNumber, long MCount)
        {
            int iReturn = 0;
            decimal dWear_Rate = this.GetWearRate(MNumber);
            iReturn = Convert.ToInt32(MCount * (dWear_Rate / 100 + 1));
            return iReturn;
        }


        /// <summary>
        /// �Τp�p�ɥͦ��d�����������ƯӥΰO��
        /// </summary>
        /// <returns></returns>
        //public List<string> getMaterialUsed(DateTime Date)
        //{
        //    List<string> lstSerielNumber = new List<string>();
        //    MATERIEL_STOCKS_USED msuModel = new MATERIEL_STOCKS_USED();
        //    try
        //    {
        //        dirValues.Clear();
        //        dirValues.Add("check_date_start", Date.ToString("yyyy/MM/dd 00:00:00"));
        //        dirValues.Add("check_date_end", Date.ToString("yyyy/MM/dd 23:59:59"));
        //        //���H�ʩM�H�d��ӥΰO���ADataSet<���ƯӥΰO��>
        //        DataSet dsMATERIAL_BY_SUBTOTAL = dao.GetList(SEL_MATERIAL_BY_SUBTOTAL, dirValues);
        //        foreach (DataRow dr in dsMATERIAL_BY_SUBTOTAL.Tables[0].Rows)
        //        {
        //            if (dr["CE_Number"].ToString() != null && dr["CE_Number"].ToString() != "")
        //            {
        //                // �O�s���ƫ~�W�s���A���P�_���ƪ��w�s�M�w������@�ǳ�
        //                if (-1 == lstSerielNumber.IndexOf(dr["CE_Number"].ToString()))
        //                {
        //                    lstSerielNumber.Add(dr["CE_Number"].ToString());
        //                }
        //                msuModel.Stock_Date = Date;
        //                msuModel.Number = this.ComputeMaterialNumber(dr["CE_Number"].ToString(), Convert.ToInt64(dr["Number1"]));
        //                msuModel.Serial_Number = dr["CE_Number"].ToString();
        //                msuModel.Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                dao.Add<MATERIEL_STOCKS_USED>(msuModel, "RID");
        //            }

        //            if (dr["EI_Number"].ToString() != null && dr["EI_Number"].ToString() != "")
        //            {
        //                // �O�s���ƫ~�W�s���A���P�_���ƪ��w�s�M�w������@�ǳ�
        //                if (-1 == lstSerielNumber.IndexOf(dr["EI_Number"].ToString()))
        //                {
        //                    lstSerielNumber.Add(dr["EI_Number"].ToString());
        //                }

        //                msuModel.Stock_Date = Date;
        //                msuModel.Number = this.ComputeMaterialNumber(dr["EI_Number"].ToString(), Convert.ToInt64(dr["Number1"]));
        //                msuModel.Serial_Number = dr["EI_Number"].ToString();
        //                msuModel.Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                dao.Add<MATERIEL_STOCKS_USED>(msuModel, "RID");
        //            }
        //        }

        //        //��DM�ӥΰO���ADataSet<DM���ƯӥΰO��>
        //        dirValues.Clear();
        //        dirValues.Add("check_date_start", Date.ToString("yyyy/MM/dd 00:00:00"));
        //        dirValues.Add("check_date_end", Date.ToString("yyyy/MM/dd 23:59:59"));

        //        DataSet MATERIAL_BY_SUBTOTAL_DM = dao.GetList(SEL_MATERIAL_BY_SUBTOTAL_DM, dirValues);
        //        foreach (DataRow dr in MATERIAL_BY_SUBTOTAL_DM.Tables[0].Rows)
        //        {
        //            if (dr["DI_Number"].ToString() != "")
        //            {
        //                // �O�s���ƫ~�W�s���A���P�_���ƪ��w�s�M�w������@�ǳ�
        //                if (-1 == lstSerielNumber.IndexOf(dr["DI_Number"].ToString()))
        //                {
        //                    lstSerielNumber.Add(dr["DI_Number"].ToString());
        //                }
        //                if (dr["Card_Type_Link_Type"].ToString() == "1" ||
        //               (dr["Card_Type_Link_Type"].ToString() == "2" && dr["CardType_RID"].ToString() != ""))
        //                {
        //                    msuModel.Stock_Date = Date;
        //                    msuModel.Number = this.ComputeMaterialNumber(dr["DI_Number"].ToString(), Convert.ToInt64(dr["Number1"]));
        //                    msuModel.Serial_Number = dr["DI_Number"].ToString();
        //                    msuModel.Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
        //                    dao.Add<MATERIEL_STOCKS_USED>(msuModel, "RID");
        //                }
        //            }
        //        }

        //        return lstSerielNumber;
        //    }
        //    catch (Exception ex)
        //    {
        //        LogFactory.Write(ex.ToString(), GlobalString.LogType.ErrorCategory);
        //        return lstSerielNumber;
        //        throw ex;
        //    }
        //}
        /// <summary>
        /// �Τp�p�ɥͦ��d�����������ƯӥΰO��
        /// </summary>
        /// <returns></returns>
        public List<string> getMaterialUsed(DateTime Date)
        {
            //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/25 start
            // Depository010BL bl = new Depository010BL();
            List<string> lstSerielNumber = new List<string>();
            try
            {
                //lstSerielNumber = bl.SaveMaterielUsedCount(Date);
                InOut000BL BL000 = new InOut000BL();
                lstSerielNumber=BL000.SaveMaterielUsedCount(Date);
                //200908CR���ƪ����ӭp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/25 end
                return lstSerielNumber;
            }            
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�Τp�p�ɥͦ��d�����������ƯӥΰO��, getMaterialUsed����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return lstSerielNumber;
                throw ex;
            }
        }
        /// <summary>
        /// �����ƪ����~�BRID���T��
        /// </summary>
        /// <param name="Serial_Number">�~�W�s��</param>
        /// <returns><DataTable>����DataTable</returns>
        public DataSet GetMateriel(string Serial_Number)
        {
            DataSet dtsMateriel = null;
            try
            {
                // �����ƪ��~�W
                this.dirValues.Clear();
                this.dirValues.Add("serial_number", Serial_Number);

                // �H��
                if ("A" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_ENVELOPE_INFO, this.dirValues);
                // �H�d��
                else if ("B" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_CARD_EXPONENT, this.dirValues);
                // DM
                else if ("C" == Serial_Number.Substring(0, 1).ToUpper())
                    dtsMateriel = (DataSet)dao.GetList(SEL_DMTYPE_INFO, this.dirValues);
                return dtsMateriel;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�����ƪ����~�BRID���T��, GetMateriel����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
        }

        /// <summary>
        /// �p��N�s�O��
        /// </summary>
        /// <returns></returns>
        public void getProjectCost(DateTime Date)
        {
            MATERIEL_STOCKS_USED msuModel = new MATERIEL_STOCKS_USED();
            PERSO_PROJECT_DETAIL prdModel = new PERSO_PROJECT_DETAIL();

            try
            {
                // ���R���鵲��Ѫ��@��N�s����
                dirValues.Clear();
                dirValues.Add("Begin_Date", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("Finish_Date", Date.ToString("yyyy/MM/dd 23:59:59"));
                dao.ExecuteNonQuery(DEL_MAKE_COST_FROM_SUBTOTAL_IMPORT, this.dirValues);

                // �ͦ��@��N�s���ضO�ΰT��
                // 1�B���p�p�ɸ�T
                dirValues.Clear();
                dirValues.Add("check_date_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("check_date_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                //���鵲�����Ѫ��p�p�ɡADataSet<�p�p��>
                //DataSet dsSUBTOTAL_PROJECT_COST = dao.GetList(SEL_SUBTOTAL_PROJECT_COST, dirValues);
              
                //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 start
                DataSet dsSUBTOTAL_PROJECT_COST = dao.GetList(SEL_SUBTOTAL_REPLACE_PROJECT_COST, dirValues);
                //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 end
                if (null != dsSUBTOTAL_PROJECT_COST &&
                    dsSUBTOTAL_PROJECT_COST.Tables.Count > 0 &&
                    dsSUBTOTAL_PROJECT_COST.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsSUBTOTAL_PROJECT_COST.Tables[0].Rows)
                    {
                        // 2�B���d�ع������@��N�s���س��
                        this.dirValues.Clear();
                        dirValues.Add("Date_Time", Date.ToString("yyyy/MM/dd 12:00:00"));
                        dirValues.Add("CTRID", dr["RID"].ToString());
                        dirValues.Add("perso_factory_rid", dr["Perso_Factory_RID"].ToString());
                        // ���d�ت��@��N�s����
                        DataSet dsPROJECT_STEP_SURPLUS = dao.GetList(SEL_PROJECT_STEP_SURPLUS, dirValues);

                        // �K�[�@��N�s���ذT��
                        if (null != dsPROJECT_STEP_SURPLUS &&
                            dsPROJECT_STEP_SURPLUS.Tables.Count > 0 &&
                            dsPROJECT_STEP_SURPLUS.Tables[0].Rows.Count > 0)
                        {
                            // 3�B�K�[�@��N�s���ضO��
                            prdModel.Unit_Price = Convert.ToDecimal(dsPROJECT_STEP_SURPLUS.Tables[0].Rows[0]["Price"]);
                            prdModel.Number = Convert.ToInt32(dr["Number"]);
                            prdModel.Sum = prdModel.Unit_Price * prdModel.Number;
                            prdModel.Use_Date = Date;
                            prdModel.Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
                            prdModel.Card_Group_RID = Convert.ToInt32(dr["CARDGROUPRID"]);
                            prdModel.CardType_RID = Convert.ToInt32(dr["RID"]);
                            prdModel.Project_RID = Convert.ToInt32(dsPROJECT_STEP_SURPLUS.Tables[0].Rows[0]["RID"]);
                            dao.Add<PERSO_PROJECT_DETAIL>(prdModel, "RID");
                        }
                    }
                }

                #region �p��N�s�O�άO�_�W�X�w���ȡA�p�G�d�L�Aĵ��

                string Year = Date.Year.ToString();
                //�p��N�s�O���`�p�]�d�^
                int CostSumCard = 0;
                //�p��N�s�O���`�p�]�ȡ^
                int CostSumBank = 0;

                dirValues.Clear();
                dirValues.Add("year", Year);
                // �p�⥻�~�ׯS��N�s�O���`�p
                DataSet dsSPECIAL_PROJECT_COST = dao.GetList(SEL_SPECIAL_PROJECT_COST, dirValues);
                // �p�⥻�~�רҥ~�N�s�O���`�p
                DataSet dsEXCEPTION_PROJECT_COST = dao.GetList(SEL_EXCEPTION_PROJECT_COST, dirValues);
                // �p�⥻�~�ץN�s���ʶO���`�p��
                DataSet dsPERSO_PROJECT_CHANGE_DETAIL = dao.GetList(SEL_PERSO_PROJECT_CHANGE_DETAIL, dirValues);
                // ���@��N�s���ضO��
                DataSet dsPERSO_PROJECT_NORMAL = dao.GetList(SEL_PERSO_PROJECT_NORMAL, dirValues);
                // ���ƥN�s�O�Φ~�׹w��A9�G�N�s�O�� �]�d�^
                DataSet dsMATERIEL_BUDGET_SUM_CARD = dao.GetList(SEL_MATERIEL_BUDGET_SUM_CARD, dirValues);
                // ���ƥN�s�O�Φ~�׹w��A10�G�N�s�O�Ρ]�ȡ^ 
                DataSet dsMATERIEL_BUDGET_SUM_BANK = dao.GetList(SEL_MATERIEL_BUDGET_SUM_BANK, dirValues);

                // �p��N�s�O���`�p�]�d�^ = �S���إN�s�O�� + 
                //                        (�ϱ��H�Υd�B�����H�Υd�BVISA DEBIT�d�s��)���@�붵�إN�s�O�� + 
                //                        (�ϱ��H�Υd�B�����H�Υd�BVISA DEBIT�d�s��)���ҥ~���إN�s�O�� - 
                //                        (�ϱ��H�Υd�B�����H�Υd�BVISA DEBIT�d�s��)����Ȳ��ʶO��  
                // �p��N�s�O���`�p�]�ȡ^ = (�������ĥd�M�{���d���s��)���@�붵�إN�s�O�� + 
                //                        (�������ĥd�M�{���d���s��)���ҥ~���إN�s�O�� - 
                //                        (�������ĥd�M�{���d���s��)����Ȳ��ʶO��  

                // �S���إN�s�O��
                if (dsSPECIAL_PROJECT_COST.Tables[0].Rows[0][0].ToString() == "")
                    CostSumCard = 0;
                else
                    CostSumCard = Convert.ToInt32(dsSPECIAL_PROJECT_COST.Tables[0].Rows[0][0]);
                // �@�붵�إN�s�O��
                foreach (DataRow dr in dsPERSO_PROJECT_NORMAL.Tables[0].Rows)
                {
                    if (dr["Group_Name"].ToString() == "�ϱ��H�Υd" ||
                        dr["Group_Name"].ToString() == "�����H�Υd" ||
                        dr["Group_Name"].ToString() == "VISA DEBIT�d")
                    {
                        CostSumCard += Convert.ToInt32(dr[1]);
                    }
                    else if (dr["Group_Name"].ToString() == "�������ĥd" ||
                            dr["Group_Name"].ToString() == "�{���d")
                    {
                        CostSumBank += Convert.ToInt32(dr[1]);
                    }
                }
                // �ҥ~���ت��N�s�O��
                foreach (DataRow dr in dsEXCEPTION_PROJECT_COST.Tables[0].Rows)
                {
                    if (dr["Group_Name"].ToString() == "�ϱ��H�Υd" ||
                        dr["Group_Name"].ToString() == "�����H�Υd" ||
                        dr["Group_Name"].ToString() == "VISA DEBIT�d")
                    {
                        CostSumCard += Convert.ToInt32(dr[1]);
                    }
                    else if (dr["Group_Name"].ToString() == "�������ĥd" ||
                            dr["Group_Name"].ToString() == "�{���d")
                    {
                        CostSumBank += Convert.ToInt32(dr[1]);
                    }
                }

                // ��Ȳ���
                foreach (DataRow dr in dsPERSO_PROJECT_CHANGE_DETAIL.Tables[0].Rows)
                {
                    if (dr["Group_Name"].ToString() == "�ϱ��H�Υd" ||
                        dr["Group_Name"].ToString() == "�����H�Υd" ||
                        dr["Group_Name"].ToString() == "VISA DEBIT�d")
                    {
                        CostSumCard += Convert.ToInt32(dr[1]);
                    }
                    else if (dr["Group_Name"].ToString() == "�������ĥd" ||
                            dr["Group_Name"].ToString() == "�{���d")
                    {
                        CostSumBank += Convert.ToInt32(dr[1]);
                    }
                }

                // �ˬd�N�s�O�Υd�O�_�ݭn��ĵ
                if (null != dsMATERIEL_BUDGET_SUM_CARD &&
                    dsMATERIEL_BUDGET_SUM_CARD.Tables.Count > 0 &&
                    dsMATERIEL_BUDGET_SUM_CARD.Tables[0].Rows.Count > 0)
                {
                    Decimal intMATERIEL_BUDGET_SUM_CARD = Convert.ToDecimal(dsMATERIEL_BUDGET_SUM_CARD.Tables[0].Rows[0]["Budget"]);
                    if ((intMATERIEL_BUDGET_SUM_CARD - CostSumCard) < intMATERIEL_BUDGET_SUM_CARD * System.Decimal.Parse("0.1"))
                    {
                        object[] arg = new object[1];
                        arg[0] = "�N�s�O�Ρ]�d�^";
                        Warning.SetWarning(GlobalString.WarningType.SurplusMaterialBuget, arg);
                    }
                }

                if (null != dsMATERIEL_BUDGET_SUM_BANK &&
                        dsMATERIEL_BUDGET_SUM_BANK.Tables.Count > 0 &&
                        dsMATERIEL_BUDGET_SUM_BANK.Tables[0].Rows.Count > 0)
                {
                    Decimal intMATERIEL_BUDGET_SUM_BANK = Convert.ToDecimal(dsMATERIEL_BUDGET_SUM_BANK.Tables[0].Rows[0]["Budget"]);
                    if ((intMATERIEL_BUDGET_SUM_BANK - CostSumBank) < intMATERIEL_BUDGET_SUM_BANK * System.Decimal.Parse("0.1"))
                    {
                        object[] arg = new object[1];
                        arg[0] = "�N�s�O�Ρ]�ȡ^";
                        Warning.SetWarning(GlobalString.WarningType.SurplusMaterialBuget, arg);
                    }
                }

                #endregion

            }

            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�p��N�s�O��, getProjectCost����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }

        /// <summary>
        /// �N�����O���аO�]�m���鵲
        /// </summary>
        /// <returns></returns>
        public void setDaySurplus(DateTime Date)
        {
            try
            {
                CARDTYPE_STOCKS csModel = new CARDTYPE_STOCKS();

                dirValues.Clear();
                dirValues.Add("check_date", Date.ToString("yyyy/MM/dd"));
                dirValues.Add("check_date_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("check_date_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                //�N�鵲�����Ѫ��J�w�����Ѭ��鵲
                //for Jacky test   dao.ExecuteNonQuery("update batch_manage set status ='N'");
                dao.ExecuteNonQuery(UPDATE_DEPOSITORY_STOCK, dirValues);
                //�N�鵲�����Ѫ��A�J�w�����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_DEPOSITORY_RESTOCK, dirValues);
                ////�R����Ѫ����Ưӥξ�
                //dao.ExecuteNonQuery(DEL_MATERIEL_STOCKS_USED, dirValues);
                //�N�鵲�����Ѫ��h�f�����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_DEPOSITORY_CANCEL, dirValues);
                //�N�鵲�����Ѫ��p�p�ɪ����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_SUBTOTAL_IMPORT, dirValues);
                //�N�鵲�����Ѫ��t�Ӳ��ʪ����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_FACTORY_CHANGE_IMPORT, dirValues);

                //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 start
                //�N�鵲�����Ѫ������e�����p�p�ɪ����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_SUBTOTAL_REPLACE_IMPORT, dirValues);
                //�N�鵲�����Ѫ������e�����t�Ӳ��ʪ����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_FACTORY_CHANGE_REPLACE_IMPORT, dirValues);
                //200908CR�N��O�έp��אּ�Τp�p�ɪ��u�����e�v�����p�� add by ���� 2009/09/03 end

                //�N�鵲�����Ѫ��d�����ʪ����Ѭ��鵲
                dao.ExecuteNonQuery(UPDATE_CARDTYPE_STOCKS_MOVE, dirValues);
                //�O�s�d�خw�s��
                DataSet dsCARDTYPE_STOCKS = dao.GetList(SEL_CARDTYPE_STOCKS, dirValues);
                foreach (DataRow dr in dsCARDTYPE_STOCKS.Tables[0].Rows)
                {
                    csModel.Stock_Date = Convert.ToDateTime(dr["Date_Time"]);
                    csModel.Stocks_Number = Convert.ToInt32(dr["Number"]);
                    csModel.Perso_Factory_RID = Convert.ToInt32(dr["Perso_Factory_RID"]);
                    csModel.CardType_RID = Convert.ToInt32(dr["RID"]);
                    dao.Add<CARDTYPE_STOCKS>(csModel, "RID");
                }
            }

            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�N�����O���аO�]�m���鵲, setDaySurplus����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                throw ex;
            }
        }
        /// <summary>
        /// �ˬd�ӰO���O�_�w�g�O�s,�p�GCount == 0���s�b,
        /// �b�����W���ܤƦ^�u�ɡ]WAFER_USELOG_ROLLBACK�^���K�[�A�����W���ܤƪ����T��
        /// </summary>
        /// <returns></returns>
        public void saveWafer_Uselog_Rollback(int WAFER_RID, DateTime Date)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("uselog_rid", WAFER_RID);
                dirValues.Add("check_date_start", Date.ToString("yyyy/MM/dd 00:00:00"));
                dirValues.Add("check_date_end", Date.ToString("yyyy/MM/dd 23:59:59"));
                dirValues.Add("check_date", Date);
                DataSet dsWAFER_USELOG_ROLLBACK = dao.GetList(CON_WAFER_USELOG_ROLLBACK, dirValues);
                // �ˬd�ӰO���O�_�w�g�O�s,�p�GCount == 0���s�b,�b�����W���ܤƦ^�u�ɡ]WAFER_USELOG_ROLLBACK�^���K�[�A�����W���ܤƪ����T��
                // �C�ѥu�O�s�@���C
                if (null != dsWAFER_USELOG_ROLLBACK &&
                    dsWAFER_USELOG_ROLLBACK.Tables.Count > 0 &&
                    dsWAFER_USELOG_ROLLBACK.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToInt32(dsWAFER_USELOG_ROLLBACK.Tables[0].Rows[0][0].ToString()) == 0)
                    {
                        // �b�����W���ܤƦ^�u�ɤ��O�s�����W���ܤƪ��T��
                        dao.ExecuteNonQuery(INSERT_WAFER_USELOG_ROLLBACK, dirValues);
                    }
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("�ˬd�ӰO���O�_�w�g�O�s, saveWafer_Uselog_Rollback����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
            }
        }
        /// <summary>
        /// ���o�̦Z�@���鵲�����
        /// </summary>

        public string GetLastStock_Date()
        {
            try
            {
                DataSet dst = dao.GetList(SEL_SURPLUS_CHECK);
                if (dst.Tables[0].Rows.Count != 0)
                {
                    return Convert.ToDateTime(dst.Tables[0].Rows[0][0].ToString()).ToString("yyyy/MM/dd");
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���o�̦Z�@���鵲�����, GetLastStock_Date����:" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return "";
            }
        }
        /// <summary>
        /// ���ݭn�@�鵲���u�@��
        /// </summary>
        public DataTable GetNeedStock_Date(string time)
        {
            try
            {
                dirValues.Clear();
                dirValues.Add("lasttime", time);
                dirValues.Add("now", DateTime.Now.ToString("yyyy/MM/dd"));
                DataTable dt = dao.GetList(SEL_WORKDATE_NOT_SURPLUS, dirValues).Tables[0];
                return dt;
            }
            catch (Exception ex)
            {
                // Legend 2018/4/13 �ɥRLog���e
                LogFactory.Write("���ݭn�@�鵲���u�@��, GetNeedStock_Date����: " + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }

        }
       


    }
}
