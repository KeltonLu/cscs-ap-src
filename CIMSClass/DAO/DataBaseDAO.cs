//******************************************************************
//*  作    者：FangBao
//*  功能說明：資料庫操作物件實現各類型的DAO功能
//*  創建日期：2008-08-11
//*  修改日期：2008-08-11 12:00
//*  修改記錄：



//*            □2008-08-11
//*              1.創建 鮑方
//*******************************************************************

using System;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Practices.EnterpriseLibrary.Data;
//using System.Web;
using System.Text;
using System.ComponentModel;
using CIMSClass;

/// <summary>
/// DataBaseDAO 的摘要描述



/// </summary>
namespace CIMSClass
{
    public class DataBaseDAO
    {
        private Database dbWorker = null;//資料庫操作基本物件



        private DbConnection conWorker = null;//管理器資料庫鏈接
        private DbTransaction trsWork = null;//事務物件
        private List<DbCommand> listLastCommands = new List<DbCommand>();//最後執行的命令集合

        private DatabaseWork dwWork = default(DatabaseWork);//支持事務的資料庫操作DAO的委託



        /// <summary>
        /// 支持事務的資料庫操作DAO的委託



        /// </summary>
        public DatabaseWork Work { get { return dwWork; } }

        //是否事務處理
        private bool blTransactionState = false;


        /// <summary>
        /// 事務處理狀態



        /// </summary>
        public bool TransactionState
        {
            get
            {
                return blTransactionState;
            }
            set
            {
                blTransactionState = value;
            }
        }

        /// <summary>
        /// 最後執行的命令集合
        /// </summary>
        public DbCommand[] LastCommands
        {
            get
            {
                return listLastCommands.ToArray();
            }
        }
        /// <summary>
        /// 事務是否開始
        /// </summary>
        private bool TransactionBegined
        {
            get
            {
                if (trsWork == null)
                {
                    conWorker.Open();
                    trsWork = conWorker.BeginTransaction();
                }
                return true;
            }
        }

        /// <summary>
        /// 當前管理器的連接狀態



        /// </summary>
        public ConnectionState NowSatae
        {
            get
            {
                return conWorker.State;
            }
        }

        /// <summary>
        /// 建構
        /// </summary>
        public DataBaseDAO()
        {
            #region Legend 2016/10/26 因企業庫升級調整寫法

            //dbWorker = DatabaseFactory.CreateDatabase();

            DatabaseProviderFactory factory = new DatabaseProviderFactory();

            dbWorker = factory.CreateDefault();

            #endregion

            conWorker = dbWorker.CreateConnection();
        }

        /// <summary>
        /// 建構
        /// </summary>
        /// <param name="work">支持事務的資料庫操作DAO的委託</param>
        public DataBaseDAO(DatabaseWork work)
        {
            #region Legend 2016/10/26 因企業庫升級調整寫法

            //dbWorker = DatabaseFactory.CreateDatabase();

            DatabaseProviderFactory factory = new DatabaseProviderFactory();

            dbWorker = factory.CreateDefault();

            #endregion

            conWorker = dbWorker.CreateConnection();
            dwWork = work;
        }

        /// <summary>
        /// 建構
        /// </summary>
        /// <param name="connectionName">鏈接組態明</param>
        public DataBaseDAO(string connectionName)
        {
            #region Legend 2016/10/26 因企業庫升級調整寫法

            //dbWorker = DatabaseFactory.CreateDatabase(connectionName);

            DatabaseProviderFactory factory = new DatabaseProviderFactory();

            dbWorker = factory.Create(connectionName);

            #endregion

            conWorker = dbWorker.CreateConnection();
        }

        /// <summary>
        /// 建構
        /// </summary>
        /// <param name="connectionName">鏈接組態明</param>
        /// <param name="work">支持事務的資料庫操作DAO的委託</param>
        public DataBaseDAO(string connectionName, DatabaseWork work)
        {
            #region Legend 2016/10/26 因企業庫升級調整寫法

            //dbWorker = DatabaseFactory.CreateDatabase(connectionName);

            DatabaseProviderFactory factory = new DatabaseProviderFactory();

            dbWorker = factory.Create(connectionName);

            #endregion

            conWorker = dbWorker.CreateConnection();
            dwWork = work;
        }



        /// <summary>
        /// 執行委託方法，自動提交、回滾、關閉鏈接



        /// </summary>
        /// <param name="work">委託</param>
        /// <returns>如果操作出現異常則不為null</returns>
        public Exception BeginDatabaseWork(params object[] values)
        {
            Exception excReturn = null;
            if (TransactionBegined && Work != default(DatabaseWork))//如果事務已經開始
            {
                listLastCommands.Clear();
                try
                {
                    dwWork.Invoke(this, values);
                    trsWork.Commit();
                }
                catch (Exception ex)
                {
                    trsWork.Rollback();
                    excReturn = ex;
                }
                trsWork = null;
                conWorker.Close();
            }
            else
            {
                excReturn = new ApplicationException("TransactionBegined or Work not find!");
            }
            return excReturn;
        }


        /// <summary>
        /// 開啟連接，並且啟動事務



        /// </summary>
        public void OpenConnection()
        {
            try
            {
                conWorker.Open();
                trsWork = conWorker.BeginTransaction();
                blTransactionState = true;
            }
            catch { }
        }

        /// <summary>
        /// 提交
        /// </summary>
        public void Commit()
        {
            trsWork.Commit();
        }

        /// <summary>
        /// 回滾
        /// </summary>
        public void Rollback()
        {
            try
            {
                trsWork.Rollback();
            }
            catch { }
        }

        /// <summary>
        /// 關閉鏈接
        /// </summary>
        public void CloseConnection()
        {
            try
            {
                trsWork = null;
                conWorker.Close();
                blTransactionState = false;
            }
            catch { }
        }

        /// <summary>
        /// 執行SQL語句，返回第一行數值



        /// </summary>
        /// <param name="commandString">執行SQL語句</param>
        /// <param name="values">參數</param>
        /// <returns></returns>
        public object ExecuteScalar(string commandString, Dictionary<string, object> values)
        {
            DbCommand dbcWork = null;

            if (values == null)//如果參數為空
            {
                dbcWork = GetCommand(commandString, false);
            }
            else
            {
                dbcWork = GetCommand(commandString, false, values);
            }

            listLastCommands.Add(dbcWork);

            if (blTransactionState)
            {
                dbcWork.Transaction = trsWork;
                return dbcWork.ExecuteScalar();
            }
            else
            {
                return dbWorker.ExecuteScalar(dbcWork);
            }
        }

        /// <summary>
        /// 執行SQL語句
        /// </summary>
        /// <param name="commandText">SQL命令字符串</param>
        /// <param name="values">參數集合</param>
        /// <param name="isProcedure">是否是存儲過程</param>
        /// <returns>受操作的行數</returns>
        public int ExecuteNonQuery(string commandText, Dictionary<string, object> values, bool isProcedure)
        {
            int intRows = -1;

            DbCommand dbcWork = null;
            DataSet ds = new DataSet();
            if (values == null)//如果參數為空
            {
                dbcWork = GetCommand(commandText, false);
            }
            else
            {
                dbcWork = GetCommand(commandText, isProcedure, values);
            }

            listLastCommands.Add(dbcWork);

            if (blTransactionState)
            {
                dbcWork.Transaction = trsWork;
                intRows = dbcWork.ExecuteNonQuery();
            }
            else
            {
                intRows = dbWorker.ExecuteNonQuery(dbcWork);
            }

            //UserActiveLog.SaveLogByCmd(dbcWork);

            return intRows;
        }

        /// <summary>
        /// 執行SQL語句
        /// </summary>
        /// <param name="commandText">SQL命令字符串</param>
        /// <param name="values">參數集合</param>
        /// <returns>受操作的行數</returns>
        public int ExecuteNonQuery(string commandText, Dictionary<string, object> values)
        {
            return ExecuteNonQuery(commandText, values, false);
        }

        /// <summary>
        /// 執行SQL語句
        /// </summary>
        /// <param name="commandText">SQL命令字符串</param>
        /// <returns>受操作的行數</returns>
        public int ExecuteNonQuery(string commandText)
        {
            return ExecuteNonQuery(commandText, null, false);
        }

        /// <summary>
        /// 獲得資料集



        /// </summary>
        /// <param name="commandText">SQL命令字符串</param>
        /// <returns>得到的資料集</returns>
        public DataSet GetList(string commandText)
        {
            return GetList(commandText, null, false);
        }

        /// <summary>
        /// 獲得資料集



        /// </summary>
        /// <param name="commandText">SQL命令字符串</param>
        /// <param name="values">參數集合</param>
        /// <returns>得到的資料集</returns>
        public DataSet GetList(string commandText, Dictionary<string, object> values)
        {
            return GetList(commandText, values, false);
        }

        /// <summary>
        /// 從資料庫獲取模型資料列表（分頁）
        /// </summary>
        /// <param name="commandString">查詢命令字符串</param>
        /// <param name="firstRowNumber">開始行號</param>
        /// <param name="lastRowNumber">結束行號</param>
        /// <param name="sortField">排序欄位</param>
        /// <param name="sortType">排序方式</param>
        /// <param name="rowCount">查到的行數</param>
        /// <returns>資料集</returns>
        public DataSet GetList(string commandString, string firstRowNumber, string lastRowNumber, string sortField, string sortType, out int rowCount)
        {
            StringBuilder stbCommand = new StringBuilder("select count(*) from (");
            stbCommand.Append(commandString);
            stbCommand.Append(") table1 ");
            stbCommand.Append("select * from (select *,ROW_NUMBER() OVER (order by ");
            stbCommand.Append(sortField);
            stbCommand.Append(" ");
            stbCommand.Append(sortType);
            stbCommand.Append(" ) as RowNumber from (");
            stbCommand.Append(commandString);
            stbCommand.Append(") table2 ) table3 where RowNumber between ");
            stbCommand.Append(firstRowNumber);
            stbCommand.Append(" and ");
            stbCommand.Append(lastRowNumber);
            DataSet dtsReturn = GetList(stbCommand.ToString(), null, false);
            rowCount = (int)dtsReturn.Tables[0].Rows[0][0];
            dtsReturn.Tables.RemoveAt(0);
            return dtsReturn;
        }

        /// <summary>
        /// 從資料庫獲取模型資料列表（分頁）
        /// </summary>
        /// <param name="commandString">查詢命令字符串</param>
        /// <param name="parameters">參數集合</param>
        /// <param name="firstRowNumber">開始行號</param>
        /// <param name="lastRowNumber">結束行號</param>
        /// <param name="sortField">排序欄位</param>
        /// <param name="sortType">排序方式</param>
        /// <param name="rowCount">查到的行數</param>
        /// <returns>資料集</returns>
        public DataSet GetList(string commandString, Dictionary<string, object> parameters, string firstRowNumber, string lastRowNumber, string sortField, string sortType, out int rowCount)
        {
            StringBuilder stbCommand = new StringBuilder("select count(*) from (");
            stbCommand.Append(commandString);
            stbCommand.Append(") table1 ");
            stbCommand.Append("select * from (select *,ROW_NUMBER() OVER (order by ");
            stbCommand.Append(sortField);
            stbCommand.Append(" ");
            stbCommand.Append(sortType);
            stbCommand.Append(" ) as RowNumber from (");
            stbCommand.Append(commandString);
            stbCommand.Append(") table2 ) table3 where RowNumber between ");
            stbCommand.Append(firstRowNumber);
            stbCommand.Append(" and ");
            stbCommand.Append(lastRowNumber);
            DataSet dtsReturn = GetList(stbCommand.ToString(), parameters, false);
            rowCount = (int)dtsReturn.Tables[0].Rows[0][0];
            dtsReturn.Tables.RemoveAt(0);
            return dtsReturn;
        }

        /// <summary>
        /// 獲得資料集



        /// </summary>
        /// <param name="commandText">SQL命令字符串</param>
        /// <param name="values">參數集合</param>
        /// <param name="isProcedure">是否是存儲過程</param>
        /// <returns>得到的資料集</returns>
        public DataSet GetList(string commandText, Dictionary<string, object> values, bool isProcedure)
        {
            DataSet dstReturn = new DataSet();
            DbDataReader ddrReader = null;

            DbCommand dbcWork = null;
            DataSet ds = new DataSet();
            if (values == null)//如果參數為空
            {
                dbcWork = GetCommand(commandText, false);
            }
            else
            {
                dbcWork = GetCommand(commandText, isProcedure, values);
            }
            // 調整超時時間從600改為1800(0.5H)
            dbcWork.CommandTimeout = 1800;
            listLastCommands.Add(dbcWork);

            if (blTransactionState)
            {
                dbcWork.Transaction = trsWork;
                ddrReader = dbcWork.ExecuteReader();
                dstReturn.Load(ddrReader, LoadOption.Upsert, "Table1");
                ddrReader.Close();
            }
            else
            {
                dstReturn = dbWorker.ExecuteDataSet(dbcWork);
            }

            return dstReturn;
        }

        /// <summary>
        /// 向資料庫新增物件
        /// </summary>
        /// <typeparam name="T">數據模型</typeparam>
        /// <param name="model">待增加數據模型</param>
        public void Add<T>(T model)
        {
            Type typModel = typeof(T);
            PropertyInfo[] ppiModels = typModel.GetProperties();
            StringBuilder stbValues = new StringBuilder();
            StringBuilder stbFields = new StringBuilder();
            Dictionary<string, object> dirValues = new Dictionary<string, object>();
                       
            for (int index = 0; index < ppiModels.Length; index++)
            {
                object propertyValue = null;

               
                if (ppiModels[index].Name.ToUpper() == "RCT")
                    propertyValue = DateTime.Now;
                if (ppiModels[index].Name.ToUpper() == "RST")
                    propertyValue = GlobalString.RST.ACTIVED;               
                if (ppiModels[index].Name.ToUpper() == "RUT")
                    propertyValue = DateTime.Now;

                if (propertyValue == null)
                    propertyValue = ppiModels[index].GetValue(model, null);

                if (propertyValue == null)
                    continue;
                else
                    if (StringUtil.IsEmpty(propertyValue.ToString()))
                        propertyValue = " ";

                if (ppiModels[index].PropertyType.Name == "DateTime")
                    propertyValue = ((DateTime)propertyValue).ToString("yyyy-MM-dd HH:mm:ss");

                dirValues.Add(ppiModels[index].Name, propertyValue);

                stbFields.Append(ppiModels[index].Name);
                stbFields.Append(",");
                stbValues.Append("@");
                stbValues.Append(ppiModels[index].Name);
                stbValues.Append(",");
            }
            StringBuilder stbCommand = new StringBuilder("insert into ");
            stbCommand.Append(typModel.Name);
            stbCommand.Append(" (");
            stbCommand.Append(stbFields.ToString().Substring(0, stbFields.Length - 1));
            stbCommand.Append(") values (");
            stbCommand.Append(stbValues.ToString().Substring(0, stbValues.Length - 1));
            stbCommand.Append(")");

            ExecuteNonQuery(stbCommand.ToString(), dirValues, false);
        }

        /// <summary>
        /// 向資料庫新增物件
        /// </summary>
        /// <param name="model">待增加數據模型</param>
        /// <param name="strKeyId">主鍵名稱</param>
        public void Add<T>(T model, string strKeyId)
        {
            Type typModel = typeof(T);
            PropertyInfo[] ppiModels = typModel.GetProperties();
            StringBuilder stbValues = new StringBuilder();
            StringBuilder stbFields = new StringBuilder();
            Dictionary<string, object> dirValues = new Dictionary<string, object>();     
            
            for (int index = 0; index < ppiModels.Length; index++)
            {
                foreach (string strKey in strKeyId.Split(','))
                {
                    if (strKey.ToLower() != ppiModels[index].Name.ToLower())
                    {
                        object propertyValue = null;
                        propertyValue = ppiModels[index].GetValue(model, null);

                        if (propertyValue == null)
                        {
                            if (ppiModels[index].Name.ToUpper() == "RCU")
                                propertyValue = GlobalString.RCU.ACTIVED;
                            if (ppiModels[index].Name.ToUpper() == "RUU")
                                propertyValue = GlobalString.RUU.ACTIVED;

                        }
                       
                        if (ppiModels[index].Name.ToUpper() == "RCT")
                            propertyValue = DateTime.Now;
                        if (ppiModels[index].Name.ToUpper() == "RST")
                            propertyValue = GlobalString.RST.ACTIVED;                        
                        if (ppiModels[index].Name.ToUpper() == "RUT")
                            propertyValue = DateTime.Now;


                        if (propertyValue == null)
                            propertyValue = ppiModels[index].GetValue(model, null);

                        if (propertyValue == null)
                            continue;
                        else
                            if (StringUtil.IsEmpty(propertyValue.ToString()))
                                propertyValue = " ";

                        if (ppiModels[index].PropertyType.Name == "DateTime")
                            propertyValue = ((DateTime)propertyValue).ToString("yyyy-MM-dd HH:mm:ss");

                        dirValues.Add(ppiModels[index].Name, propertyValue);

                        stbFields.Append(ppiModels[index].Name);
                        stbFields.Append(",");
                        stbValues.Append("@");
                        stbValues.Append(ppiModels[index].Name);
                        stbValues.Append(",");

                    }
                }
            }
            StringBuilder stbCommand = new StringBuilder("insert into ");
            stbCommand.Append(typModel.Name);
            stbCommand.Append(" (");
            stbCommand.Append(stbFields.ToString().Substring(0, stbFields.Length - 1));
            stbCommand.Append(") values (");
            stbCommand.Append(stbValues.ToString().Substring(0, stbValues.Length - 1));
            stbCommand.Append(")");

            ExecuteNonQuery(stbCommand.ToString(), dirValues, false);
        }

        /// <summary>
        /// 新增，返回當前ID
        /// </summary>
        /// <typeparam name="T">數據模型</typeparam>
        /// <param name="model">待增加數據模型</param>
        /// <param name="strKeyId">需要返回的主鍵ID</param>
        /// <returns></returns>
        public object AddAndGetID<T>(T model, string strKeyId)
        {
            Add<T>(model, strKeyId);

            Type typModel = typeof(T);

            return ExecuteScalar("select max(" + strKeyId + ") from " + typModel.Name, null);
        }


        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="strTableName">表名</param>
        /// <param name="values">數據集</param>
        /// <returns></returns>
        public int Delete(string strTableName, Dictionary<string, object> values)
        {
            StringBuilder stbFields = new StringBuilder();
            foreach (KeyValuePair<string, object> entry in values)
            {
                stbFields.Append(entry.Key);
                stbFields.Append("=");
                stbFields.Append("@");
                stbFields.Append(entry.Key);
                stbFields.Append(" and ");
            }
            StringBuilder stbCommand = new StringBuilder("delete ");
            stbCommand.Append(strTableName);
            stbCommand.Append(" where ");
            stbCommand.Append(stbFields.ToString().Substring(0, stbFields.Length - 4));

            return ExecuteNonQuery(stbCommand.ToString(), values, false);
        }

        /// <summary>
        /// 修改
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="model">數據模型</param>
        /// <param name="keyName">條件字段，逗號隔開</param>
        public void Update<T>(T model, string keyName)
        {
            Type typModel = typeof(T);
            PropertyInfo[] ppiModels = typModel.GetProperties();
            StringBuilder stbFields = new StringBuilder();
            StringBuilder stbKeyFields = new StringBuilder();
            Dictionary<string, object> dirValues = new Dictionary<string, object>();
            for (int index = 0; index < ppiModels.Length; index++)
            {
                if (ppiModels[index].Name.ToUpper() == "RCU")
                    continue;
                if (ppiModels[index].Name.ToUpper() == "RCT")
                    continue;

                bool isKey = false;
                foreach (string strKeyName in keyName.Split(','))
                {
                    if (strKeyName == ppiModels[index].Name)
                    {
                        isKey = true;
                        stbKeyFields.Append(strKeyName);
                        stbKeyFields.Append("=@");
                        stbKeyFields.Append(strKeyName);
                        stbKeyFields.Append(" and ");

                        dirValues.Add(ppiModels[index].Name, ppiModels[index].GetValue(model, null));
                        break;
                    }
                }

                if (!isKey)
                {
                    object propertyValue = null;

                    //if (ppiModels[index].Name.ToUpper() == "RUU")
                    //    propertyValue = ((Users)HttpContext.Current.Session[GlobalString.SessionAndCookieKeys.USER]).UserID;
                    if (ppiModels[index].Name.ToUpper() == "RUT")
                        propertyValue = DateTime.Now;

                    if (propertyValue == null)
                        propertyValue = ppiModels[index].GetValue(model, null);

                    if (propertyValue == null)
                        continue;
                    else
                        if (StringUtil.IsEmpty(propertyValue.ToString()))
                            propertyValue = " ";

                    if (ppiModels[index].PropertyType.Name == "DateTime")
                        propertyValue = ((DateTime)propertyValue).ToString("yyyy-MM-dd HH:mm:ss");

                    dirValues.Add(ppiModels[index].Name, propertyValue);


                    stbFields.Append(ppiModels[index].Name);
                    stbFields.Append("=@");
                    stbFields.Append(ppiModels[index].Name);
                    stbFields.Append(",");
                }
            }
            StringBuilder stbCommand = new StringBuilder("update ");
            stbCommand.Append(typModel.Name);
            stbCommand.Append(" set ");
            stbCommand.Append(stbFields.ToString().Substring(0, stbFields.Length - 1));
            stbCommand.Append(" where ");
            stbCommand.Append(stbKeyFields.ToString().Substring(0, stbKeyFields.Length - 4));

            ExecuteNonQuery(stbCommand.ToString(), dirValues, false);
        }

        /// <summary>
        /// 獲取命令物件
        /// </summary>
        /// <param name="commandString">SQL命令</param>
        /// <param name="isProcedure">是否是存儲過程</param>
        /// <param name="useProcess">是否使用存儲過程</param>
        /// <returns>命令物件</returns>
        private DbCommand GetCommand(string commandString, bool isProcedure)
        {
            DbCommand dbcWork = (isProcedure ? dbWorker.GetStoredProcCommand(commandString) : dbWorker.GetSqlStringCommand(commandString));
            dbcWork.Connection = conWorker;
            return dbcWork;
        }

        /// <summary>
        /// 獲取命令物件
        /// </summary>
        /// <param name="commandString">SQL命令</param>
        /// <param name="isProcedure">是否是存儲過程</param>
        /// <param name="values">參數集合</param>
        /// <param name="useProcess">是否使用存儲過程</param>
        /// <returns>命令物件</returns>
        private DbCommand GetCommand(string commandString, bool isProcedure, Dictionary<string, object> values)
        {
            DbCommand dbcWork = GetCommand(commandString, isProcedure);
            FillParametersToCommand(dbcWork, values);
            return dbcWork;
        }

        /// <summary>
        /// 依據Dictionary<string, object>填充Command物件參數集



        /// </summary>
        /// <param name="command">命令物件</param>
        /// <param name="parameters">參數</param>
        private void FillParametersToCommand(IDbCommand command, Dictionary<string, object> parameters)
        {
            foreach (string parameterName in parameters.Keys)
            {
                IDataParameter parameter = command.CreateParameter();
                parameter.ParameterName = "@" + parameterName;
                if (parameters[parameterName] == null || parameters[parameterName].ToString() == "")//如果參數沒有值
                {
                    parameter.Value = DBNull.Value;
                }
                else
                {
                    parameter.Value = parameters[parameterName];
                }
                command.Parameters.Add(parameter);
            }
        }

        /// <summary>
        /// 清除命令日誌
        /// </summary>
        public void ClearLogCommand()
        {
            listLastCommands.Clear();
        }

        /// <summary>
        /// 從資料庫獲取模型物件
        /// </summary>
        /// <typeparam name="T">數據模型</typeparam>
        /// <typeparam name="TKey">主鍵字段</typeparam>
        /// <param name="strKeyName">主鍵名稱</param>
        /// <param name="keyValue">參數</param>
        /// <returns></returns>
        public T GetModel<T, TKey>(string strKeyName, TKey keyValue)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>(1);
            parameters.Add(strKeyName, keyValue);
            Type typModel = typeof(T);

            return GetModel<T>("select * from " + typModel.Name + " where " + string.Concat("[", strKeyName, "]=@", strKeyName), parameters);
        }

        /// <summary>
        /// 從資料庫獲取模型物件
        /// </summary>
        /// <typeparam name="T">數據模型</typeparam>
        /// <param name="commandString">SQL命令</param>
        /// <param name="parameters">參數</param>
        /// <returns>模型物件</returns>
        public T GetModel<T>(string commandString, Dictionary<string, object> parameters)
        {
            DataRow modelRow = GetRow(commandString, parameters, false);
            return GetModelByDataRow<T>(modelRow);
        }



        /// <summary>
        /// 依據一行模型資料生成模型物件



        /// </summary>
        /// <typeparam name="T">數據模型</typeparam>
        /// <param name="modelRow">資料行</param>
        /// <returns>模型物件</returns>
        public T GetModelByDataRow<T>(DataRow modelRow)
        {
            Type typModelType = typeof(T);//模型類型
            PropertyInfo[] piPropertyInfos = typModelType.GetProperties();//模型類型的屬性信息集合




            T retrunModel = (T)typModelType.Assembly.CreateInstance(typModelType.FullName);
            if (modelRow != null)
            {
                for (int index = 0; index < piPropertyInfos.Length; index++)
                {
                    try
                    {
                        object o = modelRow[piPropertyInfos[index].Name];
                    }
                    catch
                    {
                        continue;
                    }

                    object propertyValue = null;

                    if (piPropertyInfos[index].Name == "LastLoginDateTime")
                    {
                        propertyValue = DateTime.Now;
                    }
                    else
                    {
                        if (piPropertyInfos[index].PropertyType == typeof(string))
                        {
                            if (modelRow[piPropertyInfos[index].Name] != DBNull.Value)
                            {
                                propertyValue = modelRow[piPropertyInfos[index].Name].ToString().Trim();
                            }
                        }
                        else
                        {
                            if (!(modelRow[piPropertyInfos[index].Name] is System.DBNull))
                            {
                                if (!StringUtil.IsEmpty(modelRow[piPropertyInfos[index].Name].ToString()))
                                    propertyValue = Convert.ChangeType(modelRow[piPropertyInfos[index].Name], piPropertyInfos[index].PropertyType);
                            }

                        }
                    }
                    piPropertyInfos[index].SetValue(retrunModel, propertyValue, null);
                }

                return retrunModel;
            }
            else
            {
                //throw new Exception(GlobalStringManager.BizB["Alert_DatabaseNoneObjectErr"]);
                return default(T);
            }
        }

        /// <summary>
        ///  從資料庫獲取一行模型資料



        /// </summary>
        /// <param name="commandString">SQL命令</param>
        /// <param name="parameters">參數</param>
        /// <param name="useProcess">是否使用存儲過程</param>
        /// <returns>資料行</returns>
        public DataRow GetRow(string commandString, Dictionary<string, object> parameters, bool useProcess)
        {
            DataSet modelList = GetList(commandString, parameters, useProcess);
            if (modelList.Tables.Count > 0 && modelList.Tables[0].Rows.Count > 0)
            {
                return modelList.Tables[0].Rows[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 判斷資料中是否有符合條件的模型數據



        /// </summary>
        /// <param name="commandString">SQL命令</param>
        /// <param name="parameters">參數</param>
        /// <returns>存在否（true:存在，false:不存在）</returns>
        public bool Contains(string commandString, Dictionary<string, object> parameters)
        {
            object DatabaseReturn = 0;

            DatabaseReturn = ExecuteScalar(commandString, parameters);

            return ((((int)DatabaseReturn) > 0) ? true : false);
        }

        /// <summary>
        /// 判斷資料中是否有符合條件的模型數據



        /// </summary>
        /// <typeparam name="T">數據模型</typeparam>
        /// <typeparam name="TKey">字段類型</typeparam>
        /// <param name="strKeyName">驗證字段</param>
        /// <param name="keyValue">字段的值</param>
        /// <returns>存在否（true:存在，false:不存在）</returns>
        public bool Contains<T, TKey>(string strKeyName, TKey keyValue)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>();
            parameters.Add(strKeyName, keyValue);
            Type typModel = typeof(T);
            return Contains("select count(*) from " + typModel.Name + " where " + string.Concat("[", strKeyName, "]=@", strKeyName), parameters);
        }
    }
}
/// <summary>
/// 資料庫操作事務委託



/// </summary>
/// <param name="dbtTransaction">要使用的事務管理物件</param>
/// <param name="values">參數</param>
public delegate void DatabaseWork(DataBaseDAO dbtTransaction, params object[] values);
