using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace InfoChange
{
    class Logger_ufssprk_tools
    {
        private OleDbConnection gibdd_con;

        private string txtErrMessage;
        private string txtMemoryLog;

        public string ErrMessage
        {
            get
            {
                return txtErrMessage;
            }
            set
            {
                txtErrMessage = value;
            }
        }


        private string txtConStrGibdd;
        public string ConStr
        {
            get
            {
                return txtConStrGibdd;
            }
            set
            {
                if (value.Trim().Length > 0)
                {
                    txtConStrGibdd = value;
                }
            }
        }

        private decimal LLogID;
        public decimal logID
        {
            get
            {
                return LLogID;
            }
            set
            {
                LLogID = value;
            }
        }

        // параметры лога (требуются при создании)
        private decimal nStatus;
        public decimal Status
        {
            get
            {
                return nStatus;
            }
            set
            {
                nStatus = value;
            }
        }

        private int nPackType;
        public int PackType
        {
            get
            {
                return nPackType;
            }
            set
            {
                
                nPackType = value;
                
            }
        }

        private string txtAgreement_code;
        public string Agreement_code
        {
            get
            {
                return txtAgreement_code;
            }
            set
            {
                if (value.Length > 0)
                {
                    txtAgreement_code = value;
                }
            }
        }

        private decimal nParent_ID;
        public decimal Parent_ID
        {
            get
            {
                return nParent_ID;
            }
            set
            {
                nParent_ID = value;
                
            }
        }

        private bool bValidCon;
        public bool ValidCon
        {
            get
            {
                return bValidCon;
            }
            set
            {
                bValidCon = value;
            }
        }

        private decimal _OspNum;
        public decimal OspNum
        {
            get
            {
                return _OspNum;
            }
            set
            {
                _OspNum = value;
            }
        }

        private int _PackCount;
        public int PackCount
        {
            get
            {
                return _PackCount;
            }
            set
            {
                _PackCount = value;
            }
        }

        private string _Filename;
        public string Filename
        {
            get
            {
                return _Filename;
            }
            set
            {
                if (value.Length >= 0)
                {
                    _Filename = value;
                }
            }
        }

        public Logger_ufssprk_tools(string pConStr, decimal pLLogID)
        {
            txtMemoryLog = "";
            txtErrMessage = "";
            ConStr = pConStr;
            ValidCon = false;

            // проверить соединение
            ConStr = pConStr;
            ValidCon = CheckConnection();
            if (ValidCon)
            {
                gibdd_con = new OleDbConnection(ConStr);
            }
            // найти в базе данных такой лог и его свойства записать в параметры класса
            // будем параметры получать как DataTable

            string txtGetLogParams = "Select first 1 * from local_logs where ID = " + pLLogID.ToString();
            DataTable dtLogParams = null;
            dtLogParams = GetDataTableFromFB(txtGetLogParams, "LogParams");

            if (dtLogParams != null && dtLogParams.Rows.Count > 0)
            {
                DataRow row = dtLogParams.Rows[0];

                if (dtLogParams.Columns.Contains("OSPNUM"))
                    OspNum = Convert.ToDecimal(row["OSPNUM"]);

                if (dtLogParams.Columns.Contains("PACK_TYPE"))
                    PackType = Convert.ToInt32(row["PACK_TYPE"]);

                if (dtLogParams.Columns.Contains("CONV_CODE"))
                    Agreement_code = Convert.ToString(row["CONV_CODE"]);

                if (dtLogParams.Columns.Contains("PACK_STATUS"))
                    Status = Convert.ToDecimal(row["PACK_STATUS"]);

                if (dtLogParams.Columns.Contains("PACK_COUNT"))
                    PackCount = Convert.ToInt32(row["PACK_COUNT"]);

                if (dtLogParams.Columns.Contains("PARENT_ID"))
                    Parent_ID = Convert.ToDecimal(row["PARENT_ID"]);
                
                if (dtLogParams.Columns.Contains("FILENAME"))
                    Filename = Convert.ToString(row["FILENAME"]);

                LLogID = pLLogID;
            }
        }

        private DataTable GetDataTableFromFB(string txtSql, string tblName)
        {
            DataSet ds = new DataSet();
            DataTable tbl = ds.Tables.Add(tblName);
            try
            {
                // проверить подключение - а то может статься что не закрыли
                if (gibdd_con != null && gibdd_con.State != ConnectionState.Closed) gibdd_con.Close();
                gibdd_con.ConnectionString = ConStr;

                gibdd_con.Open();
                OleDbTransaction tran = gibdd_con.BeginTransaction(IsolationLevel.RepeatableRead);
                OleDbCommand cmd = new OleDbCommand(txtSql, gibdd_con, tran);

                using (OleDbDataReader rdr = cmd.ExecuteReader(CommandBehavior.Default))
                {
                    ds.Load(rdr, LoadOption.OverwriteChanges, tbl);
                    rdr.Close();
                }

                tran.Commit();
                gibdd_con.Close();

            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    ErrMessage += "Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState + ".\n";
                }
            }
            catch (Exception ex)
            {
                ErrMessage += "Ошибка приложения. Message: " + ex.ToString() + ".\n";
            }
            return tbl;
        }


        public decimal GetLogByFileName(int pPackType, string pAgreement_code, string txtFileName, int iMonthPeriod)
        {
            if (txtFileName.Trim().Length > 0)
            {

                string txtSql = "SELECT first 1 ID from LOCAL_LOGS where UPPER(FILENAME) LIKE '" + txtFileName.ToUpper() + "'";
                if (pAgreement_code.Trim().Length > 0) txtSql += " and CONV_CODE = '" + pAgreement_code + "'";
                if (pPackType != 0) txtSql += " and PACK_TYPE = " + pPackType;
                if (iMonthPeriod <= 0)
                {
                    DateTime dtPeriod = DateTime.Today.AddMonths(iMonthPeriod);
                    txtSql += " and '" + dtPeriod.ToShortDateString() + "' < packdate";
                }
                string txtResult = "";
                txtResult = SelectSqlScalar(gibdd_con, txtSql);
                // заглушка 
                if (txtResult.Trim() == "") return -1;
                return Convert.ToDecimal(txtResult);
            }
            return -1;

        }

        public Logger_ufssprk_tools(string pConStr, decimal pStatus, int pPackType, string pAgreement_code, decimal pParent_ID, decimal pOspNum, string txtLogMess)
        {
            PackCount = 0;
            txtMemoryLog = "";
            txtErrMessage = "";

            ConStr = pConStr;
            Status = pStatus;
            PackType = pPackType;
            Agreement_code = pAgreement_code;
            Parent_ID = pParent_ID;
            OspNum = pOspNum;

            ValidCon = false;
            
            // проверить соединение
            ConStr = pConStr;
            ValidCon = CheckConnection();

            // создать LLogID
            LLogID = CreateLLog(txtLogMess);
                        
            // вопрос о режиме записи:
            // отложить на потом, вот его суть - можно писать по запросу, а можно накопить и потом записать в конце
            // для этого нужно хранить режим записи - cummulative/permanent

        }

        private bool CheckConnection()
        {
            bool canConnect= false;
            OleDbConnection connection = new OleDbConnection(ConStr);

            try
            {
                using (connection)
                {
                    connection.Open();
                    canConnect = true;
                    gibdd_con = connection;
                }
            }
            catch (Exception exception)
            {
                ErrMessage += "Ошибка при проверке соединения. ConString = ''" + ConStr + "''\n";
                ErrMessage += "Message = ''" + exception.Message + "''\n";
            }
            finally
            {
                connection.Close();
            }

            return canConnect;

        }



        private decimal CreateLLog(string txtLOG)
        {
            if (!ValidCon) return 0; // проверить что соединение есть
            decimal nID = 0;
            decimal nOspNum = 0;
            OleDbCommand cmd, cmdIns;
            OleDbTransaction tran = null;

            try
            {
                if (gibdd_con != null && gibdd_con.State != ConnectionState.Closed) gibdd_con.Close();
                gibdd_con.ConnectionString = ConStr;
                gibdd_con.Open();
                tran = gibdd_con.BeginTransaction(IsolationLevel.ReadCommitted);

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(GEN_LOCAL_LOG_ID, 1) FROM RDB$DATABASE", gibdd_con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());

                // получить OSPNUM
                nOspNum = 10000 + OspNum;

                // вставить DOCUMENT
                // decimal nStatus, int nPackType, string txtAgreement_code, decimal nParent_ID,
                cmdIns = new OleDbCommand();
                cmdIns.Connection = gibdd_con;
                cmdIns.Transaction = tran;
                cmdIns.CommandText = "insert into LOCAL_LOGS (ID, OSPNUM, PACKDATE, PACK_TYPE, CONV_CODE, PACK_STATUS, PACK_COUNT, PARENT_ID, LOG)";
                cmdIns.CommandText += " VALUES (:ID ,:OSPNUM, :PACKDATE, :PACK_TYPE, :CONV_CODE, :PACK_STATUS, 0, :PARENT_ID, :TXT_LOG)";

                cmdIns.Parameters.Add(new OleDbParameter(":ID", Convert.ToDecimal(nID)));
                cmdIns.Parameters.Add(new OleDbParameter(":OSPNUM", Convert.ToInt32(nOspNum)));
                cmdIns.Parameters.Add(new OleDbParameter(":PACKDATE", DateTime.Now));
                cmdIns.Parameters.Add(new OleDbParameter(":PACK_TYPE", PackType));
                cmdIns.Parameters.Add(new OleDbParameter(":CONV_CODE", Agreement_code));
                cmdIns.Parameters.Add(new OleDbParameter(":PACK_STATUS", Status));
                cmdIns.Parameters.Add(new OleDbParameter(":PARENT_ID", Parent_ID));
                cmdIns.Parameters.Add(new OleDbParameter(":TXT_LOG", txtLOG));

                if (cmdIns.ExecuteNonQuery() == -1)
                {
                    Exception ex = new Exception("Error inserting new row to document table LOCAL_LOGS = " + nID.ToString());
                    throw ex;
                }

                tran.Commit();
                gibdd_con.Close();

            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    ErrMessage += "Ошибка при работе с данными.\n";
                    ErrMessage += "Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState + ".\n";
                }
            }
            catch (Exception ex)
            {
                ErrMessage += "Ошибка приложения. Message: " + ex.ToString() + "\n";
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
            }

            return nID;
        }

        // функция для отложенной записи в лог - чтобы записывать сразу только отрицательные записи,
        // а положительные хранить и писать в конце

        public bool MemoryLLog(string txtText)
        {
            txtMemoryLog += txtText;
            return true;
        }

        //public bool WriteLLog(string txtText)
        //{
        //    bool res;
        //    // если есть отложенная запись - записать.
        //    if (txtMemoryLog.Length > 0) txtText = txtMemoryLog + txtText;

        //    txtText = txtText.Replace("'", "''"); // эскапируем кавычки
        //    string txtSql = "Update LOCAL_LOGS set LOG = LOG || '" + txtText + "' where ID = " + LLogID.ToString();
        //    res = UpdateSqlExecute(gibdd_con, txtSql);
            
        //    // если отложенная запись прошла - очистить txtMemoryLog
        //    if (res) txtMemoryLog = "";

        //    return res;
        //}

        public bool WriteTofile(string txtText, string outfile)
        {
            using (StreamWriter sw = new StreamWriter(outfile, true))
            {
                sw.WriteLine(txtText);
                sw.Close();
            }
            return true;

        }


        public bool WriteLLog(string txtText)
        {
            bool res;
            // если есть отложенная запись - записать.
            if (txtMemoryLog.Length > 0) txtText = txtMemoryLog + txtText;
            
            // проверить длину (д.б. не больше чем 32768 (UTF-8) или 65536 (cp1251)? - 64 (длина остального SQL)
            // 100 отведу на эскапирование кавычек - вдруг их там тьма
            int nMaxLen = 65536 - 164;
            string txtSql = "";

            while (txtText.Length > nMaxLen)
            {
                string txtSubstr = txtText.Substring(0, nMaxLen);
                txtSubstr = txtSubstr.Replace("'", "''"); // эскапируем кавычки
                txtSql = "Update LOCAL_LOGS set LOG = LOG || '" + txtSubstr + "' where ID = " + LLogID.ToString();
                UpdateSqlExecute(gibdd_con, txtSql);
                txtText = txtText.Substring(nMaxLen);
            }
            // последняя итерация
            txtText = txtText.Replace("'", "''"); // эскапируем кавычки
            txtSql = "Update LOCAL_LOGS set LOG = LOG || '" + txtText + "' where ID = " + LLogID.ToString();
            res = UpdateSqlExecute(gibdd_con, txtSql);

            // если отложенная запись прошла - очистить txtMemoryLog
            if (res) txtMemoryLog = "";
            return res;
        }

        public bool UpdateLLogStatus(int pStatus)
        {
            Status = pStatus;
            string txtSql = "Update LOCAL_LOGS set PACK_STATUS = " + Status.ToString() + " where ID = " + LLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        public bool UpdateLLogFileName(string txtFileName)
        {
            bool res = false;
            if (txtFileName.Length > 300) txtFileName = txtFileName.Substring(0, 300);
            string txtSql = "UPDATE LOCAL_LOGS ll set ll.filename = '" + txtFileName + "' where ID = " + LLogID.ToString();
            res = UpdateSqlExecute(gibdd_con, txtSql);
            if(res) Filename = txtFileName;
            return res;
        }


        public bool UpdateLLogCount(int pCount)
        {
            PackCount = pCount;
            string txtSql = "Update LOCAL_LOGS set PACK_COUNT = " + PackCount.ToString() + " where ID = " + LLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        public bool AppendLLogCount(int pCount)
        {
            return UpdateLLogCount(PackCount + pCount);
        }

        public bool UpdateLLogParent(decimal pParentID)
        {
            Parent_ID = pParentID;
            string txtSql = "Update LOCAL_LOGS set PARENT_ID = " + Parent_ID.ToString() + " where ID = " + LLogID.ToString();
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        public bool UpdateLLogParentStatus(int pStatus)
        {
            string txtSql = "Update LOCAL_LOGS set PACK_STATUS = " + pStatus.ToString() + " where ID in (select PARENT_ID from LOCAL_LOGS where ID = " + LLogID.ToString() + ")";
            return UpdateSqlExecute(gibdd_con, txtSql);
        }

        public bool UpdateLLogFlag(int Status, string txtFlagName)
        {
            string txtSql = "";

            switch (txtFlagName)
            {
                case "FL_FIND":
                    txtSql = "Update LOCAL_LOGS set FL_FIND = " + Status.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + LLogID.ToString() + ")";
                    break;
                case "FL_NOFIND":
                    txtSql = "Update LOCAL_LOGS set FL_NOFIND = " + Status.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + LLogID.ToString() + ")";
                    break;
                case "FL_E_TOFIND":
                    txtSql = "Update LOCAL_LOGS set FL_E_TOFIND = " + Status.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + LLogID.ToString() + ")";
                    break;
            }

            if (txtSql != "")
                return UpdateSqlExecute(gibdd_con, txtSql);
            else return false;
        }

        public decimal GetLLogStatus(decimal pLLogID)
        {
            string txtSql = "SELECT PACK_STATUS from LOCAL_LOGS where ID = " + pLLogID.ToString();
            string txtResult = "";
            txtResult = SelectSqlScalar(gibdd_con, txtSql);
            // заглушка 
            if (txtResult.Trim() == "") return 0;
            return Convert.ToDecimal(txtResult);
        }

        public string SelectSqlScalar(OleDbConnection con, string txtSql)
        {
            string res = "";
            try
            {
                if (con != null && con.State != ConnectionState.Closed) con.Close();
                con.Open();
                OleDbTransaction tran = con.BeginTransaction(IsolationLevel.ReadCommitted);
                OleDbCommand cmd = new OleDbCommand(txtSql, con, tran);
                res = Convert.ToString(cmd.ExecuteScalar());
                tran.Rollback();
                con.Close();
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    ErrMessage += "Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState + ".\n";
                }
            }
            catch (Exception ex)
            {
                ErrMessage += "Ошибка приложения. Message: " + ex.ToString() + ".\n";
            }
            return res;
        }


        private bool UpdateLLogFlag(decimal pLLogID, int pStatus, string txtFlagName)
        {
            string txtSql = "";

            switch (txtFlagName)
            {
                case "FL_FIND":
                    txtSql = "Update LOCAL_LOGS set FL_FIND = " + pStatus.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + pLLogID.ToString() + ")";
                    break;
                case "FL_NOFIND":
                    txtSql = "Update LOCAL_LOGS set FL_NOFIND = " + pStatus.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + pLLogID.ToString() + ")";
                    break;
                case "FL_E_TOFIND":
                    txtSql = "Update LOCAL_LOGS set FL_E_TOFIND = " + pStatus.ToString() + " where ID = (select first 1 parent_id from local_logs where id =" + pLLogID.ToString() + ")";
                    break;
            }

            if (txtSql != "")
                return UpdateSqlExecute(gibdd_con, txtSql);
            else return false;
        }

        private decimal CopyLLogParent(decimal pOldLLogID, string txtNewAgreementCode)
        {
            bool bUpdated = true;
            decimal nID = 0;
            OleDbCommand cmd, cmdIns;
            OleDbTransaction tran = null;

            try
            {
                if (gibdd_con != null && gibdd_con.State != ConnectionState.Closed) gibdd_con.Close();
                gibdd_con.ConnectionString = ConStr;
                gibdd_con.Open();
                tran = gibdd_con.BeginTransaction(IsolationLevel.ReadCommitted);

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(GEN_LOCAL_LOG_ID, 1) FROM RDB$DATABASE", gibdd_con, tran);
                nID = Convert.ToDecimal(cmd.ExecuteScalar());

                if (nID * pOldLLogID != 0)
                {
                    cmdIns = new OleDbCommand();
                    cmdIns.Connection = gibdd_con;
                    cmdIns.Transaction = tran;
                    cmdIns.CommandText = "insert into LOCAL_LOGS (ID, OSPNUM, PACKDATE, PACK_TYPE, CONV_CODE, PACK_STATUS, PACK_COUNT, PARENT_ID,LOG) select :new_ID as ID, OSPNUM, PACKDATE, PACK_TYPE, :newAgrCode as CONV_CODE, PACK_STATUS, PACK_COUNT, PARENT_ID, LOG from LOCAL_LOGS WHERE ID = :old_ID";

                    cmdIns.Parameters.Add(new OleDbParameter(":new_ID", nID));
                    cmdIns.Parameters.Add(new OleDbParameter(":newAgrCode", txtNewAgreementCode));
                    cmdIns.Parameters.Add(new OleDbParameter(":old_ID", pOldLLogID));

                    if (cmdIns.ExecuteNonQuery() == -1)
                    {
                        bUpdated = false;
                    }
                }

                tran.Commit();
                gibdd_con.Close();

                if (!bUpdated)
                {
                    Exception ex = new Exception("Error. Can't copy doc_deposit id = " + pOldLLogID.ToString());
                    throw ex;
                }
            }
            catch (OleDbException ole_ex)
            {
                foreach (OleDbError err in ole_ex.Errors)
                {
                    ErrMessage += "Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState + ".\n";
                }
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
            }
            catch (Exception ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
                ErrMessage += "Ошибка приложения. Message: " + ex.ToString() + ".\n";
            }
            return nID;

        }

        private bool UpdateSqlExecute(OleDbConnection con, string txtUpdSql)
        {

            OleDbCommand cmd;
            OleDbTransaction tran = null;

            try
            {
                if ((con == null) || (con.State.Equals(ConnectionState.Closed)))
                {
                    con.Open();
                }

                tran = con.BeginTransaction(IsolationLevel.ReadCommitted);

                cmd = new OleDbCommand();
                cmd.Connection = con;
                cmd.Transaction = tran;
                cmd.CommandText = txtUpdSql;
                cmd.ExecuteNonQuery();
                tran.Commit();
                con.Close();

                return true;
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (con != null)
                {
                    con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    ErrMessage += "Ошибка при работе с данными. Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState + "\n";
                }
            }
            catch (Exception ex)
            {
                ErrMessage += "Ошибка приложения. Message: " + ex.ToString() + "\n"; ;
                if (con != null)
                {
                    con.Close();
                }
            }
            return false;
        }

        public decimal GenExtID()
        {
            decimal nExtID = 0;

            if (!ValidCon) return 0; // проверить что соединение есть
            OleDbCommand cmd;
            OleDbTransaction tran = null;

            try
            {
                if (gibdd_con != null && gibdd_con.State != ConnectionState.Closed) gibdd_con.Close();
                gibdd_con.ConnectionString = ConStr;
                gibdd_con.Open();
                tran = gibdd_con.BeginTransaction(IsolationLevel.ReadCommitted);

                // получить новый ключ
                cmd = new OleDbCommand("SELECT gen_id(GEN_EXT_ID, 1) FROM RDB$DATABASE", gibdd_con, tran);
                nExtID = Convert.ToDecimal(cmd.ExecuteScalar());
                tran.Commit();
                gibdd_con.Close();
            }
            catch (OleDbException ole_ex)
            {
                if (tran != null)
                {
                    tran.Rollback();
                }
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
                foreach (OleDbError err in ole_ex.Errors)
                {
                    ErrMessage += "Ошибка при работе с данными.\n";
                    ErrMessage += "Message: " + err.Message + "Native Error: " + err.NativeError + "Source: " + err.Source + "SQL State   : " + err.SQLState + ".\n";
                }
            }
            catch (Exception ex)
            {
                ErrMessage += "Ошибка приложения. Message: " + ex.ToString() + "\n";
                if (gibdd_con != null)
                {
                    gibdd_con.Close();
                }
            }

            return nExtID;
        }

    }
}
