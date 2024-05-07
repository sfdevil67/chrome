using MySql.Data.MySqlClient;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ChromeBrowserMacro
{
    public class Common
    {
        /// <summary>
        /// ConnectionString
        /// </summary>
        public static string DBConnectString = string.Empty;

        /// <summary>
        /// Domain
        /// </summary>
        public static string Domain = string.Empty;

        /// <summary>
        /// DocumentURL
        /// </summary>
        public static string DocumentURL = string.Empty;

        /// <summary>
        /// HtmlHeader
        /// </summary>
        public static string HtmlHeader = string.Empty;

        /// <summary>
        /// HtmlFooter
        /// </summary>
        public static string HtmlFooter = string.Empty;

        /// <summary>
        /// WorkingFolder
        /// </summary>
        public static string WorkingFolder = string.Empty;

        /// <summary>
        /// HtmlDocPath
        /// </summary>
        public static string HtmlDocPath = string.Empty;

        /// <summary>
        /// PDFDocPath
        /// </summary>
        public static string PDFDocPath = string.Empty;

        /// <summary>
        /// PDF_Program_PATH
        /// </summary>
        public static string PDF_Program_PATH = string.Empty;

        /// <summary>
        /// LogPath
        /// </summary>
        public static string LogPath = string.Empty;

        /// <summary>
        /// ErrorLogPath
        /// </summary>
        public static string ErrorLogPath = string.Empty;

        /// <summary>
        /// LoginID
        /// </summary>
        public static string LoginID = string.Empty;

        /// <summary>
        /// LoginPWD
        /// </summary>
        public static string LoginPWD = string.Empty;

        /// <summary>
        /// SynapIP
        /// </summary>
        public static string SynapIP = string.Empty;

        /// <summary>
        /// SynapConvertURL
        /// </summary>
        public static string SynapConvertURL = string.Empty;

        /// <summary>
        /// SynapHtmlPath
        /// </summary>
        public static string SynapHtmlPath = string.Empty;

        /// <summary>
        /// SynapLoginID
        /// </summary>
        public static string SynapLoginID = string.Empty;

        /// <summary>
        /// SynapLoginPWD
        /// </summary>
        public static string SynapLoginPWD = string.Empty;

        public Common() { Init(); }

        #region { Init : 초기 세팅 }
        /// <summary>
        /// 프로그램에 필요한 초기 세팅한다.
        /// </summary>
        protected static void Init()
        {
            try
            {
                // App Config 조회
                DBConnectString = GetAppConfigValue("DBConnectString");
                Domain = GetAppConfigValue("Domain");
                DocumentURL = GetAppConfigValue("DocumentURL").Replace("{Domain}", Domain);
                HtmlHeader = GetAppConfigValue("HtmlHeader").Replace("{Domain}", Domain);
                HtmlFooter = GetAppConfigValue("HtmlFooter");
                WorkingFolder = GetAppConfigValue("WorkingFolder");
                HtmlDocPath = GetAppConfigValue("HtmlDocPath");
                PDFDocPath = GetAppConfigValue("PDFDocPath");
                PDF_Program_PATH = GetAppConfigValue("PDF_Program_PATH");
                LogPath = GetAppConfigValue("LogPath");
                ErrorLogPath = GetAppConfigValue("ErrorLogPath");
                LoginID = GetAppConfigValue("LoginID");
                LoginPWD = GetAppConfigValue("LoginPWD");
                SynapIP = GetAppConfigValue("SynapIP");
                SynapConvertURL = GetAppConfigValue("SynapConvertURL");
                SynapHtmlPath = GetAppConfigValue("SynapHtmlPath");
                SynapLoginID = GetAppConfigValue("SynapLoginID");
                SynapLoginPWD = GetAppConfigValue("SynapLoginPWD");
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
                Application.Exit();
            }
        }
        #endregion

        #region { RemoveTag : 불필요한 태그 제거 }
        /// <summary>
        /// 불필요한 태그 제거
        /// </summary>
        /// <param name="pDocument">pDocument</param>
        /// <param name="pPattern">pPattern</param>
        /// <returns>태그 제거된 본문 ( STIRNG )</returns>
        public static string RemoveTag(string pDocument, string pPattern)
        {
            string returnDocument = string.Empty;
            try
            {
                returnDocument = Regex.Replace(pDocument, pPattern, "", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
            return returnDocument;
        }
        #endregion

        #region { ReplaceTag : 태그변경 }
        /// <summary>
        /// 태그변경
        /// </summary>
        /// <param name="pDocument">pDocument</param>
        /// <param name="pPattern">pPattern</param>
        /// <param name="pChageString">pChageString</param>
        /// <returns>태그 제거된 본문 ( STIRNG )</returns>
        public static string ReplaceTag(string pDocument, string pPattern, string pChageString)
        {
            string returnDocument = string.Empty;
            try
            {
                returnDocument = Regex.Replace(pDocument, pPattern, pChageString, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
            return returnDocument;
        }
        #endregion

        #region { GetAppConfigValue : App Config 값 조회 }
        /// <summary>
        /// App Config 값을 조회 한다.
        /// </summary>
        /// <param name="pKey"></param>
        /// <returns></returns>
        private static string GetAppConfigValue(string pKey)
        {
            if (string.IsNullOrWhiteSpace(pKey)) return "";
            string ret = string.Empty;
            try
            {
                ret = ConfigurationManager.AppSettings[pKey].ToString();
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
            return ret;
        }
        #endregion

        #region { WriteErrorLog : Log 파일 작성 }
        /// <summary>
        /// Error Log 파일을 작성 한다.
        /// </summary>
        /// <param name="pMethod">Call Method Name</param>
        /// <param name="pException">Exception</param>
        public static void WriteErrorLog(string pMethod, Exception pException, string pErrorMsg = "")
        {
            try
            {
                string logPath = ErrorLogPath;
                string message = !string.IsNullOrWhiteSpace(pErrorMsg) ? pErrorMsg : pException.Message + Environment.NewLine + pException.StackTrace;

                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }

                string file = String.Format("{0}{1}.log", logPath, DateTime.Now.ToString("yyyy-MM-dd"));

                using (StreamWriter sw = new StreamWriter(file, true, Encoding.UTF8))
                {
                    sw.WriteLine(String.Format("{0}:{1} {2} => {3}", DateTime.Now.ToString("HH:mm:ss"), DateTime.Now.Millisecond.ToString(), pMethod, message));
                }
            }
            catch { }
        }
        #endregion

        #region { WriteTextLog : Log 파일 작성 }
        /// <summary>
        /// Log 파일을 작성 한다.
        /// </summary>
        /// <param name="pMmessage">Mmessage</param>
        public static void WriteTextLog(string pMmessage)
        {
            try
            {
                string logPath = LogPath;
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }

                string file = String.Format("{0}{1}.log", logPath, DateTime.Now.ToString("yyyy-MM-dd"));

                using (StreamWriter sw = new StreamWriter(file, true, Encoding.UTF8))
                {
                    sw.WriteLine(pMmessage);
                }
            }
            catch { }
        }
        #endregion

        #region { SetConvertStatus : 변환작업 상태 변경 (P:진행중/S:완료/E:오류) }
        /// <summary>
        /// 변환작업 상태 변경 (P:진행중/S:완료/E:오류)
        /// </summary>
        public static void SetConvertStatus(string status, string msg, string uniqueID, string mhtLinkUrl = "")
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(uniqueID))
                {
                    using (MySqlCommand oCommand = new MySqlCommand())
                    {
                        // 변환할 데이터 상태 변경함 (P:진행중)
                        if (status.Equals("E"))
                        {
                            oCommand.CommandText = String.Format("UPDATE covi_approval4j_store.migration_docinfo SET CONVERTSTATUS = '{0}', CONVERTMSG = '{1}', MHTLINKURL = '{3}' WHERE UNIQUEID = '{2}'", status, msg, uniqueID, mhtLinkUrl);
                        }
                        else
                        {
                            oCommand.CommandText = String.Format("UPDATE covi_approval4j_store.migration_docinfo SET CONVERTSTATUS = '{0}', BODY_TYPE = '{1}', MHTLINKURL = '{3}' WHERE UNIQUEID = '{2}'", status, msg, uniqueID, mhtLinkUrl);
                        }
                        Common.ExecuteNonQueryMySql(oCommand);
                    }
                }
            }
            catch (Exception ex)
            {
                Common.WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
        }
        #endregion

        #region { ExecuteDataSet : SqlCommand 실행 후 DataSet 반환 }
        /// <summary>
        /// DB연결 후 SqlCommand 실행하고 DataSet을 반환한다.
        /// </summary>
        /// <param name="pCommand"></param>
        /// <returns></returns>
        public static DataSet ExecuteDataSet(SqlCommand pCommand)
        {
            DataSet ds = new DataSet();
            try
            {
                using (SqlConnection oConnection = new SqlConnection(DBConnectString))
                {
                    pCommand.CommandType = CommandType.StoredProcedure;
                    pCommand.Connection = oConnection;
                    oConnection.Open();
                    pCommand.ExecuteNonQuery();

                    using (SqlDataAdapter oSqlDataAdapter = new SqlDataAdapter(pCommand))
                    {
                        oSqlDataAdapter.Fill(ds);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
            return ds;
        }
        #endregion

        #region { ExecuteNonQuery : SqlCommand 실행 }
        /// <summary>
        /// DB연결 후 pCommand 실행한다.
        /// </summary>
        /// <param name="pSqlCommand"></param>
        public static void ExecuteNonQuery(SqlCommand pCommand)
        {
            try
            {
                using (SqlConnection oConnection = new SqlConnection(DBConnectString))
                {
                    pCommand.CommandType = CommandType.StoredProcedure;
                    pCommand.Connection = oConnection;
                    oConnection.Open();
                    pCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
        }
        #endregion

        #region { ExecuteDataSetMySql : SqlCommand 실행 후 DataSet 반환 }
        /// <summary>
        /// DB연결 후 SqlCommand 실행하고 DataSet을 반환한다.
        /// </summary>
        /// <param name="pCommand"></param>
        /// <returns></returns>
        public static DataSet ExecuteDataSetMySql(MySqlCommand pCommand)
        {
            DataSet ds = new DataSet();
            try
            {
                using (MySqlConnection oConnection = new MySqlConnection(DBConnectString))
                {
                    pCommand.CommandType = CommandType.Text;
                    pCommand.Connection = oConnection;
                    oConnection.Open();
                    pCommand.ExecuteNonQuery();

                    using (MySqlDataAdapter oDataAdapter = new MySqlDataAdapter(pCommand))
                    {
                        oDataAdapter.Fill(ds);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
            return ds;
        }
        #endregion

        #region { ExecuteNonQueryMySql : SqlCommand 실행 }
        /// <summary>
        /// DB연결 후 SqlCommand 실행한다.
        /// </summary>
        /// <param name="pSqlCommand"></param>
        public static void ExecuteNonQueryMySql(MySqlCommand pCommand)
        {
            try
            {
                using (MySqlConnection oConnection = new MySqlConnection(DBConnectString))
                {
                    pCommand.CommandType = CommandType.Text;
                    pCommand.Connection = oConnection;
                    oConnection.Open();
                    pCommand.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                WriteErrorLog(System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), ex);
            }
        }
        #endregion

    }
}