//******************************************************************
//*  作    者：Ray
//*  功能說明：FTP操作
//*  創建日期：2008-11-14
//*  修改日期：
//*  修改記錄：
//*******************************************************************
using System;
using System.Net;
using System.IO;
using System.Text;
using System.Configuration;

namespace CIMSClass.FTP
{

    public class FTPFactory
    {
        #region Attribute
        private string remoteHost, remotePath, remoteUser, remotePass, locPath;
        private int remotePort;

        private FtpWebRequest ftp;
        private WebResponse response;
        const int bufferSize = 2048; //下載緩沖2M
        #endregion
        #region 構造函數
        public FTPFactory()
        {
            remoteHost = ConfigurationManager.AppSettings["FTPRemoteIP"]; //FTP IP地址
            remotePath = ConfigurationManager.AppSettings["FTPRemotePathDefault"]; //FTP 檔案目錄
            remoteUser = ConfigurationManager.AppSettings["FTPUser"]; //FTP User
            remotePass = ConfigurationManager.AppSettings["FTPPassword"]; //FTP Password
            remotePort = Convert.ToInt32(ConfigurationManager.AppSettings["FTPRemotePort"]); //FTP Port
            locPath = ConfigurationManager.AppSettings["SubTotalFilesPath"]; // Local 檔案存儲目錄
        }


        public FTPFactory(string FtpType)
        {
            remoteHost = ConfigurationManager.AppSettings[FtpType + "FTPRemoteIP"]; //FTP IP地址
            remotePath = ConfigurationManager.AppSettings["FTPRemotePathDefault"]; //FTP 檔案目錄
            remoteUser = ConfigurationManager.AppSettings[FtpType + "FTPUser"]; //FTP User
            remotePass = ConfigurationManager.AppSettings[FtpType + "FTPPassword"]; //FTP Password
            remotePort = Convert.ToInt32(ConfigurationManager.AppSettings[FtpType + "FTPRemotePort"]); //FTP Port
            locPath = ConfigurationManager.AppSettings["SubTotalFilesPath"]; // Local 檔案存儲目錄
        }
        #endregion
        #region 登入FTP
        /// <summary>
        /// 登陸FTP
        /// </summary>
        /// <param name="Path">FTP子路徑</param>
        /// <returns></returns>
        private bool Login(string Path)
        {
            return Login(Path, "");

        }
        /// <summary>
        /// 登陸FTP
        /// </summary>
        /// <param name="Path">FTP子路徑</param>
        /// <param name="remFileName">FTP檔名</param>
        /// <returns></returns>
        private bool Login(string Path, string remFileName)
        {
            return Login(remoteHost, Path, remFileName);

        }
        private bool Login(string Host, string Path, string remFileName)
        {
            return Login(Host, Path, remFileName, remoteUser, remotePass);

        }
        private bool Login(string Host, string Path, string remFileName, string remoteUser, string remotePass)
        {
            try
            {
                string FTPURL = "ftp://" + Host + "/";
                if (!StringUtil.IsEmpty(Path))
                {
                    FTPURL = FTPURL + Path + "/" + remFileName;
                }

                ftp = (FtpWebRequest)FtpWebRequest.Create(FTPURL);
                ftp.UseBinary = true;
                ftp.Credentials = new NetworkCredential(remoteUser, remotePass);
                ftp.KeepAlive = false;
                if (ConfigurationManager.AppSettings["FTPModel"].ToString() == "2")
                    ftp.UsePassive = false;
                return true;
            }
            catch (Exception ex)
            {
                LogFactory.Write("FTP登陸失敗：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }
        #endregion
        #region 取得FTP檔案列表
        /// <summary>
        /// 取得檔案列表
        /// </summary>
        /// <returns></returns>
        public string[] GetFileList()
        {
            return GetFileList(remotePath);
        }
        /// <summary>
        /// 取得檔案列表
        /// </summary>
        /// <param name="remotePath">FTP子目錄</param>
        /// <returns></returns>
        public string[] GetFileList(string remotePath)
        {
            return GetFileList(remotePath, "");
        }
        /// <summary>
        /// 取得檔案列表
        /// </summary>
        /// <param name="remotePath">FTP子目錄</param>
        /// <param name="remFileName">FTP檔案名稱</param>
        /// <returns></returns>
        public string[] GetFileList(string remotePath, string remFileName)
        {
            string line;
            StringBuilder dicList = new StringBuilder();
            if (true)
            {
                Login(remotePath, remFileName);

            }

            try
            {
                ftp.Method = WebRequestMethods.Ftp.ListDirectory;
                WebResponse response = ftp.GetResponse();
                StreamReader sr = new StreamReader(response.GetResponseStream());
                line = sr.ReadLine();
                while (line != null)
                {
                    dicList.Append(line);
                    dicList.Append("\n");
                    line = sr.ReadLine();
                }
                sr.Close();
                response.Close();
                return dicList.ToString().Split('\n');

            }
            catch (Exception ex)
            {
                LogFactory.Write("獲取FTP目錄remotePath:" + remotePath + "remFileName：" + remFileName + "錯誤：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return null;
            }
        }
        #endregion
        #region 下載檔案
        /// <summary>
        /// 下載FTP檔案
        /// </summary>
        /// <param name="remotePath">FTP子目錄名稱</param>
        /// <param name="remFileName">FTP檔案名稱</param>
        /// <param name="locFileName">Local檔案名稱</param>
        /// <returns></returns>
        public bool Download(string remotePath, string remFileName, string locPath, string locFileName)
        {
            int readCount;
            byte[] buffer = new byte[bufferSize];
            try
            {
                if (GetFileList(remotePath, remFileName) == null)
                {
                    LogFactory.Write("檔案 未找到！", GlobalString.LogType.ErrorCategory);
                    return false;
                }
                MKDir(locPath);
                FileStream fs = new FileStream(locPath + locFileName, FileMode.Create);
                if (true)
                {
                    Login(remotePath, remFileName);

                }
                ftp.Method = WebRequestMethods.Ftp.DownloadFile;
                response = (FtpWebResponse)ftp.GetResponse();
                Stream stream = response.GetResponseStream();
                readCount = stream.Read(buffer, 0, bufferSize);
                while (readCount > 0)
                {
                    fs.Write(buffer, 0, readCount);
                    readCount = stream.Read(buffer, 0, bufferSize);
                }
                stream.Close();
                fs.Close();
                response.Close();
                return true;
            }
            catch (Exception ex)
            { 
                LogFactory.Write("下載檔案 失敗：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }

        }
        /// <summary>
        /// 下載FTP檔案
        /// </summary>
        /// <param name="remFileName"></param>
        /// <param name="locFileName"></param>
        /// <returns></returns>
        public bool Download(string remFileName, string locFileName)
        {
            return Download(remFileName, locPath, locFileName);

        }
        public bool Download(string remFileName, string locPath, string locFileName)
        {
            return Download(remotePath, remFileName, locPath, locFileName);

        }
        #endregion
        #region 創建目錄
        /// <summary>
        /// 創建目錄
        /// </summary>
        /// <param name="FolderPath">目錄名稱</param>
        /// <returns></returns>
        public bool MKDir(string FolderPath)
        {
            try
            {
                if (!Directory.Exists(FolderPath))
                    Directory.CreateDirectory(FolderPath);
                return true;
            }
            catch (Exception ex)
            {
                LogFactory.Write("目錄 " + FolderPath + " 建制失敗：" + ex.ToString(),

GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        #endregion
        #region 刪除檔案
        /// <summary>
        /// 刪除FTP檔案
        /// </summary>
        /// <param name="remFileName">檔案名稱</param>
        /// <returns></returns>
        public bool Delete(string remFileName)
        {
            return Delete(remotePath, remFileName);
        }
        /// <summary>
        /// 刪除FTP檔案
        /// </summary>
        /// <param name="remotePath"></param>
        /// <param name="remFileName"></param>
        /// <returns></returns>
        public bool Delete(string remotePath, string remFileName)
        {
            try
            {
                if (true)
                {
                    Login(remotePath, remFileName);
                }
                ftp.Method = WebRequestMethods.Ftp.DeleteFile;
                response = (FtpWebResponse)ftp.GetResponse();
                response.Close();
                return true;
            }
            catch (Exception ex)
            { 
                LogFactory.Write("刪除FTP檔案 失敗：" + ex.ToString(), GlobalString.LogType.ErrorCategory);
                return false;
            }
        }
        #endregion
    }

}
