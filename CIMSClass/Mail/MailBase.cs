//******************************************************************
//*  作    者：Ray
//*  功能說明：郵件發送，適用于SMTP協議郵件伺服器
//*  創建日期：2008-11-14
//*  修改日期：
//*  修改記錄：
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Net.Mail;
using System.Net;

namespace CIMSClass.Mail
{
    internal class MailBase
    {
        private static string MailServer = ConfigurationManager.AppSettings["MailServer"]; //郵件伺服器地址
        private static string ManagerMail = ConfigurationManager.AppSettings["ManagerMail"];//接受者郵箱
        private static string MailFrom = ConfigurationManager.AppSettings["MailFrom"];//寄件者郵箱
        private static string MailUser = ConfigurationManager.AppSettings["MailUser"];//寄件者郵箱
        private static string MailPassword = ConfigurationManager.AppSettings["MailPassword"];//寄件者郵箱密碼
        private static string Subject = ConfigurationManager.AppSettings["MailSubject"];//郵件主旨
        private static string Body = ConfigurationManager.AppSettings["MailBody"];//郵件正文

        private static SmtpClient smtp; //SMTP對象
        private static MailMessage mail;//Mail主體對象
        
        public MailBase()
        {
                      
        }
        /// <summary>
        /// 郵件發送，不含不含附件
        /// </summary>
        /// <returns></returns>
        public static bool SendMail()
        {

            return SendMail(ManagerMail, Subject, Body);
        }
        ///// <summary>
        ///// 郵件發送，含附件，多個附件路徑用逗號分隔
        ///// </summary>
        ///// <param name="LogFile">附件地址字符串</param>
        ///// <returns></returns>
        //public static bool SendMail(string LogFile)
        //{
        //    try
        //    {
        //        BuildMail();
        //        AddAttach(LogFile);
        //        smtp.Send(mail);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteErrLog(ex.ToString());
        //        return false;
        //    }
        //}

        /// <summary>
        /// WEB中的郵件發送方法
        /// </summary>
        /// <param name="MailTo"></param>
        /// <param name="MailSubject"></param>
        /// <param name="MailBody"></param>
        public static bool SendMail(string MailTo, string MailSubject, string MailBody)
        {
            string strTestType = ConfigurationManager.AppSettings["TestType"].ToString();
            if (strTestType == "2")
            {

                string MailFrom = ConfigurationManager.AppSettings["MailServerFrom"];
                string MailServer = ConfigurationManager.AppSettings["MailServer"];
                string MailServerID = ConfigurationManager.AppSettings["MailServerID"];
                string MailServerPWD = ConfigurationManager.AppSettings["MailServerPWD"];
                try
                {
                    SmtpClient client = new SmtpClient();
                    client.Host = MailServer;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(MailServerID, MailServerPWD);

                    client.DeliveryMethod = SmtpDeliveryMethod.Network;

                    System.Net.Mail.MailMessage message = new MailMessage(MailFrom, MailTo);
                    message.Subject = MailSubject;
                    message.Body = MailBody;
                    message.BodyEncoding = System.Text.Encoding.GetEncoding("GB2312");
                    message.IsBodyHtml = true;



                    client.Send(message);
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            else
            {

                try
                {

                    MailMessage mail = new MailMessage();

                    string MailFrom = ConfigurationManager.AppSettings["MailServerFrom"];
                    string MailFromName = ConfigurationManager.AppSettings["MailFromName"];
                    string MailServer = ConfigurationManager.AppSettings["MailServer"];
                    string MailServerID = ConfigurationManager.AppSettings["MailServerID"];
                    string MailServerPWD = ConfigurationManager.AppSettings["MailServerPWD"];
                    string MailDomainName = ConfigurationManager.AppSettings["MailDomainName"];

                    //郵件發送者
                    mail.From = new MailAddress(MailFrom, MailFromName, Encoding.Default);

                    //郵件接受者
                    mail.To.Add(new MailAddress(MailTo));

                    //郵件標題
                    mail.Subject = MailSubject;
                    mail.SubjectEncoding = Encoding.Default;

                    //郵件主體��
                    mail.IsBodyHtml = true;
                    mail.Body = MailBody;
                    mail.BodyEncoding = Encoding.Default;


                    //SMTP
                    SmtpClient smtp = new SmtpClient(MailServer, 25);
                    smtp.UseDefaultCredentials = true;



                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(MailServerID, MailServerPWD, MailDomainName);

                    //發送
                    smtp.Send(mail);
                    mail.Dispose();
                    return true;
                }
                catch
                {
                    return false;
                }
            }
        }


        //public static bool SendMail(string MailAddress, string Subject, string Body)
        //{
        //    try
        //    {
        //        BuildMail(MailAddress, Subject, Body);
        //        smtp.Send(mail);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        WriteErrLog(ex.ToString());
        //        return false;
        //    }
        //}
        /// <summary>
        /// 創建郵件主體
        /// </summary>
        private static void BuildMail()
        {
            BuildMail(MailServer, MailFrom, ManagerMail, Subject, Body, MailUser, MailPassword);
        }
        private static void BuildMail(string MailAddress, string Subject, string Body)
        {
            BuildMail(MailServer, MailFrom, MailAddress, Subject, Body, MailUser, MailPassword);
        }
        /// <summary>
        /// 創建郵件主體
        /// </summary>
        /// <param name="MailServer">伺服器地址</param>
        /// <param name="MailFrom">發件人</param>
        /// <param name="ManagerMail">收件人，多個地址用逗號分隔</param>
        /// <param name="Subject">主旨</param>
        /// <param name="Body">正文內容</param>
        /// <param name="MailUser">郵箱賬號</param>
        /// <param name="MailPassword">郵箱密碼</param>
        private static void BuildMail(string MailServer,string MailFrom,string ManagerMail,string Subject,string Body,string MailUser,string MailPassword)
        {
            smtp = new SmtpClient(MailServer);
            mail = new MailMessage(MailFrom, ManagerMail);
            mail.Subject = Subject;
            mail.Body = Body;
            mail.Priority = MailPriority.High;
            smtp.Credentials = new NetworkCredential(MailUser, MailPassword);
        }
        private static void AddAttach(string LogFiles)
        {
            foreach (string LogFile in LogFiles.Split(','))
            {
                Attachment attatch = new Attachment(LogFile);
                mail.Attachments.Add(attatch);
            }
        }
        /// <summary>
        /// 異常郵件發送
        /// </summary>
        /// <param name="ErrDes">系統異常信息</param>
        private static void WriteErrLog(string ErrDes)
        {
            StringBuilder ErrStr = new StringBuilder("郵件發送錯誤：");
            ErrStr.Append(ErrDes);
            ErrStr.Append(";");     
            LogFactory.Write(ErrStr.ToString(), GlobalString.LogType.ErrorCategory);
        }
    }
}
