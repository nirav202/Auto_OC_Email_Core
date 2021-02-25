using System;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using Utilities_Core;
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.Xml;
using System.Collections.Generic;
using System.Net.Mail;
using MsgReader;

namespace Auto_OC_Email_Core
{
    class Program
    {
        private static string strEmailMSGTemplate = "";
        //private static string RegEmailPat = @"<([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)>";
        private static string RegEmailPat = @"([ <(]?)([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)[> \r\n)]?";

        private static string strEmailMSGArchive = ConfigurationManager.AppSettings.Get("ArchiveEmailMSG");
        static void Main(string[] args)
        {
            string strLogFile = ConfigurationManager.AppSettings.Get("LogFile");
            if (!Directory.Exists(strLogFile))
            {
                Directory.CreateDirectory(strLogFile);
            }
            string strorderNo="";
            string strLogFileName = strLogFile + "\\" + "Log_" + DateTime.Today.ToShortDateString().Replace('/', '-') + ".txt";
            string strEmailUserName = ConfigurationManager.AppSettings.Get("EmailUserName");
            string strEmailUserPwd = ConfigurationManager.AppSettings.Get("EmailUserPwd");
            string strEmailFrom = ConfigurationManager.AppSettings.Get("EmailFrom");
            string strEmailMSGLocation = ConfigurationManager.AppSettings.Get("EmailMSGLocation");

            string strOCFileDollar = ConfigurationManager.AppSettings.Get("OCFileDollar");
            string strOCFileDim = ConfigurationManager.AppSettings.Get("OCFileDim");
            string strErrorEmail = ConfigurationManager.AppSettings.Get("ErrorEmail");
            string strNotificationEmail = ConfigurationManager.AppSettings.Get("NotificationEmail");
            strEmailMSGTemplate = ConfigurationManager.AppSettings.Get("EmailMSGTeplate");
            double doubleHoursToWait = Convert.ToDouble(ConfigurationManager.AppSettings.Get("OCNotificationHoursToWait"));

            SqlConnection sqlcon = new SqlConnection(ConfigurationManager.ConnectionStrings["dbcon"].ToString());
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adpt = new SqlDataAdapter();
            DataTable dtOrder = new DataTable();
            DataTable dtHWD = new DataTable();

            string OCDollarFile = "", OCDimFile = "",strMSGFileErr="";
            try
            {
                

                
                cmd.Connection = sqlcon;
                cmd.CommandType = CommandType.Text;
                adpt.SelectCommand = cmd;


                foreach (string strmsgfiles in Directory.GetFiles(strEmailMSGLocation, "*.msg"))
                {
                  
                    strMSGFileErr = strmsgfiles;
                    bool dimflag = true;
                    strorderNo = Path.GetFileName(strmsgfiles).Split('-', '_', ' ')[0];
                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Processing - " + strmsgfiles);

                    //Console.WriteLine(Path.GetFileName(strmsgfiles).Split('-', '_', ' ')[0]);

                    cmd.CommandText = "select OH.DeliveryDate,Isnull(OC.Email,'') 'Email',Isnull(OC.ShipToEmail,'') 'ShipToEmail',Isnull(OC.DeliverySlipToEmail,'') 'DeliverySlipToEmail' from data.OrderHeader OH inner join lookup.OrderCustomer OC on OC.InternalId = OH.CustomerName where OH.DocumentNumber ='" + strorderNo + "'";
                    dtOrder.Clear();
                    adpt.Fill(dtOrder);
                    if (dtOrder.Rows.Count > 0)
                    {
                        OCDollarFile = ""; OCDimFile = "";
                        foreach (string strocfiles in Directory.GetFiles(strOCFileDollar, "*.pdf"))
                        {
                            if (Path.GetFileName(strocfiles).Split('-', '_', ' ')[0].StartsWith(strorderNo))
                            {
                                //Console.WriteLine(Path.GetFileName(strocfiles));
                                OCDollarFile = strocfiles;
                                break;
                            }
                        }

                        //Check for if its Hardware only order.
                        cmd.CommandText = "SELECT * FROM data.OrderLine where DocumentNumber = '"+strorderNo+"' and ProductCode not in ('Hardware')";
                        dtHWD.Clear();
                        adpt.Fill(dtHWD);
                        if (dtHWD.Rows.Count > 0)
                        {
                            foreach (string strocfilesdim in Directory.GetFiles(strOCFileDim, "*.pdf"))
                            {
                                if (Path.GetFileName(strocfilesdim).Remove(0, 12).Split('-', '_', ' ')[0].StartsWith(strorderNo))
                                {
                                    //Console.WriteLine(Path.GetFileName(strocfilesdim));
                                    OCDimFile = strocfilesdim;
                                    break;
                                }
                            }
                        }
                        else
                        { dimflag = false; }
                        if (OCDollarFile.Length > 0)
                        {
                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " OC File with dollar information found - " + OCDollarFile);
                            if (dimflag==false || (dimflag==true && OCDimFile.Length > 0))
                            {

                                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " File with dimension information found - " + OCDimFile);
                                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Reading Email MSG file - " + strmsgfiles);
                                StreamReader sr = new StreamReader(strmsgfiles, Encoding.Default);
                                
                                string strFullString = sr.ReadToEnd(), strEmailTo, strEmailCC;
                                
                                sr.Close();
                                
                                //MsgReader.Mime.Message message = new MsgReader.Mime.Message(System.Text.Encoding.ASCII.GetBytes(strFullString));
                                //MsgReader.Reader reader = new Reader();
                                MsgReader.Outlook.Storage.Message EmailMsg = new MsgReader.Outlook.Storage.Message(strmsgfiles);
                                strEmailTo= Regex.Match(EmailMsg.GetEmailSender(false, false), RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");
                                strEmailCC = funGetCCEmails(EmailMsg.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false));//, RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");
                                strEmailCC = strEmailCC +(strEmailCC.Trim().Length>0 ? ",":"") + funGetCCEmails(EmailMsg.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false));
                                string strEmailSubject = EmailMsg.SubjectNormalized;

                                EmailMsg.Dispose();
                                //string strEmailSubject = "PGI " + strorderNo + " " + funGetContent("\r\nSubject: ", strFullString, "\r\n");

                               // strFullString = Regex.Replace(strFullString, @"\0", "");
                                //strEmailTo = funGetToCCEmails(strFullString.Remove(0, strFullString.IndexOf("\r\nFrom: ") + 2)).Replace("<", "").Replace(">", "").Replace("(","").Replace(")","");
                                //strEmailTo = Regex.Match(funGetContent("\r\nFrom: ", strFullString, "\r\n"), RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");

                                if (strEmailTo.Contains("efax"))
                                {
                                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " EFax order looking for email in database...");
                                    strEmailTo = "";
                                    if (dtOrder.Rows[0]["Email"].ToString().Length > 0)
                                    {
                                        strEmailTo = strEmailTo + dtOrder.Rows[0]["Email"].ToString();
                                    }
                                    if (dtOrder.Rows[0]["ShipToEmail"].ToString().Length > 0)
                                    {
                                        strEmailTo = strEmailTo + "," + dtOrder.Rows[0]["ShipToEmail"].ToString();
                                    }
                                    if (dtOrder.Rows[0]["DeliverySlipToEmail"].ToString().Length > 0)
                                    {
                                        strEmailTo = strEmailTo + "," + dtOrder.Rows[0]["DeliverySlipToEmail"].ToString();
                                    }
                                    //strEmailCC = funGetContent("\r\nTo: ", strFullString, "\r\n");
                                }
                                //else
                                //{
                                //    strEmailCC = funGetToCCEmails(strFullString.Remove(0, strFullString.IndexOf("\r\nTo: ") + 2));
                                //    if (strFullString.IndexOf("\r\nCC: ") > 0)
                                //    {
                                //        strEmailCC = strEmailCC +(strEmailCC.Trim().Length>0?",":"")+ funGetToCCEmails(strFullString.Remove(0, strFullString.IndexOf("\r\nCC: ") + 2));
                                //    }
                                //}
                                if (strEmailCC.Trim().Length > 0)
                                {
                                    strEmailCC = funRemoveBlaklistDom(strEmailCC);
                                    //strEmailCC = strEmailCC.Replace("sales@precisionglassindustries.com", "");
                                }

                                string strDeliveryDT = Convert.ToDateTime(dtOrder.Rows[0]["DeliveryDate"]).ToShortDateString();



                                if (strEmailTo.Length <= 0)
                                {
                                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Not able to extract To addresses.");


                                    string strEmailSub = "Not able to extract TO email addresses for OC email to customer. Order# " + strorderNo;
                                    string strEmailBody = "Hello, \r\n\r\nNot able to extract TO email addresses from customer email files for Order# " + strorderNo + " not found.";
                                    string strEmailAttachment = "";
                                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Not able to extract TO email addresses for OC email to customer. Sending notification Email... To : " + strNotificationEmail);
                                    clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strNotificationEmail, strEmailSub, strEmailBody, strEmailAttachment, "");

                                }
                                else
                                {
                                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Sending OC Email... To : " + strEmailTo +(strEmailCC.Trim().Length>0 ? " CC : "+strEmailCC : ""));
                                    
                                    try
                                    {
                                        string strEmailBody;
                                        StreamReader streamReader = new StreamReader(strEmailMSGTemplate + @"\EmailSalesTeam.htm");
                                        strEmailBody = streamReader.ReadToEnd();


                                        string[] strEmailAttachment = new string[2];
                                        strEmailAttachment[0] = OCDollarFile;
                                        strEmailAttachment[1] = OCDimFile;

                                        SmtpClient smtpClient = new SmtpClient();
                                        MailAddress fromAddress = new MailAddress(strEmailUserName);
                                        smtpClient.Host = "smtp.office365.com";
                                        smtpClient.UseDefaultCredentials = false;

                                        smtpClient.Credentials = new System.Net.NetworkCredential(strEmailUserName, strEmailUserPwd);
                                        smtpClient.Port = 25;  // 587;
                                        smtpClient.Timeout = (60 * 5 * 1000);
                                        smtpClient.EnableSsl = true;



                                        MailMessage mail = new MailMessage();

                                        Attachment att1 = new Attachment(strEmailMSGTemplate + @"\image001.jpg");
                                        att1.ContentDisposition.Inline = true;
                                        Attachment att2 = new Attachment(strEmailMSGTemplate + @"\image002.png");
                                        att2.ContentDisposition.Inline = true;
                                        Attachment att3 = new Attachment(strEmailMSGTemplate + @"\image003.png");
                                        att3.ContentDisposition.Inline = true;
                                        Attachment att4 = new Attachment(strEmailMSGTemplate + @"\image004.png");
                                        att4.ContentDisposition.Inline = true;
                                        Attachment att5 = new Attachment(strEmailMSGTemplate + @"\image005.png");
                                        att5.ContentDisposition.Inline = true;

                                        mail.From = new MailAddress(strEmailFrom);

                                        mail.To.Add(strEmailTo);

                                        if (strEmailCC.Trim().Length > 0)
                                            mail.CC.Add(strEmailCC);

                                        //Testing 
                                        //mail.To.Add("npatel@precisionglassindustries.com");
                                        //if (strEmailCC.Length > 0)
                                        //    mail.CC.Add(",npatel@precisionglassindustries.com");
                                        //Testing End

                                        mail.Subject = strEmailSubject;
                                        mail.Body = String.Format(strEmailBody, strDeliveryDT, att1.ContentId, att2.ContentId, att3.ContentId, att4.ContentId, att5.ContentId);    //, strEmailTo, strEmailCC);

                                        mail.IsBodyHtml = true;
                                        mail.Attachments.Add(att1);
                                        mail.Attachments.Add(att2);
                                        mail.Attachments.Add(att3);
                                        mail.Attachments.Add(att4);
                                        mail.Attachments.Add(att5);

                                        foreach (string attachmentfile in strEmailAttachment)
                                        {
                                            if (attachmentfile.Length > 0)
                                                mail.Attachments.Add(new Attachment(attachmentfile));
                                        }

                                        smtpClient.Send(mail);
                                        mail.Dispose();

                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Email send successfully.");
                                        string timestamp = "-" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                                        //Archive OC ....
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Archiving email MSG file and OC with dollar PDF file.");
                                        strEmailMSGArchive = funCreateFileStructure();
                                        string strOCDollarArcive, strMSGArcive;
                                        strOCDollarArcive = Path.GetFileName(OCDollarFile).Replace(".pdf", timestamp + ".pdf");
                                        Directory.Move(OCDollarFile, strEmailMSGArchive + "\\" + strOCDollarArcive);
                                        //Archive Email MSG file ....
                                        strMSGArcive = Path.GetFileName(strmsgfiles).Replace(".msg", timestamp + ".msg");
                                        Directory.Move(strmsgfiles, strEmailMSGArchive + "\\" + strMSGArcive);


                                    }
                                    catch (Exception ex)
                                    {
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Email send Error." + ex.Message);
                                        string strEmailSub = "Error sending email Order# " + strorderNo;
                                        string strEmailBody = "Hello, \r\n\r\nError occured while sending email files for Order# " + strorderNo + ".";
                                        string strEmailAttachment = "";
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Error occured while sending email files. Sending notification Email... To : " + strErrorEmail);
                                        clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strErrorEmail, strEmailSub, strEmailBody, strEmailAttachment, "");

                                        string timestamp = "-" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                                        //Moving MSG file to Error folder ....
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Moving email MSG file and OC with dollar PDF file to Error folder.");
                                        string strEmailMSGErr = ConfigurationManager.AppSettings.Get("ErrorMSG");
                                        string strOCDollarArcive, strMSGArcive;
                                        strOCDollarArcive = Path.GetFileName(OCDollarFile).Replace(".pdf", timestamp + ".pdf");
                                        Directory.Move(OCDollarFile, strEmailMSGErr + "\\" + strOCDollarArcive);
                                        //Archive Email MSG file ....
                                        strMSGArcive = Path.GetFileName(strMSGFileErr).Replace(".msg", timestamp + ".msg");
                                        Directory.Move(strMSGFileErr, strEmailMSGErr + "\\" + strMSGArcive);
                                    }

                                }
                            }
                            else
                            {

                                string strEmailTo = strNotificationEmail;
                                string strEmailCC = "";
                                string strEmailSubject = "Missing file with dimensions for OC email to customer. Order# " + strorderNo;
                                string strEmailBody = "Hello, \r\n\r\nDimension files for Order# " + strorderNo + " not found.";
                                string strEmailAttachment = "";
                                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Dimension file missing. Sending notification Email... To : " + strEmailTo);
                                clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strEmailTo, strEmailSubject, strEmailBody, strEmailAttachment, strEmailCC);
                            }
                        }
                        else
                        {
                            FileInfo fileInfo = new FileInfo(strmsgfiles);
                            //Console.WriteLine(fileInfo.CreationTime.ToString());
                            //Console.WriteLine(fileInfo.CreationTime.AddHours(doubleHoursToWait).ToString());
                            if (fileInfo.CreationTime.AddHours(doubleHoursToWait) < DateTime.Now)
                            {
                                string strEmailTo = strNotificationEmail;
                                string strEmailCC = "";
                                string strEmailSubject = "Missing file for OC email to customer. Order# " + strorderNo;
                                string strEmailBody = "Hello, \r\n\rOrder confirmation files for Order# " + strorderNo + " not found.";
                                string strEmailAttachment = "";
                                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " OC with dollar amount file missing. Sending notification Email... To : " + strEmailTo);
                                clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strEmailTo, strEmailSubject, strEmailBody, strEmailAttachment, strEmailCC);
                            }
                        }
                    }
                    else
                    {
                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Order not found in Orders database : " + strorderNo);

                        //send notification if set time is elasped and order does not found in database...
                        FileInfo fileInfo = new FileInfo(strmsgfiles);
                        if (fileInfo.CreationTime.AddHours(doubleHoursToWait) < DateTime.Now)
                        {
                            string strEmailTo = strNotificationEmail;
                            string strEmailCC = "";
                            string strEmailSubject = "Order not found in Orders database. Order# " + strorderNo;
                            string strEmailBody = "Hello, \r\n\rOrder not found in Orders database. Order# " + strorderNo + " not found.";
                            string strEmailAttachment = "";
                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Order not found in Orders database. Sending notification Email... To : " + strEmailTo);
                            clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strEmailTo, strEmailSubject, strEmailBody, strEmailAttachment, strEmailCC);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": Exception occured : " + ex.Message);
                string strEmailSub = "Error processing. Order# " + strorderNo;
                string strEmailBody = "Hello, \r\n\r\nError occured while processing email files for Order# " + strorderNo + ".";
                string strEmailAttachment = "";
                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Error occured while processing email files. Sending notification Email... To : " + strErrorEmail);
                clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strErrorEmail, strEmailSub, strEmailBody, strEmailAttachment, "");

                string timestamp = "-" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                //Moving MSG file to Error folder ....
                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Moving email MSG file and OC with dollar PDF file to Error folder.");
                string strEmailMSGErr = ConfigurationManager.AppSettings.Get("ErrorMSG");
                string strOCDollarArcive, strMSGArcive;
                strOCDollarArcive = Path.GetFileName(OCDollarFile).Replace(".pdf", timestamp + ".pdf");
                Directory.Move(OCDollarFile, strEmailMSGErr + "\\" + strOCDollarArcive);
                //Archive Email MSG file ....
                strMSGArcive = Path.GetFileName(strMSGFileErr).Replace(".msg", timestamp + ".msg");
                Directory.Move(strMSGFileErr, strEmailMSGErr + "\\" + strMSGArcive);
            }
        }

        private static string funCreateFileStructure()
        {
            string strEmailMSGArc = ConfigurationManager.AppSettings.Get("ArchiveEmailMSG");
            #region Create Archive File Structure...
            if (!Directory.Exists(strEmailMSGArc))
            {
                Directory.CreateDirectory(strEmailMSGArc);
            }
            strEmailMSGArc = strEmailMSGArc + "\\" + DateTime.Today.Year.ToString();
            if (!Directory.Exists(strEmailMSGArc))
            {
                Directory.CreateDirectory(strEmailMSGArc);
            }
            strEmailMSGArc = strEmailMSGArc + "\\" + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00");
            if (!Directory.Exists(strEmailMSGArc))
            {
                Directory.CreateDirectory(strEmailMSGArc);
            }
            strEmailMSGArc = strEmailMSGArc + "\\" + DateTime.Today.Year.ToString() + DateTime.Today.Month.ToString("00") + DateTime.Today.Day.ToString("00");
            if (!Directory.Exists(strEmailMSGArc))
            {
                Directory.CreateDirectory(strEmailMSGArc);
            }
            #endregion
            return strEmailMSGArc;
        }

        private static string funGetCCEmails(string strEmailAddress)
        {
            string strEmails = "";
            foreach(string strTemp in strEmailAddress.Split(";"))
            {
                if (!Regex.Match(strTemp, RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "").Trim().Equals("sales@precisionglassindustries.com"))
                {
                    strEmails = strEmails + (strEmails.Trim().Length > 0 ? "," : "") + Regex.Match(strTemp, RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", ""); 
                }
            }
            return strEmails;
        }
        private static string funGetToCCEmails(string strFullString)
        {
            string strEmails = "";

            foreach (string strTemp in strFullString.Split("\r\n"))
            {
                if (Regex.IsMatch(strTemp, RegEmailPat))
                {

                    if (!Regex.Match(strTemp, RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "").Trim().Equals("sales@precisionglassindustries.com"))
                    {
                        strEmails = strEmails + (strEmails.Trim().Length > 0 ? "," : "") + Regex.Match(strTemp, RegEmailPat).Value;
                    }
                }
                else
                {
                    break;
                }
            }
            return strEmails.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");
        }
        private static string funRemoveBlaklistDom(string strEmails)
        {

            List<string> strBlackList = funGetBlackListDom();

            string FinalEmailTo = "";
            char[] charSplit = new char[] { ',', ';' };

            foreach (string strtemp in strEmails.Split(charSplit))
            {
                string strEmailDom = strtemp.Remove(0,strtemp.IndexOf('@')+1);
                if (strEmailDom.LastIndexOf('.') > 0)
                {
                    strEmailDom = strEmailDom.Substring(0, strEmailDom.LastIndexOf('.'));

                    bool IsBlack = false;
                    foreach (string strtempBlack in strBlackList)
                    {
                        if (strEmailDom.Equals(strtempBlack))
                            IsBlack = true;
                    }
                    if (IsBlack == false)
                    {
                        FinalEmailTo = FinalEmailTo + (FinalEmailTo.Trim().Length > 0 ? "," : "") + strtemp;
                    }
                }
            }

            return FinalEmailTo;
        }

        private static List<string> funGetBlackListDom()
        {
            List<string> strBlackList = new List<string>();

            // Start with XmlReader object  
            //here, we try to setup Stream between the XML file using xmlReader  
            using (XmlReader reader = XmlReader.Create(strEmailMSGTemplate+@"\BlackList.xml"))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement())
                    {
                        //return only when you have START tag  
                        switch (reader.Name.ToString())
                        {
                            case "name":
                                strBlackList.Add(reader.ReadString());
                                break;
                        }
                    }
                }
            }

            return strBlackList;
        }

       private static string funGetContent(string strStartKey, string strSearchString,String strEndKey)
        {
            int iStart;
            int iLast;

            if (strSearchString.IndexOf(strStartKey) > -1)
            {
                iStart = strSearchString.IndexOf(strStartKey);
                strSearchString = strSearchString.Remove(0, iStart + strStartKey.Length);

                iLast = strSearchString.IndexOf(strEndKey)<0 ? 0 : strSearchString.IndexOf(strEndKey);
                return strSearchString.Substring(0, iLast);
            }

            return "";
        }
    }
}





