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
using System.Threading;

namespace Auto_OC_Email_Core
{
    class Program
    {
        private static string strEmailMSGTemplate = "";
        //private static string RegEmailPat = @"<([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)>";
        //private static string RegEmailPat = @"([ <(]?)([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)[> \r\n)]?";
        private static string RegEmailPat = @"(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|""(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])";

        private static string strEmailMSGArchive = ConfigurationManager.AppSettings.Get("ArchiveEmailMSG");
        static void Main(string[] args)
        {

            string strLogFile = ConfigurationManager.AppSettings.Get("LogFile");
            if (!Directory.Exists(strLogFile))
            {
                Directory.CreateDirectory(strLogFile);
            }
            string strorderNo = "", strPGIOrderNo = "";
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
            double doublePDFPurgeHoursToWait = Convert.ToDouble(ConfigurationManager.AppSettings.Get("PDFPurgeHoursToWait"));

            SqlConnection sqlcon = new SqlConnection(ConfigurationManager.ConnectionStrings["dbcon"].ToString());
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter adpt = new SqlDataAdapter();
            DataTable dtOrder = new DataTable();
            DataTable dtHWD = new DataTable();

            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": Started Auto Email OC application...");

            string OCDollarFile = "", OCDimFile = "",strMSGFileErr="";
            
            
                cmd.Connection = sqlcon;
                cmd.CommandType = CommandType.Text;
                adpt.SelectCommand = cmd;

            foreach (string strmsgfiles in Directory.GetFiles(strEmailMSGLocation))
            {
                try
                {
                    //Process files that are email message files that ends with .msg or .eml.....
                    if (strmsgfiles.EndsWith(".msg") || strmsgfiles.EndsWith(".eml"))
                    {

                        strMSGFileErr = strmsgfiles;
                        bool dimflag = true;
                        strorderNo = Path.GetFileName(strmsgfiles).Split('-', '_', ' ')[0];
                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Processing - " + strmsgfiles);

                        //Console.WriteLine(Path.GetFileName(strmsgfiles).Split('-', '_', ' ')[0]);

                        cmd.CommandText = "select OH.DocumentNumber,OH.DeliveryDate,Isnull(OC.Email,'') 'Email',case Isnull(OC.ShipToEmail,'') when 'null' then '' else Isnull(OC.ShipToEmail,'') end  'ShipToEmail',case Isnull(OC.DeliverySlipToEmail,'') when 'null' then '' else Isnull(OC.DeliverySlipToEmail,'') end 'DeliverySlipToEmail',case ISNULL(OC.OrderConfirmationEmail,'') when 'null' then '' else ISNULL(OC.OrderConfirmationEmail,'') end 'OrderConfirmationEmail' from data.OrderHeader OH inner join lookup.OrderCustomer OC on OC.InternalId = OH.CustomerName where OH.DocumentNumber Like '%" + strorderNo + "%'";
                        dtOrder.Clear();
                        adpt.Fill(dtOrder);
                        if (dtOrder.Rows.Count > 0)
                        {
                            strPGIOrderNo = dtOrder.Rows[0]["DocumentNumber"].ToString();

                            //get the additional email list for OrderConfirmation if they exist in database...
                            string additinalEmailList = dtOrder.Rows[0]["OrderConfirmationEmail"].ToString();

                            OCDollarFile = ""; OCDimFile = "";
                            foreach (string strocfiles in Directory.GetFiles(strOCFileDollar, "*.pdf"))
                            {
                                //see if OC PDF file exist or not....
                                //if (Path.GetFileName(strocfiles).Split('-', '_', ' ')[0].StartsWith(strorderNo))
                                if (Path.GetFileName(strocfiles).Contains(strorderNo))
                                {
                                    //Console.WriteLine(Path.GetFileName(strocfiles));
                                    OCDollarFile = strocfiles;
                                    break;
                                }
                            }

                            //Check for if its Hardware only order.
                            cmd.CommandText = "SELECT * FROM data.OrderLine where DocumentNumber Like '%" + strorderNo + "%' and ProductCode not in ('Hardware')";
                            dtHWD.Clear();
                            adpt.Fill(dtHWD);
                            if (dtHWD.Rows.Count > 0)
                            {
                                foreach (string strocfilesdim in Directory.GetFiles(strOCFileDim, "*.pdf"))
                                {
                                    //if (Path.GetFileName(strocfilesdim).Remove(0, 12).Split('-', '_', ' ')[0].StartsWith(strorderNo))
                                    if (Path.GetFileName(strocfilesdim).Contains("Glass Order " + strorderNo))
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
                                if (dimflag == false || (dimflag == true && OCDimFile.Length > 0))
                                {
                                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " File with dimension information found - " + OCDimFile);
                                    string strEmailTo = "", strEmailCC = "";
                                    string strEmailSubject = "";
                                    //For .msg Files ----Outlook messages files...
                                    if (strmsgfiles.Substring(strmsgfiles.Length - 3, 3) == "msg")
                                    {
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Reading Email MSG file - " + strmsgfiles);
                                        MsgReader.Outlook.Storage.Message EmailMsg = new MsgReader.Outlook.Storage.Message(strmsgfiles);
                                        strEmailTo = Regex.Match(EmailMsg.GetEmailSender(false, false), RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");
                                        strEmailCC = funGetCCEmails(EmailMsg.GetEmailRecipients(MsgReader.Outlook.RecipientType.To, false, false));//, RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");
                                        strEmailCC = strEmailCC + (strEmailCC.Trim().Length > 0 ? "," : "") + funGetCCEmails(EmailMsg.GetEmailRecipients(MsgReader.Outlook.RecipientType.Cc, false, false));
                                        //strEmailSubject = "PGINo " + strorderNo + " " + EmailMsg.SubjectNormalized;
                                        strEmailSubject = EmailMsg.SubjectNormalized;
                                        EmailMsg.Dispose();
                                    }
                                    else // For .eml files......
                                    {
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Reading Email EML file - " + strmsgfiles);
                                        FileInfo fileEml = new FileInfo(strmsgfiles);
                                        MsgReader.Mime.Message EmailEml = MsgReader.Mime.Message.Load(fileEml);

                                        strEmailTo = Regex.Match(EmailEml.Headers.From.MailAddress.Address, RegEmailPat).Value.Replace("<", "").Replace(">", "").Replace("(", "").Replace(")", "");
                                        foreach (var recipient in EmailEml.Headers.To)
                                        {
                                            strEmailCC = strEmailCC + (strEmailCC.Trim().Length > 0 ? "," : "") + funGetCCEmails(recipient.MailAddress.Address);
                                        }
                                        foreach (var recipient in EmailEml.Headers.Cc)
                                        {
                                            strEmailCC = strEmailCC + (strEmailCC.Trim().Length > 0 ? "," : "") + funGetCCEmails(recipient.MailAddress.Address);
                                        }
                                        //strEmailSubject = "PGINo " + strorderNo + " " + EmailEml.Headers.Subject;
                                        strEmailSubject = EmailEml.Headers.Subject;

                                    }

                                    //Add additional emails to strEmailTo for some customer if it is in our database.
                                    strEmailTo = strEmailTo + (additinalEmailList.Trim().Length > 0 ? "," + additinalEmailList : "");

                                    //StreamReader sr = new StreamReader(strmsgfiles, Encoding.Default);
                                    //string strFullString = sr.ReadToEnd();
                                    //sr.Close();
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
                                            strEmailTo = strEmailTo + dtOrder.Rows[0]["ShipToEmail"].ToString().Trim() == "null" ? "" : "," + dtOrder.Rows[0]["ShipToEmail"].ToString();
                                        }
                                        if (dtOrder.Rows[0]["DeliverySlipToEmail"].ToString().Length > 0)
                                        {
                                            strEmailTo = strEmailTo + dtOrder.Rows[0]["DeliverySlipToEmail"].ToString().Trim() == "null" ? "" : "," + dtOrder.Rows[0]["DeliverySlipToEmail"].ToString();
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
                                    //Sending Email.....
                                    else
                                    {
                                        clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Sending OC Email... To : " + strEmailTo + (strEmailCC.Trim().Length > 0 ? " CC : " + strEmailCC : ""));

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
                                            smtpClient.Port = 587; //25
                                            smtpClient.Timeout = (60 * 5 * 1000);
                                            smtpClient.EnableSsl = true;



                                            MailMessage mail = new MailMessage();

                                            Attachment att1 = new Attachment(strEmailMSGTemplate + @"\image001.jpg");
                                            att1.ContentDisposition.Inline = true;
                                            Attachment att2 = new Attachment(strEmailMSGTemplate + @"\img_phone.png");
                                            att2.ContentDisposition.Inline = true;
                                            Attachment att3 = new Attachment(strEmailMSGTemplate + @"\img_mail.png");
                                            att3.ContentDisposition.Inline = true;
                                            Attachment att4 = new Attachment(strEmailMSGTemplate + @"\img_maps.png");
                                            att4.ContentDisposition.Inline = true;
                                            Attachment att5 = new Attachment(strEmailMSGTemplate + @"\img_facebook.png");
                                            att5.ContentDisposition.Inline = true;
                                            Attachment att6 = new Attachment(strEmailMSGTemplate + @"\img_linkedin.png");
                                            att6.ContentDisposition.Inline = true;
                                            Attachment att7 = new Attachment(strEmailMSGTemplate + @"\img_instagram.png");
                                            att7.ContentDisposition.Inline = true;
                                            Attachment att8 = new Attachment(strEmailMSGTemplate + @"\img_showers.png");
                                            att8.ContentDisposition.Inline = true;

                                            mail.From = new MailAddress(strEmailFrom);

                                            mail.To.Add(strEmailTo);

                                            if (strEmailCC.Trim().Length > 0)
                                                mail.CC.Add(strEmailCC);

                                            //Testing
                                            //mail.To.Add("npatel@precisionglassindustries.com");
                                            //if (strEmailCC.Length > 0)
                                            //    mail.CC.Add("npatel@precisionglassindustries.com");
                                            //Testing End
                                            mail.Subject = strEmailSubject;
                                            mail.Body = String.Format(strEmailBody, strDeliveryDT, strPGIOrderNo, att1.ContentId, att2.ContentId, att3.ContentId, att4.ContentId, att5.ContentId, att6.ContentId, att7.ContentId, att8.ContentId);    //, strEmailTo, strEmailCC);

                                            mail.IsBodyHtml = true;
                                            mail.Attachments.Add(att1);
                                            mail.Attachments.Add(att2);
                                            mail.Attachments.Add(att3);
                                            mail.Attachments.Add(att4);
                                            mail.Attachments.Add(att5);
                                            mail.Attachments.Add(att6);
                                            mail.Attachments.Add(att7);
                                            mail.Attachments.Add(att8);

                                            foreach (string attachmentfile in strEmailAttachment)
                                            {
                                                if (attachmentfile.Length > 0)
                                                    mail.Attachments.Add(new Attachment(attachmentfile));
                                            }

                                            smtpClient.Send(mail);
                                            mail.Dispose();
                                            Thread.Sleep(6000);
                                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Email send successfully.");
                                            //Updating Order Difference Table  to set EmailConfirmationToCustomerSent field......

                                            SqlConnection sqlcon1 = new SqlConnection(ConfigurationManager.ConnectionStrings["dbcon"].ToString());
                                            SqlCommand cmd1 = new SqlCommand();
                                            cmd1.Connection = sqlcon1;
                                            cmd1.CommandType = CommandType.Text;
                                            cmd1.CommandText = "update data.OrderDiff set EmailConfirmationToCustomerSent = 1 where DocumentNumber like '" + strorderNo + "%'";
                                            sqlcon1.Open();
                                            cmd1.ExecuteNonQuery();
                                            sqlcon1.Close();
                                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " data.OrderDiff Table has been updated to set EmailConfirmationToCustomerSent flag.");

                                            string timestamp = "-" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                                            //Archive OC ....
                                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Archiving email file and OC with dollar PDF file.");
                                            strEmailMSGArchive = funCreateFileStructure();
                                            string strOCDollarArcive, strMSGArcive;
                                            strOCDollarArcive = Path.GetFileName(OCDollarFile).Replace(".pdf", timestamp + ".pdf");
                                            Directory.Move(OCDollarFile, strEmailMSGArchive + "\\" + strOCDollarArcive);
                                            //Archive Email file ....
                                            strMSGArcive = Path.GetFileName(strmsgfiles).Replace(".msg", timestamp + ".msg").Replace(".eml", timestamp + ".eml");
                                            Directory.Move(strmsgfiles, strEmailMSGArchive + "\\" + strMSGArcive);


                                        }
                                        catch (Exception ex)
                                        {
                                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Email send Error." + ex.Message);
                                            string strEmailSub = "Application:- Auto_OC_Email_Core Error sending email Order# " + strorderNo;
                                            string strEmailBody = "Hello, \r\n\r\nError occured while sending email files for Order# " + strorderNo + ". \r\nException occured : " + ex.Message;
                                            string strEmailAttachment = "";
                                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Error occured while sending email files. Sending notification Email... To : " + strErrorEmail);
                                            clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strErrorEmail, strEmailSub, strEmailBody, strEmailAttachment, "");

                                            string timestamp = "-" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                                            //Moving MSG file to Error folder ....
                                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Moving email file and OC with dollar PDF file to Error folder.");
                                            string strEmailMSGErr = ConfigurationManager.AppSettings.Get("ErrorMSG");
                                            string strOCDollarArcive, strMSGArcive;
                                            strOCDollarArcive = Path.GetFileName(OCDollarFile).Replace(".pdf", timestamp + ".pdf");
                                            Directory.Move(OCDollarFile, strEmailMSGErr + "\\" + strOCDollarArcive);
                                            //Archive Email MSG file ....
                                            strMSGArcive = Path.GetFileName(strMSGFileErr).Replace(".msg", timestamp + ".msg").Replace(".eml", timestamp + ".eml");
                                            Directory.Move(strMSGFileErr, strEmailMSGErr + "\\" + strMSGArcive);
                                        }

                                    }
                                }
                                else
                                {

                                    //    string strEmailTo = strNotificationEmail;
                                    //    string strEmailCC = "";
                                    //    string strEmailSubject = "Missing file with dimensions for OC email to customer. Order# " + strorderNo;
                                    //    string strEmailBody = "Hello, \r\n\r\nDimension files for Order# " + strorderNo + " not found.";
                                    //    string strEmailAttachment = "";
                                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Dimension file missing.");
                                    //    clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strEmailTo, strEmailSubject, strEmailBody, strEmailAttachment, strEmailCC);
                                }
                            }
                            else
                            {
                                //FileInfo fileInfo = new FileInfo(strmsgfiles);
                                ////Console.WriteLine(fileInfo.CreationTime.ToString());
                                ////Console.WriteLine(fileInfo.CreationTime.AddHours(doubleHoursToWait).ToString());
                                //if (fileInfo.CreationTime.AddHours(doubleHoursToWait) < DateTime.Now)
                                //{
                                //    string strEmailTo = strNotificationEmail;
                                //    string strEmailCC = "";
                                //    string strEmailSubject = "Missing file for OC email to customer. Order# " + strorderNo;
                                //    string strEmailBody = "Hello, \r\n\rOrder confirmation files for Order# " + strorderNo + " not found.";
                                //    string strEmailAttachment = "";
                                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " OC with dollar amount file missing.");
                                //    clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strEmailTo, strEmailSubject, strEmailBody, strEmailAttachment, strEmailCC);
                                //}
                            }
                        }
                        else
                        {
                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Order not found in Orders database : " + strorderNo);

                            //send notification if set time is elasped and order does not found in database...
                            FileInfo fileInfo = new FileInfo(strmsgfiles);
                            if (fileInfo.CreationTime.AddHours(doubleHoursToWait) < DateTime.Now)
                            {
                                //string strEmailTo = strNotificationEmail;
                                //string strEmailCC = "";
                                //string strEmailSubject = "Order not found in Orders database. Order# " + strorderNo;
                                //string strEmailBody = "Hello, \r\n\rOrder not found in Orders database. Order# " + strorderNo + " not found.";
                                //string strEmailAttachment = "";
                                clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Order not found in Orders database.");
                                //clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strEmailTo, strEmailSubject, strEmailBody, strEmailAttachment, strEmailCC);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": Exception occured : " + ex.Message);
                    //////Removed sending error message emails as most of the time error message is for  emails files are being used by another process.
                    ///Ammar's process is in middle of copying files as a result this process gives exception that file is being used by another process.
                    //////string strEmailSub = "Application:- Auto_OC_Email_Core Error processing. Order# " + strorderNo;
                    //////string strEmailBody = "Hello, \r\n\r\nError occured while processing email files for Order# " + strorderNo + ". \r\nException occured : " + ex.Message;
                    //////string strEmailAttachment = "";
                    //////clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Error occured while processing email files. Sending notification Email... To : " + strErrorEmail);
                    //////clsEmail.SendEmail(strEmailUserName, strEmailUserPwd, strEmailFrom, strErrorEmail, strEmailSub, strEmailBody, strEmailAttachment, "");
                    /////////////////////////////////////////////////////////////////////
                    


                    //string timestamp = "-" + DateTime.Now.Hour.ToString() + "_" + DateTime.Now.Minute.ToString() + "_" + DateTime.Now.Second.ToString();
                    //Moving MSG file to Error folder ....
                    //clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Moving email file and OC with dollar PDF file to Error folder.");
                    //string strEmailMSGErr = ConfigurationManager.AppSettings.Get("ErrorMSG");
                    //string strOCDollarArcive, strMSGArcive;
                    //strOCDollarArcive = Path.GetFileName(OCDollarFile).Replace(".pdf", timestamp + ".pdf");
                    //Directory.Move(OCDollarFile, strEmailMSGErr + "\\" + strOCDollarArcive);
                    ////Archive Email MSG file ....
                    //strMSGArcive = Path.GetFileName(strMSGFileErr).Replace(".msg", timestamp + ".msg").Replace(".eml", timestamp + ".eml");
                    //Directory.Move(strMSGFileErr, strEmailMSGErr + "\\" + strMSGArcive);
                }
            }

                Thread.Sleep(3000);
                //Process PDF for purge...
                foreach (string strPDFfiles in Directory.GetFiles(strEmailMSGLocation, "*.pdf"))
                    {
                        //send notification if set time is elasped and order does not found in database...
                        FileInfo fileInfo = new FileInfo(strPDFfiles);
                        if ((fileInfo.CreationTime.DayOfWeek == DayOfWeek.Friday ? fileInfo.CreationTime.AddHours(doublePDFPurgeHoursToWait + 48) : fileInfo.CreationTime.AddHours(doublePDFPurgeHoursToWait)) < DateTime.Now)
                        {
                            strorderNo = Path.GetFileName(strPDFfiles).Split('-', '_', ' ')[0];
                            clsWriteLog.funWriteLog(strLogFileName, DateTime.Now.ToString() + ": " + strorderNo + " Deleted " + strPDFfiles + ". No .msg or .eml file found for " + doublePDFPurgeHoursToWait.ToString() + " hours.");
                            File.Delete(strPDFfiles);
                            //Directory.Delete(strPDFfiles);

                        }
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





