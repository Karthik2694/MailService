# MailService
namespace MailService
{
    using System;
    using System.Configuration;
    using log4net;
    using SRA.Proof.Common;
    using System.Data;
    using System.Net.Mail;
    using System.IO;
    using System.Net.Mime;
    using System.Transactions;

    public class EmailSenderHost
    {
        #region Private fields

        /// <summary>
        /// A <see cref="log4net.ILog"/> instance for logging.
        /// </summary>
        private readonly ILog _sysLog;

        /// <summary>
        /// A <see cref="TaskManager{I,J}"/> that is used for
        /// scheduling tasks in specified intervals.
        /// </summary>
        private static TaskManager<string, string> taskManager;

        #endregion

        #region Default Constructor

        /// <summary>
        /// Default Constructor
        /// </summary>
        public EmailSenderHost()
        {
            _sysLog = LogManager.GetLogger(GetType());
        }
        #endregion

        #region Methods

        /// <summary>
        /// Method to initialize the service at the scheduled interval
        /// </summary>
        public void ServiceInitialize()
        {
            try
            {
                _sysLog.Debug("Entering ServiceInitialize");

                int intervalInMins = (int.TryParse(ConfigurationManager.
                AppSettings[Constants.JobInterval], out intervalInMins) ?
                intervalInMins : 1);

                _sysLog.DebugFormat("IntervalInMinutes = {0}", intervalInMins);

                taskManager = new TaskManager<string, string>();

                TaskBase<string, string> task =
                    new RecurringTask<string, string>("1",
                        DateTime.Now.AddSeconds(20), (intervalInMins * 60),
                        false);                              

                taskManager.Add(task);
                
                taskManager.InvokeTask += OnInvokeTask;                

                _sysLog.Debug("Exiting ServiceInitialize");
            }
            catch (Exception ex)
            {
                _sysLog.Error("ERROR in ServiceInitialize", ex);

                throw;
            }
        }
          

        /// <summary>
        /// Represents a method that sends the mail to the users.
        /// </summary>
        public void SendMail()
        {
            string[] toBeDeleted = new string[0];

            int i = 0;

            CryptoManager crypto = null;

            string exceptionId=null;

            try
            {
                DBHandler handler = new DBHandler();

                DataTable table = handler.GetEmailDetails();

                int noOfEmails = table.Rows.Count;

                toBeDeleted = new string[noOfEmails];

                _sysLog.DebugFormat("No of Emails found: {0}", noOfEmails);

                foreach (DataRow row in table.Rows)
                {
                    using (TransactionScope scope =
                      new TransactionScope(TransactionScopeOption.RequiresNew, System.TimeSpan.MaxValue))
                    {

                        string id = row[Constants.ID].ToString();

                        exceptionId = id;

                        string from = row[Constants.From].ToString();

                        string to = row[Constants.To].ToString();

                        string Cc = row[Constants.Cc].ToString();

                        string Bcc = row[Constants.BCC].ToString();

                        string subject = row[Constants.Subject].ToString();

                        bool isHTML = (bool)row[Constants.IsHTML];

                        string attachmentPath =
                            row[Constants.Attachment].ToString();

                        string body = row[Constants.Content].ToString();

                        _sysLog.Debug(string.Format(
                            "{0}:{1} {2}:{3} {4}:{5} {6}:{7} {8}:{9} {10}:{11} {12}:{13}",
                            Constants.From, from, Constants.To, to,
                            Constants.Cc, Cc, Constants.BCC, Bcc,
                            Constants.Subject, subject,
                            Constants.IsHTML, isHTML.ToString(),
                            Constants.Attachment, attachmentPath));

                        MailMessage mail = null;
                        mail = new MailMessage(from, to);

                        if (!string.IsNullOrEmpty(subject))
                        {
                            mail.Subject = subject;
                            _sysLog.Debug("Subject:" + mail.Subject.ToString());
                        }

                        if (!string.IsNullOrEmpty(Cc))
                        {

                            string[] CCId = Cc.Split(',');

                            foreach (string CCEmail in CCId)
                            {
                                mail.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id
                                _sysLog.Debug("cc identified" + CCEmail.ToString());
                            }

                        }

                        if (!string.IsNullOrEmpty(Bcc))
                        {

                            string[] BCCId = Bcc.Split(',');

                            foreach (string BCCEmail in BCCId)
                            {
                                mail.Bcc.Add(new MailAddress(BCCEmail)); //Adding Multiple BCC email Id
                                _sysLog.Debug("Bcc identified" + BCCEmail.ToString());
                            }

                        }


                        string destFile = string.Empty;

                        if (!string.IsNullOrEmpty(attachmentPath))
                        {

                            _sysLog.Debug("***Entering if file contain Attachment***");

                            string targetFilePath = ConfigurationManager.
                                AppSettings[Constants.FileTargetPath]; ;

                            _sysLog.Debug("***Entering if file contain Attachment***");

                            string sourcePath =
                                Path.GetDirectoryName(attachmentPath);

                            _sysLog.Debug("***Entering if file contain Attachment***");

                            string srcFileName =
                                Path.GetFileName(attachmentPath);

                            _sysLog.Debug("***Entering if file contain Attachment***");

                            string targetPath = targetFilePath;

                            string targetFileName = string.Format("{0}{1}",
                                Guid.NewGuid().ToString(), srcFileName);

                            string sourceFile = System.IO.
                                Path.Combine(sourcePath, srcFileName);

                            destFile = System.IO.
                                Path.Combine(targetPath, targetFileName);

                            if (!System.IO.Directory.Exists(targetPath))
                            {
                                System.IO.Directory.CreateDirectory(targetPath);
                            }

                            System.IO.File.Copy(sourceFile, destFile, true);

                            Attachment attachment = new Attachment
                                (destFile, MediaTypeNames.Application.Octet);

                            ContentDisposition disposition =
                                attachment.ContentDisposition;

                            disposition.CreationDate =
                                File.GetCreationTime(destFile);

                            disposition.ModificationDate =
                                File.GetLastWriteTime(destFile);

                            disposition.ReadDate =
                                File.GetLastAccessTime(destFile);

                            disposition.FileName = srcFileName; 

                            disposition.Size =
                                new FileInfo(destFile).Length;

                            disposition.DispositionType =
                                DispositionTypeNames.Attachment;

                            mail.Attachments.Add(attachment);

                            toBeDeleted[i] = destFile;

                            i++;

                            _sysLog.Debug("***Exiting if file contain Attachment***");
                        }

                        mail.Body = body;

                        _sysLog.Debug("Body:" + mail.Body.ToString());

                        mail.IsBodyHtml = isHTML;

                        _sysLog.Debug("IsBodyHtml:" + mail.IsBodyHtml.ToString());

                        _sysLog.Debug("***Entering to get mail credential***");

                        SmtpClient smtp = new SmtpClient();

                        string host = ConfigurationManager.
                            AppSettings[Constants.Host];                        

                        string userName = ConfigurationManager.
                            AppSettings[Constants.UserName];
                                                
                        string pwd = ConfigurationManager.
                            AppSettings[Constants.Password];
                        
                        int port = Convert.ToInt32(ConfigurationManager.
                            AppSettings[Constants.Port]);
                        
                        bool enableSsl =
                            Convert.ToBoolean(ConfigurationManager.
                            AppSettings[Constants.EnableSsl]);
                                                
                        bool useDefaultCredentials =
                            Convert.ToBoolean(ConfigurationManager.
                            AppSettings[Constants.UseDefaultCredentials]);

                        _sysLog.Debug(string.Format("Credential detail host - {0} username - {1}, Password - {2}, Port - {3} ",host,userName,pwd,port));
                        
                        smtp.Host = host;

                        smtp.EnableSsl = enableSsl;

                        System.Net.NetworkCredential credentials =
                            new System.Net.NetworkCredential();

                        crypto = new CryptoManager();

                        //credentials.UserName = crypto.Decrypt(userName);

                        //credentials.Password = crypto.Decrypt(pwd);

                        credentials.UserName = userName;

                        credentials.Password = pwd;

                        smtp.UseDefaultCredentials = useDefaultCredentials;

                        smtp.Credentials = credentials;

                        smtp.Port = port;

                        _sysLog.Debug("***Exiting to get mail credential***");

                        _sysLog.Debug("***Entering to sending Mail***");

                        smtp.Send(mail);

                        _sysLog.Debug("***Exiting from after sending Mail***");

                        _sysLog.Debug("***Update Query primary key:" + id);

                        handler.UpdateSentMail(id);

                        scope.Complete();
                    }
                }
            }
            catch (Exception ex)
            {
                _sysLog.Error("ERROR in Email Send so details moving to Exceptional details table - " + ex);

                DBHandler handler = new DBHandler();

                handler.UpdateExceptionDetail(exceptionId,ex.ToString());

                SendMail();

            }

            finally
            {
                foreach (string fileName in toBeDeleted)
                {
                    if (File.Exists(fileName))
                        File.Delete(fileName);
                }
            }
        }

        #endregion

        #region Event handler methods

        /// <summary>
        /// Represents the method that shall be triggered based in the 
        /// configured interval. This shall initiate the send mail operation
        /// </summary>
        private void OnInvokeTask(object sender, TaskEventArgs<string, string> e)
        {
            try
            {
                if (_sysLog.IsDebugEnabled)
                    _sysLog.Debug("Entering into OnInvokeTask");
                
                DBHandler handler = new DBHandler();

                int flag = 1, status;                

                _sysLog.Debug("Before entering into GetTaskManagerFlag Method");

                status = handler.GetTaskManagerFlag(flag);

                _sysLog.Debug("Task Manager flag is - " + status.ToString());

                if (status == 0)
                {
                    _sysLog.Debug("Entering into Update the InvokeTaskStatus to stop the next task invoke.");

                    handler.UpdateInvokeTask(2);

                    SendMail();

                    _sysLog.Debug("Entering into Update the InvokeTaskStatus to release the next task invoke.");

                    handler.UpdateInvokeTask(3);
                }
                else
                {
                    _sysLog.Debug("Already one task is running.Denying this invoke task.");
                }
                //End of change 
                if (_sysLog.IsDebugEnabled)
                    _sysLog.Debug("Exiting OnInvokeTask");
            }

            catch (Exception ex)
            {
                if (_sysLog.IsErrorEnabled)
                    _sysLog.Error("ERROR in OnInvokeTask", ex);
            }
        }
        #endregion
    }
}
