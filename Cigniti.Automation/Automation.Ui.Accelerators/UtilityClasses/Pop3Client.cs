using OpenPop.Mime;
using OpenPop.Pop3;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace Automation.Ui.Accelerators.UtilityClasses
{
    /// <summary>
    /// Message Model
    /// </summary>
    public class MessageModel
    {
        public string MessageID;
        public string FromID;
        public string FromName;
        public string Subject;
        public string Body;
        public string Html;
        public string FileName;
        public List<MessagePart> Attachment;
    }

    public class Pop3EmailClient : IDisposable
    {
        private Pop3Client mPop3CLient = null;
        public Pop3Client Pop3ClientObj
        {
            get
            {
                this.mPop3CLient = new Pop3Client();
                mPop3CLient.Connect(ConfigurationManager.AppSettings.Get("EmailHostName"), Convert.ToInt32(ConfigurationManager.AppSettings.Get("EmailPort")), Convert.ToBoolean(ConfigurationManager.AppSettings.Get("EmailUseSsl")));
                mPop3CLient.Authenticate(ConfigurationManager.AppSettings.Get("EmailId"), ConfigurationManager.AppSettings.Get("Password"));
                return this.mPop3CLient;
            }
        }

        public MessageModel GetEmailContent(int messageNumber)
        {
            MessageModel message = new MessageModel();

            MessagePart plainTextPart = null, HTMLTextPart = null;

            Message objMessage = Pop3ClientObj.GetMessage(messageNumber);

            message.MessageID = objMessage.Headers.MessageId == null ? "" : objMessage.Headers.MessageId.Trim();

            message.FromID = objMessage.Headers.From.Address.Trim();
            message.FromName = objMessage.Headers.From.DisplayName.Trim();
            message.Subject = objMessage.Headers.Subject.Trim();

            plainTextPart = objMessage.FindFirstPlainTextVersion();
            message.Body = (plainTextPart == null ? "" : plainTextPart.GetBodyAsText().Trim());

            HTMLTextPart = objMessage.FindFirstHtmlVersion();
            message.Html = (HTMLTextPart == null ? "" : HTMLTextPart.GetBodyAsText().Trim());

            List<MessagePart> attachment = objMessage.FindAllAttachments();

            if (attachment.Count > 0)
            {
                message.FileName = attachment[0].FileName.Trim();
                message.Attachment = attachment;
            }

            return message;
        }

        /// <summary>
        /// Download Files
        /// </summary>
        /// <param name="message"></param>
        /// <param name="fileNames"></param>
        public void DownloadFile(MessageModel message, params string[] fileNames)
        {
            try
            {
                List<MessagePart> attachment = message.Attachment;
                List<MessagePart> finalAttachment = new List<MessagePart>();

                for (int i = 0; i < fileNames.Length; i++)
                {
                    for (int j = 0; j < attachment.Count; j++)
                    {
                        if (attachment[j].FileName.Contains(fileNames[i]))
                        {
                            finalAttachment.Add(attachment[j]);
                        }
                    }
                }

                if (finalAttachment.Count.Equals(0))
                {
                    throw new Exception("No file found in the attachment");
                }
                for (int i = 0; i < finalAttachment.Count; i++)
                {
                    byte[] content = finalAttachment[i].Body;

                    string downloadPath = Directory.GetCurrentDirectory() + "\\EmailDownloads";
                    if (!Directory.Exists(downloadPath))
                    {
                        Directory.CreateDirectory(downloadPath);
                    }
                    using (StreamWriter sw = new StreamWriter(downloadPath + "\\" + finalAttachment[i].FileName))
                    {
                        sw.BaseStream.Write(content, 0, content.Length);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw new Exception(string.Format("Failed at DownloadFile(). ErrorMessage: {0}. InnerExceptionMessage: {1}", ex.Message, ex.InnerException.Message));
            }
        }

        /// <summary>
        /// Gets the Email Count
        /// </summary>
        /// <returns></returns>
        public int EmailCount()
        {
            try
            {
                return Pop3ClientObj.GetMessageCount();
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format("Failed at EmailCount(). ErrorMessage: {0}. InnerExceptionMessage: {1}", ex.Message, ex.InnerException.Message));
            }
        }

        public void Dispose()
        {
            Pop3ClientObj.Disconnect();
            Pop3ClientObj.Dispose();
        }
    }
}