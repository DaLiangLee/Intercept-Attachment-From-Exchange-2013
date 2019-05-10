// ***************************************************************
// <copyright file="InterceptAttachment.cs" company="Microsoft">
//     Copyright (C) Microsoft Corporation.  All rights reserved.
// </copyright>
// <summary>
// Sends incoming messages to an out-of-process COM server, which
// asynchronously examines the message and returns a modified
// version of the message.
// </summary>
// ***************************************************************

namespace Microsoft.Exchange.Samples.InterceptAttachment
{
    using System;
    using System.IO;
    using System.Diagnostics;
    using Microsoft.Exchange.Data.Transport;
    using Microsoft.Exchange.Data.Transport.Routing;
    using System.Threading.Tasks;

    /// <summary>
    /// The agent factory for the ComInterop sample.
    /// </summary>
    public class InterceptAttachmentAgentFactory : RoutingAgentFactory
    {
        /// <summary>
        /// Creates a new InterceptAttachmentAgent.
        /// </summary>
        /// <param name="server">An Exchange gateway server.</param>
        /// <returns>A new InterceptAttachmentAgent.</returns>
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            return new InterceptAttachmentAgent();
        }
    }

    /// <summary>
    /// SmtpReceiveAgent for the ComInterop sample.
    /// </summary>
    public class InterceptAttachmentAgent : RoutingAgent
    {
        /// <summary>
        /// The constructor registers an end-of-data event handler.
        /// </summary>
        public InterceptAttachmentAgent()
        {
            this.OnSubmittedMessage += new SubmittedMessageEventHandler(this.SubmittedMessageHandler);
        }

        /// <summary>
        /// Invoked by Exchange when a message has been submitted.
        /// </summary>
        /// <param name="source">The source of this event.</param>
        /// <param name="args">Arguments for this event.</param>
        void SubmittedMessageHandler(SubmittedMessageEventSource source, QueuedMessageEventArgs args)
        {
            string sPath = @"C:\test\";
            if (args.MailItem.Message.Attachments != null && args.MailItem.Message.Attachments.Count > 0)
            {
                var mailItem = args.MailItem;

                try
                {
                    /// 获取收件人
                    foreach (var recipient in mailItem.Recipients)
                    {
                        string dirName = sPath + recipient.Address.LocalPart + '@' + recipient.Address.DomainPart;
                        if (!Directory.Exists(dirName))
                        {
                            Directory.CreateDirectory(dirName);
                        }
                        MoveFile(dirName  + "\\", args.MailItem.Message.Attachments);
                    }
                }
                catch (Exception ex)
                {
                }
                
            }
        }
        public static void MoveFile(string sPath, Data.Transport.Email.AttachmentCollection attachments)
        {
            foreach (var attachment in attachments) {

                var readStream = attachment.GetContentReadStream();
                // Save the file in the attachment
                byte[] bytes = new byte[readStream.Length];
                readStream.Read(bytes, 0, bytes.Length);
                try
                {
                    // 设置当前流的位置为流的开始
                    readStream.Seek(0, SeekOrigin.Begin);
                    // 把 byte[] 写入文件
                    string fullPath = sPath + attachment.FileName;

                    // 判断是否存在同名文件
                    if (File.Exists(fullPath))
                    {
                        string directory = Path.GetDirectoryName(fullPath);
                        string filename = Path.GetFileNameWithoutExtension(fullPath);
                        string extension = Path.GetExtension(fullPath);
                        int counter = 1;

                        while (File.Exists(fullPath))
                        {
                            string newFilename = string.Format("{0}({1}){2}", filename, counter, extension);
                            fullPath = Path.Combine(directory, newFilename);
                            counter++;
                        }
                    }

                    FileStream fs = new FileStream(fullPath, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);

                    bw.Write(bytes);

                    bw.Close();
                    fs.Close();
                }
                catch(Exception ex)
                {
                }
            }
        }
    }
}

