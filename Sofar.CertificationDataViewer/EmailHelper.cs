using MailKit.Net.Smtp;
using MimeKit;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Sofar.DataBaseHelper;

public class EmailHelper
{

    public void sendMail(string body,string recipients)
    {
        //公司邮箱不支持System.Net.Mail，使用MailKit可行  
        try
        {
            string mailbox = System.Configuration.ConfigurationManager.AppSettings["mailbox"].ToString(); 
            string password = System.Configuration.ConfigurationManager.AppSettings["password"].ToString();
            string IsEncrypt = System.Configuration.ConfigurationManager.AppSettings["IsEncrypt"].ToString();
            if (string.IsNullOrEmpty(mailbox)) return;
            if (string.IsNullOrEmpty(password)) return;
            if (IsEncrypt == "YES")
            {
                mailbox = Encrypt.DecryptPassword(mailbox);
                password = Encrypt.DecryptPassword(password);
            }

            var message = new MimeMessage();
            //邮件标题
            message.Subject = $"🚨【预警】{DateTime.Now:yyyy-MM-dd} 认证证书预警信息";
            //发件人邮箱地址
            message.From.Add(new MailboxAddress($"认证证书预警平台", mailbox));

            // 设置收件人、抄送人和密送人

            //收件人变更为由预警信息得出         
            if (!string.IsNullOrWhiteSpace(recipients))
            {
                ProcessEmailAddresses(recipients, message.To, "收件人");
            }
            else
            {
                MessageBox.Show("收件人邮箱地址不能为空");
                return;
            }

            string ccEmails = System.Configuration.ConfigurationManager.AppSettings["Cc_EmailAddress"].ToString();
            if (!string.IsNullOrWhiteSpace(ccEmails))
            {
                ProcessEmailAddresses(ccEmails, message.Cc, "抄送人");
            }

            string bccEmails = System.Configuration.ConfigurationManager.AppSettings["Bcc_EmailAddress"]?.ToString();
            if (!string.IsNullOrWhiteSpace(bccEmails))
            {
                ProcessEmailAddresses(bccEmails, message.Bcc, "密送人");
            }

            // 邮件正文部分
            var bodyPart = new TextPart("html")
            {
                Text = body
            };

            //  Multipart 用于组合正文和附件
            var multipart = new Multipart("mixed");
            multipart.Add(bodyPart);
        
            string IsAddTableAttachment = System.Configuration.ConfigurationManager.AppSettings["IsAddTableAttachment"].ToString();            
            //是否添加表格附件
            if (IsAddTableAttachment == "YES")
            {           
                string TableAttachmentName = System.Configuration.ConfigurationManager.AppSettings["TableAttachmentName"].ToString();
                if (string.IsNullOrEmpty(TableAttachmentName)) return;

                // 添加表格文件附件
                var attachmentPath = Directory.GetCurrentDirectory() + "\\" + TableAttachmentName;
                var attachment = new MimePart("application", "octet-stream")
                {
                    Content = new MimeContent(File.OpenRead(attachmentPath), ContentEncoding.Default),
                    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                    ContentTransferEncoding = ContentEncoding.Base64,
                    FileName = Path.GetFileName(attachmentPath)
                };
                multipart.Add(attachment);
            }
            ////添加压缩包附件
            //var attachmentData = File.ReadAllBytes(@"D:\VS2022_CodeClone\B5_认证数据工具\Sofar.CertificationDataViewer\bin\Debug.zip");
            //var attachmentStream = new MemoryStream(attachmentData);
            //var attachment1 = new MimePart("application", "zip")
            //{
            //    Content = new MimeContent(attachmentStream, ContentEncoding.Default),
            //    ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
            //    ContentTransferEncoding = ContentEncoding.Base64,
            //    FileName = "认证证书统计工具.zip" // 自定义附件名称  
            //};
            //multipart.Add(attachment1);

            // 设置邮件的内容
            message.Body = multipart;
            using (var client = new SmtpClient())
            {
                //绕过默认的证书验证逻辑
                ServicePointManager.ServerCertificateValidationCallback =
                   delegate (Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                   {
                       return true;
                   };

                //SMTP服务器地址\SSL端口\true为启用SSL
                client.Connect("smtp.sofarsolar.com", 465, true); 

                // 发件人邮箱账号和密码
                client.Authenticate($"{mailbox}", $"{password}");
               
                client.Send(message);
                //MessageBox.Show("邮件发送成功");
                client.Disconnect(true);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("邮件发送失败:"+ex);
        }
    }

    /// <summary>
    /// 处理邮箱地址(收件人、抄送人和密送人)
    /// </summary>
    /// <param name="emailAddresses"></param>
    /// <param name="addressList"></param>
    /// <param name="addressType"></param>
    public void ProcessEmailAddresses(string emailAddresses, InternetAddressList addressList, string addressType)
    {
        var addresses = emailAddresses.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(email => email.Trim())
                                      .Where(email => !string.IsNullOrWhiteSpace(email))
                                      .ToList();

        foreach (var email in addresses)
        {
            if (MailboxAddress.TryParse(email, out var mailboxAddress))
            {
                addressList.Add(mailboxAddress);
            }
            else
            {
                MessageBox.Show($"{addressType}邮箱地址非法: {email}");
            }
        }
    }
}
