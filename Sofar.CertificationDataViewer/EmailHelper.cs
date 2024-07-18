using MailKit.Net.Smtp;
using MimeKit;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Sofar.DataBaseHelper;

public class EmailHelper
{

    public void sendMail(List<string> toEmails, string body)
    {
        //公司邮箱不支持System.Net.Mail，使用MailKit可行  
        try
        {
            string mailbox = System.Configuration.ConfigurationManager.AppSettings["mailbox"].ToString();
            if (string.IsNullOrEmpty(mailbox)) return;

            mailbox = Encrypt.DecryptPassword(mailbox);

            string password = System.Configuration.ConfigurationManager.AppSettings["password"].ToString();
            if (string.IsNullOrEmpty(password)) return;

            password = Encrypt.DecryptPassword(password);

            var message = new MimeMessage();
            //邮件标题
            message.Subject = $"🚨【预警】{DateTime.Now:yyyy-MM-dd} 认证证书预警信息";
            //发件人邮箱地址
            message.From.Add(new MailboxAddress($"认证证书预警平台", mailbox)); 

            //同时发送给多个收件人
            foreach (var email in toEmails)
            {
                message.To.Add(new MailboxAddress(null, email));

            }

            // 设置抄送（Cc）  
            message.Cc.Add(new MailboxAddress("", "cengtongnian@sofarsolar.com"));

            //// 设置密送（Bcc）  
            //message.Bcc.Add(new MailboxAddress("", "xxx.com"));

            // 邮件正文部分
            var bodyPart = new TextPart("html")
            {
                Text = body
            };

            //  Multipart 用于组合正文和附件
            var multipart = new Multipart("mixed");
            multipart.Add(bodyPart);

            // 添加表格文件附件
            var attachmentPath = Directory.GetCurrentDirectory() + "\\" + "产品认证统计表_更新.xlsx"; 
            var attachment = new MimePart("application", "octet-stream")
            {
                Content = new MimeContent(File.OpenRead(attachmentPath), ContentEncoding.Default),
                ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                ContentTransferEncoding = ContentEncoding.Base64,
                FileName = Path.GetFileName(attachmentPath)
            };
            multipart.Add(attachment);

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
            //MessageBox.Show("邮件发送失败");
        }
    }
}
