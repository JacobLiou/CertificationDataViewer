using MailKit.Net.Smtp;
using MimeKit;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Sofar.DataBaseHelper;
using ATE.Business;

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
            message.Subject = $"🚨【预警】{DateTime.Now:yyyy-MM-dd} 认证证书预警信息(每周二更新)";
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
                //绕过默认的证书验证逻辑，信任SSL所有证书（不安全，谨慎使用）
                ServicePointManager.ServerCertificateValidationCallback =
                   delegate (Object obj, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
                   {
                       return true;
                   };

                //SMTP服务器地址\SSL端口\true为启用SSL
                client.Connect("smtp.sofarsolar.com", 465, true); 

                // 发件人邮箱账号和密码
                client.Authenticate($"{mailbox}", $"{password}");
               
                bool emailSent = false; // 邮件发送标志，true代表邮件发送成功
                int retryCount = 0; // 重发次数

                while (!emailSent && retryCount < 3) // 最多尝试重发3次
                {
                    try
                    {
                        client.Send(message); // 发送邮件
                        emailSent = true; // 如果成功发送，邮件发送标志为真                     
                        var emailInfo = $"邮件发送成功\n主题: {message.Subject}\n发件人: {message.From}\n" +
                                        $"收件人: {string.Join(", ", message.To.Mailboxes.Select(m => m.Address))}\n" +
                                        $"抄送人: {string.Join(", ", message.Cc.Mailboxes.Select(m => m.Address))}\n" +
                                        $"密送人: {string.Join(", ", message.Bcc.Mailboxes.Select(m => m.Address))}\n" +
                                        $"正文: {bodyPart.Text.Substring(0, Math.Min(bodyPart.Text.Length, 100))}"; // 限制正文长度为100个字符                     
                        // 记录成功发送邮件的简短信息
                        ATLog.DLog(emailInfo);                     
                    }
                    // 处理各种 SMTP 命令异常
                    catch (SmtpCommandException ex) when (
                        ex.ErrorCode  == SmtpErrorCode.RecipientNotAccepted ||  // 收件人地址被拒绝
                        ex.ErrorCode  == SmtpErrorCode.SenderNotAccepted    ||  // 发件人地址被拒绝
                        ex.ErrorCode  == SmtpErrorCode.MessageNotAccepted   ||  // 邮件内容被拒绝
                        ex.ErrorCode  == SmtpErrorCode.UnexpectedStatusCode ||  // 未预料到的状态代码
                        ex.StatusCode == SmtpStatusCode.MailboxUnavailable      // 邮箱不可用                                                // 
                    )
                    {
                        // 根据不同的错误代码进行日志记录
                        if (ex.ErrorCode == SmtpErrorCode.UnexpectedStatusCode)
                        {      
                            ATLog.DLog($"未预料到的状态代码: {ex.StatusCode}, 错误信息: {ex.Message}");                                               
                            // 根据不同的未预期状态代码进行处理
                            switch (ex.StatusCode)
                            {                               
                                case SmtpStatusCode.ServiceNotAvailable:
                                    // 记录服务不可用的情况
                                    ATLog.DLog($"状态代码: {ex.StatusCode}服务当前不可用，请稍后重试。");
                                    break;
                                case SmtpStatusCode.TransactionFailed:
                                    // 记录事务失败的情况
                                    ATLog.DLog($"状态代码: {ex.StatusCode}邮件传输过程中发生了错误，请检查邮件内容或重试。");
                                    break;                             
                                default:
                                    // 记录所有其他未定义的情况
                                    ATLog.DLog($"状态代码: {ex.StatusCode}发生了未知错误，请检查服务器状态或联系管理员。");
                                    break;
                            }                         
                        }

                        //移除无法发送的收件人邮箱地址
                        var failedRecipient = ex.Mailbox?.Address; // 获取无法发送的邮箱地址
                        if (!string.IsNullOrEmpty(failedRecipient))
                        {
                            // 从收件人列表中移除该地址
                            var mailboxToRemove = message.To.Mailboxes.FirstOrDefault(m => m.Address == failedRecipient);
                            if (mailboxToRemove != null)
                            {
                                message.To.Remove(mailboxToRemove);
                                ATLog.DLog($"已移除无法发送的收件人邮箱地址: {failedRecipient}");
                            }                         

                            // 从抄送人列表中移除该地址
                            mailboxToRemove = message.Cc.Mailboxes.FirstOrDefault(m => m.Address == failedRecipient);
                            if (mailboxToRemove != null)
                            {
                                message.Cc.Remove(mailboxToRemove);
                                ATLog.DLog($"已移除不可用的抄送人邮箱地址: {failedRecipient}");
                            }

                            // 从密送人列表中移除该地址
                            mailboxToRemove = message.Bcc.Mailboxes.FirstOrDefault(m => m.Address == failedRecipient);
                            if (mailboxToRemove != null)
                            {
                                message.Bcc.Remove(mailboxToRemove);
                                ATLog.DLog($"已移除不可用的密送人邮箱地址: {failedRecipient}");
                            }
                        }

                        retryCount++; // 增加重试计数
                        if (retryCount < 3 && message.To.Count > 0)
                        {
                            // 如果仍有剩余的收件人且未达到最大重试次数，则继续尝试发送
                            ATLog.DLog("重新尝试发送邮件...");
                        }
                        else
                        {
                            throw; // 如果达到最大重试次数或没有剩余收件人，则抛出异常
                        }
                    }
                    catch (Exception ex)
                    {
                        // 捕获其他可能的异常并显示错误信息
                        ATLog.DLog("邮件发送失败:" + ex.Message);
                        throw; // 抛出异常
                    }
                }

                if (emailSent)
                {
                    ATLog.DLog("邮件发送成功"); // 如果邮件成功发送，显示成功消息
                }

                client.Disconnect(true); // 断开与 SMTP 服务器的连接
            }
        }
        catch (Exception ex)
        {
            // 异常处理，用于捕获所有未处理的异常
            ATLog.DLog("邮件发送失败:" + ex.Message);
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
                try
                {
                    // 将邮箱地址添加到地址列表中
                    addressList.Add(mailboxAddress);
                }
                catch (Exception ex)
                {
                    // 如果添加过程中发生异常，记录错误信息或进行其他处理
                    ATLog.DLog($"{addressType}邮箱地址添加失败: {email} - 错误: {ex.Message}");
                }
            }
            else
            {
                // 记录无效邮箱地址，继续处理其他地址
                ATLog.DLog($"{addressType}邮箱地址非法: {email}");
            }
        }
    }
}
