using Org.BouncyCastle.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace MailTest {
    /// <summary>
    /// Sytem.Net.Mail 只支持显示SSL
    /// Explicit SSL
    /// System.net.mail only supports "explicit SSL". Explicit SSL starts as unencrypted on port 25, then issues a startdls and switches to an encrypted connection.See RFC 2228.
    /// Explicit sll wowould go something like:Connect on 25-> starttls (starts to encrypt)-> authenticate-> send data
    /// If the SMTP server expects SSL/TLS connection right from the start then this will not work
    /// 
    /// Implicit SSL
    /// There is no way to use implicit SSL (smtps) with system. net. mail. implicit SSL wowould have the entire connection is wrapped in an SSL layer. A specific port wocould be used (Port 465 is common ). there is no formal RFC covering implicit SSL.
    /// Implicit sll wowould go something like:Start SSL (start encryption)-> connect-> authenticate-> send data
    /// This is not considered a bug, it's a feature request. There are two types of SSL authentication for SMTP, and we only support one (by design)-explicit SSL.
    /// Address: http://blogs.msdn.com/webdav_101/archive/2008/06/02/system-net-mail-with-ssl-to-authenticate-against-port-465.aspx
    /// </summary>
    public class NetMailHelper {
        public static void SendMail(bool enableSSL = false, int port = 25) {
            var host = "smtp.qiye.163.com";
            //var port = 25;//不加密端口，加密端口是456
            var UserName = "";
            var PassWord = "";
            var DisplayName = "test";
            var address = "";
            
            var msg = $"send email by System.Net.Mail host={host} ssl={enableSSL} port={port}";

            //System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            //ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };

            //邮件发送类 
            MailMessage mail = new MailMessage();
            try {
                //是谁发送的邮件 
                mail.From = new MailAddress(UserName, DisplayName);
                //发送给谁 

                mail.To.Add(address);

                //标题 
                mail.Subject = msg;
                //编码 
                mail.BodyEncoding = System.Text.Encoding.GetEncoding("UTF-8");
                mail.SubjectEncoding = System.Text.Encoding.GetEncoding("UTF-8");
                //发送优先级 
                mail.Priority = MailPriority.Normal;
                //邮件内容 
                mail.Body = msg;//.Replace("\r\n", "<br>\r\n");
                                //是否HTML形式发送 
                mail.IsBodyHtml = true;
                //邮件服务器和端口 
                SmtpClient smtp = new SmtpClient(host, port);
                //指定发送方式 
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                //指定登录名和密码 
                smtp.Credentials = new System.Net.NetworkCredential(UserName, PassWord);
                //超时时间 
                smtp.EnableSsl = enableSSL;//是否加密传输
                smtp.Timeout = 10000;


                var fileName = "测试附件.xlsx";
                var filePath = Path.Combine(AppContext.BaseDirectory, "Data", fileName);
                var fileBytes = File.ReadAllBytes(filePath);
                var ms = new MemoryStream(fileBytes);                
                var attachment = new Attachment(ms, MediaTypeNames.Application.Octet) {
                    Name = fileName
                };
                mail.Attachments.Add(attachment);

                try {
                    smtp.Send(mail);
                    Console.WriteLine("发送成功");
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
            finally {
                mail.Dispose();
            }
        }
    }
}
