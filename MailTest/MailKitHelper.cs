using MailKit.Net.Imap;
using MailKit;
using MailKit.Net.Pop3;
using MailKit.Net.Smtp;
using MimeKit;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MimeKit.Utils;

namespace MailTest {
    public class MailKitHelper {


        public static void SendEmail(bool enableSSL = false, int port = 25) {
            var host = "smtp.qiye.163.com";
            //var port = 25;//不加密端口，加密端口是456
            var userName = "";
            var password = "";
            var toAddress = "";
            var ccAddress = "";
            var msg = $"测试 send email by MailKit host={host} ssl={enableSSL} port={port}";

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress("fang", userName));
            message.To.Add(new MailboxAddress("fang", toAddress));
            message.Cc.Add(new MailboxAddress("fang", ccAddress));
            message.Subject = msg;


            var textPart = new TextPart("plain");
            textPart.SetText(Encoding.UTF8, msg);

            var multipart = new Multipart("mixed");
            multipart.Add(textPart);

            var fileName = $"测试非常非常非常非常非常非常非常非常非常非常非常非常非常非常长的附件文件名-{DateTime.Now.ToString("yyyyMMdd")}.xlsx";
            //var filePath = Path.Combine(AppContext.BaseDirectory, "Data", fileName);
            //var fileBytes = File.ReadAllBytes(filePath);

            var fileBytes = AsposeHelper.CreateWorkBook("测试");

            var ms = new MemoryStream(fileBytes);

            var contentDisposition = new ContentDisposition(ContentDisposition.Attachment) {
                FileName = fileName
            };

            //重写附件名称编码Rfc2047，支持长文件名
            if (contentDisposition.Parameters.TryGetValue("filename", out Parameter parameter))
                parameter.EncodingMethod = ParameterEncodingMethod.Rfc2047;

            var mimePart = new MimePart("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
                Content = new MimeContent(ms),
                ContentDisposition = contentDisposition,
                ContentTransferEncoding = ContentEncoding.Base64
                //FileName = fileName//有些client接收邮件后，附件名太长会被截断，重写附件名称编码Rfc2047
            };

            multipart.Add(mimePart);

            message.Body = multipart;


            try {
                using var client = new SmtpClient();
                //using var client = new SmtpClient(new ProtocolLogger(Console.OpenStandardOutput()));//输出日志
                client.Connect(host, port, enableSSL);

                // Note: only needed if the SMTP server requires authentication
                client.Authenticate(userName, password);
                client.Send(message);
                client.Disconnect(true);

                Console.WriteLine("发送成功");
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }


        public static void ReceiveEmailPop3(bool enableSSL = false, int port = 110) {

            var host = "pop.qiye.163.com";
            var userName = "";
            var password = "";

            try {
                using (var client = new Pop3Client()) {
                    client.Connect(host, port, enableSSL);

                    client.Authenticate(userName, password);

                    Console.WriteLine($"receive by MailKit Pop3 host={host} ssl={enableSSL} port={port}");

                    //for (int i = 0; i < client.Count; i++) {
                    //    var message = client.GetMessage(i);
                    //    Console.WriteLine("Subject: {0}", message.Subject);
                    //}
                    if (client.Count > 0) {
                        var message = client.GetMessage(client.Count - 1);
                        Console.WriteLine("Subject: {0}", message.Subject);
                    }

                    client.Disconnect(true);
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }

        public static void ReceiveEmailImap(bool enableSSL = false, int port = 143) {

            var host = "imap.qiye.163.com";
            var userName = "";
            var password = "";

            try {
                using (var client = new ImapClient()) {
                    client.Connect(host, port, enableSSL);

                    client.Authenticate(userName, password);

                    // The Inbox folder is always available on all IMAP servers...
                    var inbox = client.Inbox;
                    inbox.Open(FolderAccess.ReadOnly);

                    Console.WriteLine("Total messages: {0}", inbox.Count);
                    Console.WriteLine("Recent messages: {0}", inbox.Recent);

                    Console.WriteLine($"receive by MailKit IMAP host={host} ssl={enableSSL} port={port}");

                    //for (int i = 0; i < inbox.Count; i++) {
                    //    var message = inbox.GetMessage(i);
                    //    Console.WriteLine("Subject: {0}", message.Subject);
                    //}
                    if (inbox.Count > 0) {
                        var message = inbox.GetMessage(inbox.Count - 1);
                        Console.WriteLine("Subject: {0}", message.Subject);
                    }


                    client.Disconnect(true);
                }
            }
            catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
