// See https://aka.ms/new-console-template for more information
using MailTest;


Console.WriteLine("是否启用SSL,输入1表示启用，0表示不启用");
var enableSSLStr = Console.ReadLine();
while(string.IsNullOrWhiteSpace(enableSSLStr) || (enableSSLStr != "1" && enableSSLStr != "0")) {
    Console.WriteLine("请输入0或1,输入1表示启用，0表示不启用");
    enableSSLStr = Console.ReadLine();
}

var enableSSL = enableSSLStr == "1";

Console.WriteLine("输入端口号：");
var portStr = Console.ReadLine();
while(string.IsNullOrWhiteSpace(portStr) || !int.TryParse(portStr,out var res)) {
    Console.WriteLine("请输入端口号：");
    portStr = Console.ReadLine();
}
var port = int.Parse(portStr);

//NetMailHelper.SendMail();

MailKitHelper.SendEmail(enableSSL,port);


//MailKitHelper.SendEmail();
//MailKitHelper.ReceiveEmailPop3();

//MailKitHelper.SendEmail(true, 465);
//MailKitHelper.ReceiveEmailPop3(true, 995);

//MailKitHelper.SendEmail();
//MailKitHelper.ReceiveEmailImap();

//MailKitHelper.SendEmail(true, 994);
//MailKitHelper.ReceiveEmailImap(true, 993);

Console.ReadKey();


