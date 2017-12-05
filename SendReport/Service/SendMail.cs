using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SendReport.Service
{
   public class SendMail
    {
        private string gmail_account = "shadow061103@gmail.com";//帳號
        private string gmail_password = "XXXXXXXXXX";
        private string gmail_mail = "shadow061103@gmail.com.tw";

        public void SendGMail(string MailBody, List<string> ToMail,List<string> attach)
        {
            //建立寄信用smtp物件 gmail
            SmtpClient SmtpServer = new SmtpClient();
            //設定使用的Port
            SmtpServer.Port = 587;
            SmtpServer.Host = "smtp.gmail.com";
            //建立使用者憑據 要設定自己帳號
            SmtpServer.Credentials = new System.Net.NetworkCredential(gmail_account, gmail_password);
            //開啟SSL
            SmtpServer.EnableSsl = true;
            //宣告信件內容物件
            MailMessage mail = new MailMessage();
            //設定來源信箱
            mail.From = new MailAddress("Service@mycard520.com.tw","MyCard");
            //設定收件者信箱
            foreach(var t in ToMail)
            {
                mail.To.Add(t);
            }

            
            //設定信件主旨
            mail.Subject = DateTime.Now.ToString("yyyyMMdd")+"報表測試信";
            //設信件內容
            mail.Body = MailBody;
            //設定信件內容為html格式
            mail.IsBodyHtml = true;

            foreach (var t in attach)
            {
                Attachment attachment = new Attachment(t);
                mail.Attachments.Add(attachment);
            }
            
            //送出信件
            SmtpServer.Send(mail);

        }
    }
}
