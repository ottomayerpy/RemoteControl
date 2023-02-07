using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MimeKit;

namespace RemoteControl
{
    public partial class FormMain : Form
    {
        private const string GlobalIDServer = "0000"; // Только 4 любых символа
        private const string MailAddress = "fivesevenom@gmail.com";
        private const string MailPassword = "health57";
        private const string SmtpServer = "smtp.gmail.com";
        private const int SmtpPort = 465;
        private const string OutMail = "ottomayer57@yandex.ru";

        public FormMain()
        {
            InitializeComponent();
            //using (ImapClient client = new ImapClient())
            //{
            //    client.Connect("imap.gmail.com", 993, true);
            //    client.Authenticate("fivesevenom@gmail.com", "health57");
            //    IMailFolder inbox = client.Inbox;
            //    inbox.Open(FolderAccess.ReadOnly);
            //    label1.Text = inbox.Count.ToString();

            //    var message = inbox.GetMessage(inbox.Count - 1);
            //    label1.Text = message.TextBody;
            //}
            string message = "string message";
            bool SendFile = true;
            string[] Com = { null, null, null, null, "C:/Обои.7z" };

            MimeMessage emailMessage = new MimeMessage();

            emailMessage.From.Add(new MailboxAddress("RemoteControl", OutMail));
            emailMessage.To.Add(new MailboxAddress(string.Empty, OutMail));
            emailMessage.Subject = "Сервер: " + GlobalIDServer;
            
            if (SendFile)
            {
                BodyBuilder builder = new BodyBuilder
                {
                    TextBody = "message builder"
                };
                builder.Attachments.Add(Com[4]);
                emailMessage.Body = builder.ToMessageBody();
            }
            else
            {
                emailMessage.Body = new TextPart(MimeKit.Text.TextFormat.Html)
                {
                    Text = message
                };
            }

            using (SmtpClient client = new SmtpClient())
            {
                client.Connect(SmtpServer, SmtpPort, true);
                client.Authenticate(MailAddress, MailPassword);
                client.Send(emailMessage);
                client.Disconnect(true);
            }
        }
    }
}
