using ETE.Data;
using Microsoft.Extensions.Options;
using System.Net;
using System.Net.Mail;

namespace ETE.Services
{
    public class EmailService
    {
        private EmailSettings Settings { get; }

        public EmailService(IOptions<EmailSettings> options)
        {
            Settings = options.Value;
        }

        public async Task SendEmailAsync(string toEmail, string subject, string body)
        {
            var mailMessage = new MailMessage
            {
                From = new MailAddress(Settings.SenderEmail, Settings.SenderName),
                Subject = subject,
                Body = body,
                IsBodyHtml = true,
                Priority = MailPriority.High
            };

            mailMessage.To.Add(toEmail);

            using var client = new SmtpClient
            {
                Host = Settings.Host,
                Port = Settings.Port,
                EnableSsl = Settings.UseSSL,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(Settings.UserName, Settings.Password)
            };

            await client.SendMailAsync(mailMessage);
        }
    }
}