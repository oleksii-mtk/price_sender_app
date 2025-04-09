using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;
using System.Diagnostics;


class Program
{
    static void Main()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var appConfig = LoadAppConfig();
        var mail = appConfig.Email;
        string folderPath = appConfig.FolderPath;
        string attachment = appConfig.Attachment;

        if (string.IsNullOrWhiteSpace(mail.Username) ||
        string.IsNullOrWhiteSpace(mail.AppPassword) ||
        string.IsNullOrWhiteSpace(mail.Subject) ||
        string.IsNullOrWhiteSpace(mail.Body) ||
        string.IsNullOrWhiteSpace(folderPath))
        {
            Console.WriteLine("❌ Missing configuration in appsettings.json.");
            Environment.Exit(1);
        }

        

        string clientsFile = Path.Combine(folderPath, "clients.xlsx");
        List<string> emails = GetEmailsFromExcel(clientsFile);
        string logFile = Path.Combine(folderPath, "log.txt");
        int count = 0;

        string priceFile = GetLatestPriceFile(folderPath, attachment);
        Console.WriteLine(priceFile);

        var nextSendTime = DateTime.Now;
        foreach (string email in emails)
        {
            try
            {
                var now = DateTime.Now;
                if (now < nextSendTime)
                {
                    var waitTime = nextSendTime - now;
                    Thread.Sleep(waitTime);
                }
                var stopwatch = Stopwatch.StartNew();

                SendEmailWithAttachment(email, priceFile, appConfig);
                stopwatch.Stop();

                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} OK     → {email}  {stopwatch.Elapsed.TotalSeconds:F1}s";
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
                Console.WriteLine(logMessage);
                nextSendTime = DateTime.Now.AddSeconds(12 - stopwatch.Elapsed.TotalSeconds);
            }
            catch (Exception ex)
            {
                string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} ERROR     → {email}: {ex.Message}";
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
                Console.WriteLine(logMessage);
            }

            count++;

            if (count % 100 == 0)
            {
                Console.WriteLine($"⏸ Pause 30 minutes after {count} emails...");
                File.AppendAllLines(logFile, new[] { $"-- {DateTime.Now:yyyy-MM-dd HH:mm:ss} — PAUSE 30 min after {count} emails --" });
                Thread.Sleep(30 * 60 * 1000);
            }
         
        }

        Console.WriteLine("Done.");

    }

    static AppConfig LoadAppConfig()
    {
        var config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        //AppConfig appConfig = new AppConfig();
        //config.Bind(appConfig);
        var email = new MailConfig
        {
            Alias = config["Email:Alias"],
            Username = config["Email:Username"],
            AppPassword = config["Email:AppPassword"],
            Subject = config["Email:Subject"],
            Body = config["Email:Body"]
        };

        return new AppConfig
        {
            FolderPath = config["FolderPath"],
            Attachment = config["Attachment"],
            Email = email
        };

     
    }

    static string GetLatestPriceFile(string folderPath, string attachment)
    {
        var files = Directory.GetFiles(folderPath, attachment);
        if (files.Length == 0)
            throw new FileNotFoundException("File not found!");
        return files.OrderByDescending(File.GetLastWriteTime).First();
    }
    
    static List<string> GetEmailsFromExcel(string filePath)
    {
        List<string> emails = new List<string>();
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                while (reader.Read())
                {
                    // Assuming the email is in the first column
                    if (reader.GetValue(0) != null)
                    {
                        emails.Add(reader.GetValue(0).ToString());
                    }
                }
            }
        }
        return emails;
    }

    static void SendEmailWithAttachment(string toEmail, string attachmentPath, AppConfig appConfig)
    {
        var fromAddress = new MailAddress(appConfig.Email.Username, appConfig.Email.Alias);
        var toAddress = new MailAddress(toEmail);

        var smtp = new SmtpClient
        {
            Host = "smtp.gmail.com",
            Port = 587,
            EnableSsl = true,
            Credentials = new NetworkCredential(appConfig.Email.Username, appConfig.Email.AppPassword)
        };

        using (var message = new MailMessage(fromAddress, toAddress)
        {
            Subject = appConfig.Email.Subject,
            Body = appConfig.Email.Body
        })
        {
            message.Attachments.Add(new Attachment(attachmentPath));
            smtp.Send(message);
        }
    }
}
