using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;

class Program
{
    static void Main()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        var config = LoadConfig();
        string emailUser = config["Email:User"];
        string emailPassword = config["Email:Password"];
        string folderPath = config["FolderPath"];
        string clientsFile = Path.Combine(folderPath, "clients.xlsx");
        //Console.WriteLine(clientsFile);

        List<string> emails = GetEmailsFromExcel(clientsFile);
        //Console.WriteLine($"Найдено адресов: {emails.Count}");

        string priceFile = GetLatestPriceFile(folderPath);
        Console.WriteLine(priceFile);

        foreach (string email in emails)
        {
            try
            {
                SendEmailWithAttachment(email, priceFile, emailUser, emailPassword);
                Console.WriteLine($"Письмо отправлено: {email}");
                Thread.Sleep(12000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при отправке на {email}: {ex.Message}");
            }
        }

        Console.WriteLine("Рассылка завершена.");

    }

    static IConfiguration LoadConfig()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
        return builder.Build();
    }

    static string GetLatestPriceFile(string folderPath)
    {
        var files = Directory.GetFiles(folderPath, "price_nks13_*.xls");
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

    static void SendEmailWithAttachment(string toEmail, string attachmentPath, string emailUser, string emailPass)
    {
        var fromAddress = new MailAddress(emailUser, "Alex from NKS13");
        var toAddress = new MailAddress(toEmail);

        var smtp = new SmtpClient
        {
            Host = "smtp.gmail.com",
            Port = 587,
            EnableSsl = true,
            Credentials = new NetworkCredential(emailUser, emailPass)
        };

        using (var message = new MailMessage(fromAddress, toAddress)
        {
            Subject = "Обновлённый прайс-лист",
            Body = "Добрый день!\n\nВо вложении актуальный прайс-лист.\n\nС уважением,\nNKS13"
        })
        {
            message.Attachments.Add(new Attachment(attachmentPath));
            smtp.Send(message);
        }
    }
}
