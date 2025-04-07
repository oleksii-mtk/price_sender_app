using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDataReader;

class Program
{
    static void Main()
    {
        string clientsFile = @"C:\root\test_mail_send\clients.xlsx";

        List<string> emails = GetEmailsFromExcel(clientsFile);

        Console.WriteLine("Считанные адреса:");
        foreach (var email in emails)
        {
            Console.WriteLine(email);
        }
    }

    static List<string> GetEmailsFromExcel(string path)
    {
        // Регистрируем кодировки для ExcelDataReader
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        var emails = new List<string>();

        using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var result = reader.AsDataSet();

            // Пытаемся найти лист "Клиенты"
            var table = result.Tables["Clients"];
            if (table == null)
            {
                Console.WriteLine("Не найден лист 'Клиенты'");
                return emails;
            }

            // Проходимся по строкам и ищем email'ы
            foreach (System.Data.DataRow row in table.Rows)
            {
                foreach (var cell in row.ItemArray)
                {
                    string value = cell?.ToString();
                    if (!string.IsNullOrEmpty(value) && value.Contains("@"))
                    {
                        emails.Add(value.Trim());
                    }
                }
            }
        }

        // Удалим дубли
        return emails.Distinct().ToList();
    }
}
