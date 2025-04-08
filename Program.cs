using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDataReader;
using Microsoft.Extensions.Configuration;

class Program
{
    static void Main()
    {
        var config = LoadConfig();
        string emailUser = config["Email:User"];
        string emailPassword = config["Email:Password"];
        string folderPath = config["FolderPath"];
        string clientsFile = Path.Combine(folderPath, "clients.xlsx");  
        Console.WriteLine("Reading emails from Excel file..."+clientsFile);
        //List<string> emails = GetEmailsFromExcel(clientsFile);


    }

    static IConfiguration LoadConfig()
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
        return builder.Build();
    }


}
