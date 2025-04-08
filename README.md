# 📧 PriceSender

**PriceSender** is a lightweight C# console application that automatically sends a daily updated price list to a list of client email addresses.

## 🔧 Features

- 📤 Sends emails with an attached `.xls` price file
- 📅 Designed to run daily at 9:00 AM
- 📄 Reads client emails from an Excel file (`clients.xlsx`, sheet name: `Clients`)
- 📁 Automatically finds the latest price file (`price_nks13_dd_mm_yy.xls`)
- 🔒 Keeps credentials out of the source code using `appsettings.json`
- 🧘 Sends one email at a time with delay to avoid spam filters

## 📁 Folder Structure

```
/PriceSender/
│
├── PriceSender.exe
├── appsettings.json         ← (excluded from git)
├── appsettings.template.json
├── clients.xlsx
├── price_nks13_07_04_2025.xls
└── README.md
```

## ⚙️ Configuration

Create an `appsettings.json` file in the same folder as the executable:

```json
{
  "Email": {
    "Username": "your_email@gmail.com",
    "AppPassword": "your_app_password"
  },
  "FolderPath": "C:\\Path\\To\\Your\\Folder"
}
```

> 💡 Use double backslashes `\\` in paths on Windows.

**Don't commit this file to GitHub**. It's already listed in `.gitignore`.

Instead, include `appsettings.template.json` as an example config.

## 🛠 Build & Publish (Self-contained EXE)

```bash
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

## ⏰ Auto Start via Task Scheduler (Windows)

1. Open **Task Scheduler**
2. Create a new task:
   - Trigger: Daily at 09:00
   - Action: Start a program → point to `PriceSender.exe`
   - Working directory: folder with `appsettings.json` and Excel files
3. Ensure the scheduled user account has access to the folder and files

## ✅ SMTP Setup (Gmail Example)

1. Enable **2-Step Verification** in your Google account
2. Generate an **App Password**
3. Use it in `appsettings.json`

For other providers (Outlook, etc.), use the correct SMTP server and credentials.

## 📌 Notes

- Default delay between messages is 12 seconds (`Thread.Sleep(12000)`)
- Total sending time for 400 emails ≈ **80 minutes**
- Gmail limits: ~500 messages/day for free accounts

## 👨‍💻 Author

**Alex Matyko**  
Software Engineer / Data Enthusiast  
[LinkedIn](https://linkedin.com/in/...) • [GitHub](https://github.com/oleksii-mtk)