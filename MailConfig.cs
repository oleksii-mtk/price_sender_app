public class MailConfig
    {
        public required string Username { get; set; }
        public required string AppPassword { get; set; }
        public required string Subject { get; set; }
        public required string Body { get; set; }
        public required string Alias { get; set; }
}

public class AppConfig
{
    public required MailConfig Email { get; set; }
    public required string FolderPath { get; set; }
    public required string Attachment { get; set; }
}