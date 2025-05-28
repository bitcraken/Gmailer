using System.Net;
using System.Net.Mail;
using Microsoft.Extensions.Configuration;

internal class Program
{
    private static async Task Main(string[] args)
    {
        // Build configuration
        var config = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // Get the SMTP settings
        var smtpSettings = config.GetSection("AppSettings:SmtpSettings").Get<SmtpSettings>()
                                            ?? throw new Exception("Cannot read from settings.json file, or missing configuration settings");

        // Get all email recipents info
        var recipents = config.GetSection("AppSettings:EmailSettings:Recipients").Get<List<RecipientInfo>>()
                                            ?? throw new Exception("Cannot read from settings.json file");

        // Get subject and body for all the emails
        var subject = config.GetSection("AppSettings:EmailSettings:Subject").Get<string>() ?? throw new Exception("Cannot read from settings.json file");
        var body = config.GetSection("AppSettings:EmailSettings:Body").Get<string>() ?? throw new Exception("Cannot read from settings.json file");

        // Get the sender info
        var sender = config.GetSection("AppSettings:EmailSettings:Sender").Get<string>() ?? throw new Exception("Cannot read from settings.json file");
        var senderEmail = config.GetSection("AppSettings:EmailSettings:SenderEmail").Get<string>() ?? throw new Exception("Cannot read from settings.json file");
        var attachmentFolder = config.GetSection("AppSettings:EmailSettings:AttachmentFolder").Get<string>() ?? throw new Exception("Cannot read from settings.json file");

        // Tasks to hold of all the messages that was sent
        var emailTasks = new List<Task<EmailStatus>>();

        // Prepare and send emails to all recipients 
        foreach (var recipient in recipents)
        {
            emailTasks.Add(SendEmailAsync(recipient, subject, body, sender, senderEmail, attachmentFolder, smtpSettings));
        }

        // Wait for all the messages to be sent
        var tasks = await Task.WhenAll(emailTasks);

        // Print out the sent status of all the messages
        foreach (var task in tasks)
        {
            Console.WriteLine($"{task.To} {task.Status}");
        }

        Console.WriteLine("");

    }

    static Task<EmailStatus> SendEmailAsync(RecipientInfo recipient, string subject, string body, string sender, string senderEmail, string attachmentFolder, SmtpSettings smtpSettings)
    {
        var tcs = new TaskCompletionSource<EmailStatus>();

        // SMTP client settings
        var smtp = new SmtpClient(smtpSettings.Host, smtpSettings.Port)
        {
            EnableSsl = smtpSettings.UseSsl,
            Credentials = new NetworkCredential(smtpSettings.UserName, smtpSettings.AppPassword),
            UseDefaultCredentials = false
        };

        // Prepare the email message
        var email = new MailMessage()
        {
            From = new MailAddress(senderEmail, sender),
            Subject = subject,
            Body = body
        };

        // Main recipient address
        email.To.Add(new MailAddress(recipient.Emails.First(), recipient.Name));
        // Add rest of the email to the CC list
        var ccList = recipient.Emails.Skip(1);
        foreach (var recipientEmail in ccList)
        {
            email.CC.Add(new MailAddress(recipientEmail, recipient.Name));
        }

        // Add all attachments to the email
        foreach (var attachment in recipient.Attachments)
        {
            var filename = Path.Combine(attachmentFolder, attachment);
            Attachment data = new Attachment(filename, System.Net.Mime.MediaTypeNames.Application.Octet);
            System.Net.Mime.ContentDisposition disposition = data.ContentDisposition;
            disposition.CreationDate = File.GetCreationTime(attachment);
            disposition.ModificationDate = File.GetLastWriteTime(attachment);
            disposition.ReadDate = File.GetLastAccessTime(attachment);

            // Add the file attachment to this email message.
            email.Attachments.Add(data);
        }

        // Event handler to check the send status of the email once it is sent to Gmail.
        // Set the status results and dispose the resources
        smtp.SendCompleted += (s, e) =>
        {
            if (e.UserState is not TaskCompletionSource<EmailStatus> state)
            {
                Console.WriteLine("Invalid state");
            }

            if (e.Cancelled)
            {
                tcs.SetResult(new EmailStatus(email.To.First().Address, "Cancelled"));
            }
            else if (e.Error != null)
            {
                tcs.SetResult(new EmailStatus(email.To.First().Address, $"Failed: {e.Error.Message}"));
            }
            else
            {
                tcs.SetResult(new EmailStatus(email.To.First().DisplayName, "has been sent."));
            }

            if (s is SmtpClient smtpClient)
                smtpClient.Dispose();

            email.Dispose();
        };

        try
        {
            // Sending the email to Gmail
            smtp.SendAsync(email, tcs);
        }
        catch (Exception ex)
        {
            // Dispose all in case there was an issue
            smtp.Dispose();
            email.Dispose();
            tcs.SetResult(new EmailStatus(email.To.First().Address, $"Exception: {ex.Message}"));
        }
        // Return the status of the email send
        return tcs.Task;
    }
}


record EmailStatus(string To, string Status);

internal class RecipientInfo
{
    public required string Name { get; set; }
    public required List<string> Emails { get; set; }
    public required List<string> Attachments { get; set; }
}

internal class SmtpSettings
{
    public string Host { get; set; } = "smtp.gmail.com";
    public int Port { get; set; } = 587;
    public bool UseSsl { get; set; } = true;
    public string UserName { get; set; } = "";
    public string AppPassword { get; set; } = "";
}
