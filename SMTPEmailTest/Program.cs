using System.Reflection;
using Outlook = Microsoft.Office.Interop.Outlook;

try
{
    // Create an Outlook application instance
    Outlook.Application outlookApp = new Outlook.Application();

    // Get the MAPI namespace
    Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

    // Log on to the default profile
    outlookNamespace.Logon(Missing.Value, Missing.Value, false, true);

    // Get the Inbox folder
    Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

    var junkBox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);


    // Get all mail items in the Inbox
    Outlook.Items mailItems = inbox.Items;

    // Target email address
    string targetSender = "rafiqrahman@gmail.com";
    string downloadFolder = @"C:\EmailAttachments\";

    // Ensure the download folder exists
    if (!Directory.Exists(downloadFolder))
    {
        Directory.CreateDirectory(downloadFolder);
    }

    Console.WriteLine($"Checking emails from: {targetSender}\n");

    // Loop through mail items
    foreach (object item in mailItems)
    {
        if (item is Outlook.MailItem mail)
        {
            if (mail.SenderEmailAddress.Equals(targetSender, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"Subject: {mail.Subject}");
                Console.WriteLine($"Received: {mail.ReceivedTime}");

                // Check if the email has attachments
                if (mail.Attachments.Count > 0)
                {
                    Console.WriteLine("Attachments found. Downloading...");

                    // Loop through each attachment and save it
                    foreach (Outlook.Attachment attachment in mail.Attachments)
                    {
                        string filePath = Path.Combine(downloadFolder, attachment.FileName);
                        attachment.SaveAsFile(filePath);
                        Console.WriteLine($"Attachment saved: {filePath}");
                    }
                }
                else
                {
                    Console.WriteLine("No attachments found.");
                }

                Console.WriteLine("----------------------------------");
            }
        }
    }

    Console.WriteLine("Finished checking emails.");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
}