using System;
using System.Collections.Generic;
using System.Configuration; // Required for ConfigurationManager
using System.IO;
using System.Linq;
using System.ServiceModel.Syndication; // For SyndicationFeed, SyndicationItem
using System.Xml; // For XmlReader
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data; // For Message
using Google.Apis.Services;
using Google.Apis.Util.Store;
using MimeKit; // For MimeMessage, MailboxAddress, TextPart

namespace GmailSender
{
    // Define a class to represent an email recipient
    public class EmailRecipient
    {
        public string Name { get; set; }
        public string EmailAddress { get; set; }

        public EmailRecipient(string name, string emailAddress)
        {
            Name = name;
            EmailAddress = emailAddress;
        }

        public override string ToString()
        {
            return $"Name: {Name}, Email: {EmailAddress}";
        }
    }

    /// <summary>
    /// RSS to Gmail new auth since simple SMTP no longer works with Gmail
    /// Now supports sending to multiple recipients configured in app.config
    /// </summary>
    class Program
    {
        /// <summary>
        /// Load app config into variables
        /// </summary>
        public static string EmailFromName = ConfigurationManager.AppSettings["EmailFromName"] ?? "Default From Name";
        public static string EmailFromAddress = ConfigurationManager.AppSettings["EmailFromAddress"] ?? "defaultfrom@example.com";
        public static string RSSFeedAddress = ConfigurationManager.AppSettings["RSSFeedAddress"] ?? "https://defaultrss.com/feed";
        // EmailToName and EmailToAddress are now handled by a list of EmailRecipient objects

        /// <summary>
        /// Parses the 'EmailRecipients' appSetting into a list of EmailRecipient objects.
        /// Expected format in app.config: <add key="EmailRecipients" value="Name1:email1@example.com;Name2:email2@example.com"/>
        /// </summary>
        /// <returns>A list of EmailRecipient objects.</returns>
        private static List<EmailRecipient> GetEmailRecipientsFromConfig()
        {
            List<EmailRecipient> emailToList = new List<EmailRecipient>();
            try
            {
                string? recipientsValue = ConfigurationManager.AppSettings["EmailRecipients"];

                if (!string.IsNullOrEmpty(recipientsValue))
                {
                    // Split by semicolon to get individual "Name:EmailAddress" pairs
                    string[] recipientPairs = recipientsValue.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string pair in recipientPairs)
                    {
                        // Split each pair by colon to get Name and EmailAddress
                        string[] parts = pair.Split(new[] { ':' }, 2); // Split into max 2 parts
                        if (parts.Length == 2)
                        {
                            string name = parts[0].Trim();
                            string email = parts[1].Trim();
                            emailToList.Add(new EmailRecipient(name, email));
                        }
                        else
                        {
                            Console.WriteLine($"Warning: Could not parse recipient entry '{pair}'. Expected 'Name:EmailAddress' format.");
                            LogError($"Could not parse recipient entry '{pair}'. Expected 'Name:EmailAddress' format.", "Configuration Parsing");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Warning: 'EmailRecipients' key is missing or empty in app.config. No emails will be sent to recipients.");
                    LogError("'EmailRecipients' key is missing or empty in app.config.", "Configuration Parsing");
                }
            }
            catch (ConfigurationErrorsException ex)
            {
                Console.WriteLine($"Error reading 'EmailRecipients' from app settings: {ex.Message}");
                LogError($"Error reading 'EmailRecipients' from app settings: {ex.Message}", "Configuration Parsing");
            }
            return emailToList;
        }


        static void Main(string[] args)
        {
            try
            {
                List<EmailRecipient> recipients = GetEmailRecipientsFromConfig();

                if (recipients == null || !recipients.Any())
                {
                    Console.WriteLine("No valid email recipients configured. Exiting.");
                    LogError("No valid email recipients configured. Program will exit.", "Main Execution");
                    return; // Exit if no recipients
                }

                // OAuth2 credentials
                string[] scopes = { GmailService.Scope.GmailSend };
                UserCredential credential;

                // Ensure credentials.json is in the output directory or provide the correct path
                string credentialsPath = "credentials.json";
                if (!File.Exists(credentialsPath))
                {
                    Console.WriteLine($"Error: {credentialsPath} not found. Please ensure it's in the application directory.");
                    LogError($"{credentialsPath} not found.", "OAuth Setup");
                    return;
                }

                using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json"; // Stores the user's access and refresh tokens.
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        scopes,
                        "user", // A unique identifier for the user.
                        System.Threading.CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + Path.GetFullPath(credPath));
                }

                // Create Gmail API service
                var gmailService = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "RSS Feed Gmail Sender",
                });

                using (XmlReader reader = XmlReader.Create(RSSFeedAddress))
                {
                    SyndicationFeed feed = SyndicationFeed.Load(reader);
                    bool anyEmailSentThisRun = false;

                    foreach (SyndicationItem item in feed.Items)
                    {
                        string itemUrl = item.Links[0].Uri.ToString();
                        if (!ValidateDuplicate(itemUrl))
                        {
                            Console.WriteLine($"Processing new RSS item: {item.Title.Text}");
                            string emailSubject = item.Title.Text;
                            string emailBodyHtml = "<html><body>";

                            if (item.Summary != null)
                            {
                                emailBodyHtml += $"<p>{item.Summary.Text}</p>";
                            }
                            // Use item.Id as the link if it's a permalink, otherwise itemUrl
                            string readMoreLink = item.Id ?? itemUrl;
                            emailBodyHtml += $"<p><a href=\"{readMoreLink}\">Read More</a></p>";
                            emailBodyHtml += "</body></html>";

                            // Send this item to all configured recipients
                            foreach (EmailRecipient recipient in recipients)
                            {
                                try
                                {
                                    var emailMessage = new MimeMessage();
                                    emailMessage.From.Add(new MailboxAddress(EmailFromName, EmailFromAddress));
                                    emailMessage.To.Add(new MailboxAddress(recipient.Name, recipient.EmailAddress));
                                    emailMessage.Subject = emailSubject;
                                    emailMessage.Body = new TextPart("html") { Text = emailBodyHtml };

                                    // Convert email to base64url
                                    var rawMessage = emailMessage.ToString();
                                    byte[] bytes = System.Text.Encoding.UTF8.GetBytes(rawMessage);
                                    string encodedEmail = Convert.ToBase64String(bytes)
                                        .Replace('+', '-')
                                        .Replace('/', '_')
                                        .Replace("=", "");

                                    var message = new Message { Raw = encodedEmail };
                                    var request = gmailService.Users.Messages.Send(message, "me"); // "me" indicates the authenticated user
                                    request.Execute();
                                    Console.WriteLine($"Email sent successfully to: {recipient.Name} <{recipient.EmailAddress}> for item: {emailSubject}");
                                    anyEmailSentThisRun = true;
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Failed to send email to {recipient.EmailAddress} for item '{emailSubject}'. Error: {ex.Message}");
                                    LogError($"Failed to send email to {recipient.EmailAddress} for item '{emailSubject}'. Error: {ex.Message}", "Email Sending");
                                }
                            }
                            LogURL(itemUrl); // Log URL after attempting to send to all recipients
                        }
                    }

                    if (anyEmailSentThisRun)
                    {
                        Console.WriteLine("RSS check complete and emails sent.");
                    }
                    else
                    {
                        Console.WriteLine("RSS check complete. No new items to send.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Program encountered an unhandled error:");
                Console.WriteLine(ex.ToString()); // Using ex.ToString() for more details including stack trace
                LogError($"Unhandled Exception: {ex.Message}, StackTrace: {ex.StackTrace}", ex.Source ?? "Unknown Source");
            }
        }

        private static void LogURL(string strMainPostURL)
        {
            try
            {
                // Using statement ensures an IDisposable object is correctly disposed.
                using (FileStream fs = new FileStream("URL_Log.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (StreamWriter log = new StreamWriter(fs))
                {
                    log.WriteLine(strMainPostURL);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error logging URL '{strMainPostURL}': {ex.Message}");
                // Optionally log this logging error to the main error log
                LogError($"Error in LogURL for '{strMainPostURL}': {ex.Message}", "Logging");
            }
        }

        private static bool ValidateDuplicate(string strMainPostURL)
        {
            string logFilePath = "URL_Log.txt";
            if (!File.Exists(logFilePath))
            {
                return false;
            }
            try
            {
                // Read all lines to avoid holding the file open for too long if it's large
                // For very large files, line-by-line reading might be better, but Contains check on full text is fine for moderate sizes.
                string text = File.ReadAllText(logFilePath);
                return text.Contains(strMainPostURL);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating duplicate URL '{strMainPostURL}': {ex.Message}");
                LogError($"Error in ValidateDuplicate for '{strMainPostURL}': {ex.Message}", "Validation");
                return false; // If error, assume not a duplicate to allow processing, or handle as critical error
            }
        }

        private static void LogError(string message, string source)
        {
            try
            {
                using (FileStream fs = new FileStream("ErrorLog.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                using (StreamWriter log = new StreamWriter(fs))
                {
                    log.WriteLine($"{DateTime.Now} Source: {source} Error: {message}");
                }
            }
            catch (Exception ex)
            {
                // If logging itself fails, write to console.
                Console.WriteLine($"CRITICAL: Could not write to ErrorLog.txt. Original error from '{source}': '{message}'. Logging error: {ex.Message}");
            }
        }
    }
}
