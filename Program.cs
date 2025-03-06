using System.Configuration;
using System.ServiceModel.Syndication;
using System.Xml;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace GmailSender
{
    /// <summary>
    /// RSS to Gmail new auth since simple SMTP no longer works with Gmail
    /// </summary>
    class Program
    {

        /// <summary>
        /// Load app config into variables
        /// </summary>
        public static string EmailFromName = ConfigurationManager.AppSettings["EmailFromName"] ?? "Default From Name";
        public static string EmailFromAddress = ConfigurationManager.AppSettings["EmailFromAddress"] ?? "default@example.com";
        public static string EmailToName = ConfigurationManager.AppSettings["EmailToName"] ?? "Default To Name";
        public static string EmailToAddress = ConfigurationManager.AppSettings["EmailToAddress"] ?? "defaultto@example.com";
        public static string RSSFeedAddress = ConfigurationManager.AppSettings["RSSFeedAddress"] ?? "https://defaultrss.com/feed";

        static void Main(string[] args)
        {
            try
            {
                // OAuth2 credentials
                string[] scopes = { GmailService.Scope.GmailSend };
                UserCredential credential;

                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        scopes,
                        "user",
                        System.Threading.CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Gmail API service
                var gmailService = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Gmail Sender",
                });

                using (XmlReader reader = XmlReader.Create(RSSFeedAddress))
                {
                    SyndicationFeed feed = SyndicationFeed.Load(reader);

                    var email = new MimeKit.MimeMessage();
                    email.From.Add(new MimeKit.MailboxAddress(EmailFromName, EmailFromAddress));
                    email.To.Add(new MimeKit.MailboxAddress(EmailToName, EmailToAddress));
                                        
                    bool UrlToEmail = false;

                    foreach (SyndicationItem item in feed.Items)
                    {
                        string URL = item.Links[0].Uri.ToString();
                        if (!ValidateDuplicate(URL))
                        {
                            UrlToEmail = true;
                            string emailBody = "<html><body>";
                            Console.WriteLine("Processing: " + item.Title.Text);
                            email.Subject = item.Title.Text;

                            if (item.Summary != null)
                            {
                                emailBody += $"<p>{item.Summary.Text}</p>";
                            }

                            emailBody += $"<p><a href=\"{item.Id}\">Read More</a></p>";

                            emailBody += "</body></html>";

                            email.Body = new MimeKit.TextPart("html")
                            {
                                Text = emailBody

                            };

                            // Convert email to base64url
                            var emailContent = email.ToString();
                            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(emailContent);
                            string encodedEmail = System.Convert.ToBase64String(bytes)
                                .Replace('+', '-')
                                .Replace('/', '_')
                                .Replace("=", "");

                            // Send email
                            var message = new Message { Raw = encodedEmail };
                            var request = gmailService.Users.Messages.Send(message, "me");
                            request.Execute();
                            Console.WriteLine("Email sent successfully.");
                            LogURL(URL);
                        }
                    }
                    if (UrlToEmail)
                    {
                        Console.WriteLine("RSS check Complete and emails sent.");
                    }
                    else
                    {
                        Console.WriteLine("RSS check complete. No emails to send.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Program encountered and error:");
                Console.WriteLine(ex.Message);
                LogError(ex.Message, ex.Source ?? "Unknown Source");
            }
        }

        private static void LogURL(string strMainPostURL)
        {
            FileStream fs = new FileStream("URL_Log.txt", FileMode.Append);
            StreamWriter log = new StreamWriter(fs);
            log.WriteLine(strMainPostURL);
            log.Close();
            fs.Close();
        }

        private static bool ValidateDuplicate(string strMainPostURL)
        {
            if (!File.Exists("URL_Log.txt"))
            {
                return false;
            }
            StreamReader streamReader = new StreamReader("URL_Log.txt");
            string text = streamReader.ReadToEnd();
            streamReader.Close();
            if (text.Contains(strMainPostURL))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private static void LogError(string Message, string source)
        {
            try
            {
                FileStream fs = null;
                fs = new FileStream("ErrorLog.txt", FileMode.Append);
                StreamWriter log = new StreamWriter(fs);
                log.WriteLine(DateTime.Now + " Error: " + Message + " Source: " + source);
                log.Close();
                fs.Close();
            }
            catch (Exception)
            {
                Console.WriteLine("Error logging previous error.");
                Console.WriteLine("Make sure the Error log is not open.");
            }
        }
    }
}
