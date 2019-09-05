using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

using MailKit;
using MailKit.Net.Imap;
using MimeKit;
using MailKit.Net.Smtp;
using MailKit.Security;
using MailKit.Search;

namespace EmailAgent
{
    class Program
    {
        private static int Classify(string subject)
        {
            if ((subject.Contains("Facebook")) || (subject.Contains("Instagram")) || (subject.Contains("Twitter")))
                return 1;
            else if ((subject.Contains("Amazon")) || (subject.Contains("Banamex")) || (subject.Contains("Costco")))
                return 2;
            else return 3;
        }       

        private static void ThreadProc()
        {
            MimeMessage message;
            BodyBuilder bodyBuilder;
            Random random;
            int i = 0, seed;
            //Email Subjects
            List<string> subjects = new List<string>()
                {
                    "MailKit - Marco, tienes una nueva solicitud de amistad en Facebook",
                    "MailKit - A Alejandro le gusta tu foto en Instagram",
                    "MailKit - Twitter - @mlunac08 ha retwitteado tu twit!",

                    "MailKit - Amazon - Nuevos artículos que pueden interesarte",
                    "MailKit - Marco, recibe 6 MSI con tu tarjeta Banamex",
                    "MailKit - Venta de Liquidación en Costco",

                    "MailKit - Canvas Notification: Daily Notification",
                    "MailKit - Starbucks Rewards",
                    "MailKit - Leer está de moda"
                };

            while (true)
            {
                try
                {
                    // Random Sleeping Time For Email
                    random = new Random(i);
                    // Seed must be multiples of 1000 (miliseconds)
                    seed = random.Next(0, 8);
                    Thread.Sleep((seed * 1000));

                    // Initialize Message
                    message = new MimeMessage();
                    bodyBuilder = new BodyBuilder();

                    // From
                    message.From.Add(new MailboxAddress("EmailAgentProgram", "email"));
                    // To
                    message.To.Add(new MailboxAddress("name", "email"));
                    // Select Random Subject From List
                    message.Subject = subjects[seed];
                    bodyBuilder.HtmlBody = " ";
                    message.Body = bodyBuilder.ToMessageBody();

                    #region SMTP Connection
                    var _client = new SmtpClient();

                    //Configure SMTP
                    _client.ServerCertificateValidationCallback = (s, c, h, e) => true;
                    _client.Connect("smtp-mail.outlook.com", 587, SecureSocketOptions.StartTls);
                    _client.Authenticate("email", "password");
                    _client.Send(message);
                    _client.Disconnect(true);
                    #endregion

                    Console.WriteLine("     -----Email Send...");
                }
                catch (Exception e)
                {
                    Console.WriteLine("     -----EmailAgent - Error {0}", e.Message);
                }                            

                // Change Random's Seed
                i++;                
            }
        }
        
        static void Main(string[] args)
        {
            var client = new ImapClient();
            var cancel = new CancellationTokenSource();
            int inboxEmail, classify, aux;
            IMailFolder matchFolder;
            string folder;

            #region Environment Setup

            #region IMAP Connection
            try
            {
                Console.WriteLine("Agent - Connecting to IMAP...");

                client.Connect("imap-mail.outlook.com", 993, true, cancel.Token);
                // If you want to disable an authentication mechanism,
                // you can do so by removing the mechanism like this:
                client.AuthenticationMechanisms.Remove("XOAUTH");
                client.Authenticate("email", "password", cancel.Token);

                Console.WriteLine("Agent - Connection Successful to email");
            }
            catch (Exception e)
            {
                Console.WriteLine("Program - Client Connection Error: {0}", e.Message);
            }
            #endregion

            // Check Inbox Status
            var inbox = client.Inbox;
            inbox.Open(FolderAccess.ReadOnly, cancel.Token);
            inboxEmail = inbox.Count;
            Console.WriteLine("Agent - Emails in inbox: {0}", inboxEmail);
            inbox.Close();

            #endregion

            // Email Sender Thread
            Thread t = new Thread(new ThreadStart(ThreadProc));
            t.Start();

            #region Agent Behavior / Step
            while (true)
            {
                Task continuation = Task.Run(() => 
                {
                    Thread.Sleep(5000);

                    try
                    {
                        Console.WriteLine("Agent - Retrieving emails from inbox...");

                        // Check if there are new emails
                        inbox.Open(FolderAccess.ReadWrite, cancel.Token);
                        if ((aux = inbox.Count) > inboxEmail)
                        {                            
                            //Get Emails
                            var query = SearchQuery.SubjectContains("MailKit");
                            var uids = client.Inbox.Search(query);
                            var items = client.Inbox.Fetch(uids, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                            Console.WriteLine("Agent - New email(s) in inbox: {0}", (inbox.Count - inboxEmail));

                            // Read new emails                            
                            foreach (var item in items)
                            {
                                var message = inbox.GetMessage(item.UniqueId);
                                Console.WriteLine("Agent - Subject: {0}", message.Subject);

                                // Classify email
                                classify = Classify(message.Subject.ToString());
                                try
                                {
                                    matchFolder = null;
                                    if (classify == 1)
                                    {
                                        matchFolder = client.GetFolder("_Social");
                                        folder = "_Social";
                                    }
                                    else if (classify == 2)
                                    {
                                        matchFolder = client.GetFolder("_Publicity");
                                        folder = "_Publicity";
                                    }
                                    else
                                    {
                                        matchFolder = client.GetFolder("_Other");
                                        folder = "_Other";
                                    }
                                    // Move to Folder
                                    if (matchFolder != null)
                                    {
                                        inbox.CopyTo(item.UniqueId, matchFolder);
                                        // Add Delete Flag in Server
                                        inbox.AddFlags(item.UniqueId, MessageFlags.Deleted, true);
                                        // Purge Folder for all deleted items
                                        inbox.Expunge();
                                    }                                        

                                    Console.WriteLine("Agent - Message Moved to {0}", folder);
                                }
                                catch(Exception e)
                                {
                                    Console.WriteLine("Agent - Couldn't move Email: {0}", e.Message);
                                }                               
                                
                                inboxEmail = inbox.Count;
                            }
                        }
                        else Console.WriteLine("Agent - No new emails...");

                        inbox.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Program - Couldn't retrieve messages: {0}", e.Message);
                    }
                });
                continuation.Wait();                
            }
            #endregion
        }
    }
}    