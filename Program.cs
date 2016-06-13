using System;
using System.Collections.Generic;

namespace KiepPhotoDownloader
{
    class Program
    {
        private static string user = "username";
        private static string pass = "password";
        private static string destination = "c:\\E-mail";
        private static int days = 7;
        private static string[] extensions = new string[] { ".jpg", ".JPG", ".JPEG", ".jpeg", ".png", ".PNG" };

        static void Main(string[] args)
        {
            try
            {
                if (!System.IO.Directory.Exists(destination))
                {
                    System.IO.Directory.CreateDirectory(destination);
                }

                System.IO.Directory.SetCurrentDirectory(destination);
            }
            catch (Exception)
            {
                System.Console.WriteLine("Folder '" + destination + "' does not exist.");
                return;
            }

            ImapX.ImapClient client;
            try
            {
                client = new ImapX.ImapClient("imap.gmail.com", 993, true);
                client.Connection();
                Console.WriteLine("Connected to Gmail IMAP.");
            }
            catch
            {
                Console.WriteLine("Connection to Gmail IMAP failed.");
                return;
            }

            try
            {
                if (!client.LogIn(user, pass))
                {
                    throw new Exception();
                }
                Console.WriteLine("Logged in as " + user + ".");
            }
            catch
            {
                Console.WriteLine("Failed to login.");
                return;
            }

            ImapX.MessageCollection messages;
            try
            {
                messages = client.Folders["INBOX"].Messages;
                messages.Reverse();
            }
            catch
            {
                Console.WriteLine("Fetching messages failed.");
                return;
            }

            try
            {
                processMessages(messages);
            }
            catch
            {
                Console.WriteLine("Processing messages failed.");
                return;
            }
        }

        private static void processMessages(ImapX.MessageCollection messages)
        {

            int indexMessages = 0;
            foreach (ImapX.Message m in messages)
            {
                indexMessages++;

                try
                {
                    m.Process();
                }
                catch (Exception)
                {
                    Console.WriteLine("Error processing message " + indexMessages + " of " + messages.Count + ".");
                    continue;
                }

                try
                {
                    if (days != 0)
                    {
                        int daysBetween = (System.DateTime.Now - m.Date).Days;
                        int maxDays = (System.DateTime.Now - DateTime.MinValue).Days;
                        if ((daysBetween > days) && (maxDays > daysBetween))
                        {
                            Console.WriteLine("Message " + indexMessages + " of " + messages.Count + " is " + daysBetween + " days old.");
                            Console.WriteLine("Hitting limit of " + days + ", stopping.");
                            return;
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Error getting date of message " + indexMessages + " of " + messages.Count + ".");
                    continue;
                }

                Console.WriteLine("Processing message " + indexMessages + " of " + messages.Count + ".");

                string sender = m.From[0].Address;
                List<ImapX.Attachment> attachments = m.Attachments;

                if (attachments.Count == 0)
                {
                    Console.WriteLine("No attachments found.");
                }
                else
                {
                    Console.WriteLine("Attachments found.");
                    processAttachments(attachments, sender);
                }
            }
        }


        private static void processAttachments(List<ImapX.Attachment> attachments, string sender)
        {
            foreach (ImapX.Attachment attachment in attachments)
            {
                string filename = "";
                try
                {
                    filename = attachment.FileName;
                    if (filename.Length == 0)
                    {
                        throw new Exception();
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine("Error getting attachment filename.");
                    continue;
                }

                Console.WriteLine("Check if " + filename + " is an image.");
                foreach (string extension in extensions)
                {
                    if (filename.EndsWith(extension))
                    {
                        SaveAttachment(attachment, sender);
                    }
                }
            }
        }

        private static void SaveAttachment(ImapX.Attachment attachment, string sender)
        {
            String filename = attachment.FileName;
            Console.WriteLine("Storing attachment '" + filename + "'.");
            
            try
            {
                if (!System.IO.Directory.Exists(sender))
                {
                    System.IO.Directory.CreateDirectory(sender);
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Creating folder for '" + sender + "' failed.");
                return;
            }

            try
            {
                attachment.SaveFile(sender + "\\");
            }
            catch (Exception)
            {
                Console.WriteLine("Saving '" + filename + "' failed.");
                return;
            }

            Console.WriteLine("Attachment '" + filename + "' stored.");
        }

    }
}