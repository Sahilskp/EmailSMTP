using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using EASendMail;
using Newtonsoft.Json;

namespace EmailSMTP
{
    class Program
    {
        static void Main(string[] args)
        {
            dynamic config;
                using(var reader = new StreamReader(@"config.json"))
            {
                String stuff = reader.ReadToEnd();
                config = JsonConvert.DeserializeObject(stuff);
            } 



            List<string> StudentNames = new List<string>();
            List<string> Emails = new List<string>();
            using(var reader = new StreamReader($"{config.Parent_info}"))
                {
                    
                    while (!reader.EndOfStream)
                    {
                        String line = reader.ReadLine();
                        String[] values = line.Split(',');
                        //Console.WriteLine(values[6]);
                        StudentNames.Add(values[7]);
                        Emails.Add(values[6]);
                    }
                }


                
            Console.WriteLine(config.Folder1);
            Emails.RemoveAt(0);
            StudentNames.RemoveAt(0);

            String[] Emails_array = Emails.ToArray();
            String[] StudentNames_array = StudentNames.ToArray();
            //String[] StudentNames = {"Aaliya"};

            for(int x = 0; x<Emails_array.Length; x++)
            {
                try
                {
                    SmtpMail oMail = new SmtpMail("TryIt");

                    // Your email address
                    oMail.From = "sahil.shafeeque@outlook.com";

                    // Set recipient email address
                    if(Emails_array[x].Contains(';'))
                    {
                       String[] split_email= Emails_array[x].Split(';');
                       oMail.To = split_email[0];
                       oMail.Cc = split_email[1];
                    }else
                    {
                    oMail.To = Emails_array[x];
                    }
                    

                    // Set email subject
                    oMail.Subject = config.Subject;

                    // Set email body
                    oMail.TextBody = config.Body;

                    //Adding Attachment
                    oMail.AddAttachment($"{config.Folder1}\\{StudentNames[x]}.pdf");
                    oMail.AddAttachment($"{config.Folder2}\\{StudentNames[x]}.pdf");
                    oMail.AddAttachment($"{config.Folder3}\\{StudentNames[x]}.pdf");

                    // Hotmail/Outlook SMTP server address
                    SmtpServer oServer = new SmtpServer("smtp.live.com");

                    // If your account is office 365, please change to Office 365 SMTP server
                    // SmtpServer oServer = new SmtpServer("smtp.office365.com");

                    // User authentication should use your
                    // email address as the user name.
                    oServer.User = "sahil.shafeeque@outlook.com";
                    oServer.Password = "$ahil450";

                    // use 587 TLS port
                    oServer.Port = 587;

                    // detect SSL/TLS connection automatically
                    oServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

                    Console.WriteLine("start to send email over TLS...");

                    SmtpClient oSmtp = new SmtpClient();
                    oSmtp.SendMail(oServer, oMail);

                    Console.WriteLine("email was sent successfully!");
            }catch(Exception ep)
                {
                    Console.WriteLine("failed to send email with the following error:");
                    Console.WriteLine(ep.Message);

                }
            }
            
        }
    }
}
