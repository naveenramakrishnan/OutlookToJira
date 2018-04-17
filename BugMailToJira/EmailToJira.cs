using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace BugMailToJira
{
    class EmailToJira
    {
        static List<string> MailSubjectList = new List<string>();
        static List<string> MailBodyList = new List<string>();
        static List<string> NewEmailList = new List<string>();
        static string MailSubject;
        static string MailBody;
        static string TodayDate;
        static string ListName;
        static int EmailCount = 0;

        public static void Main(string[] args)
        {
            ReadOutlook();
        }

        public static void ReadOutlook()
        {
            Outlook._Application olApp = new Outlook.ApplicationClass();
            Outlook._NameSpace olNS = olApp.GetNamespace("MAPI");
            olNS.Logon("@OutlookEmail", "@OutlookPassword", false, false);
            Outlook.MAPIFolder oFolder = olNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            TodayDate = DateTime.Now.ToString("MM/dd/yyyy");
            Outlook.Items oItems = oFolder.Items.Restrict("[ReceivedTime] >= '"+TodayDate+"'");
            //Outlook.Items oItems = oFolder.Items.Restrict("[UnRead] = true");

            for (int i = 1; i <= oItems.Count; i++)
            {
                Microsoft.Office.Interop.Outlook.MailItem oMsg = (Microsoft.Office.Interop.Outlook.MailItem)oItems[i];
                MailSubject = oMsg.Subject.ToString();
                MailBody = System.Web.HttpUtility.JavaScriptStringEncode(oMsg.Body);

                if (MailSubject.StartsWith("RE:") || MailSubject.StartsWith("FW:") || MailSubject.StartsWith("Re:") || MailSubject.StartsWith("Fw:") || MailSubject.StartsWith("Fwd:"))
                {
                    // Do not fetch mail with contains word above (reply and forward type)
                }
                else
                {
                    EmailCount++;
                    NewEmailList.Add(MailSubject);
                    Console.WriteLine(MailSubject + MailBody);
                    CreatJiraIssue(MailSubject, MailBody);
                }
            }
            ListName = string.Join("\n", NewEmailList.ToArray());
            Console.WriteLine(EmailCount);
            Microsoft.Office.Interop.Outlook.MailItem oMsgSend = (Outlook.MailItem)olApp.CreateItem(Outlook.OlItemType.olMailItem);
            oMsgSend.To = "######Recieved_Email######";
            oMsgSend.To = "######Sender_Email######";
            oMsgSend.Subject = "Summary Auto Email Fetching " + TodayDate;
            oMsgSend.Body = "All New Email count: " + EmailCount + "\n\nFetching complete: " + EmailCount + "\n\nList Email names: \n\n" + ListName;
            oMsgSend.Save();
            oMsgSend.Send();
        }

        public static void CreatJiraIssue(string title, string body)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            string postUrl = "######Jira_Url######";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(postUrl);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            httpWebRequest.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes("Username:Password"));

            title = System.Web.HttpUtility.JavaScriptStringEncode(title);
            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                try
                {
                   string json = "{\"fields\":{\"project\":{\"key\":\"STAR\"},\"issuetype\":{\"id\":\"1\"},\"summary\":\"" + "[BUGS@] " + title + "\",\"description\":\"" + body + "\",\"environment\":\"BUGS\"}}";

                    streamWriter.Write(json);
                    streamWriter.Flush();
                    streamWriter.Close();

                    HttpWebResponse httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                        Console.WriteLine("Item Submitted Successfully");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

    }
}
