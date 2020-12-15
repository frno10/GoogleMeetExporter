using System;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleCalendar
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            Events events = new Events();
            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 100;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            int index = 1;
            var tomorrow = DateTime.Now.AddDays(1).Date;

            do
            {
                if(events.Items?.Count > 0)
                {
                    request.TimeMin = events.Items.Last().Start.DateTime.Value.Date;
                }
                else 
                {
                    request.TimeMin = new DateTime(2019, 2, 1);
                }

                // List events.
                events = request.Execute();
                Console.WriteLine("Events with attachments :");
                if (events.Items != null && events.Items.Count > 0)
                {
                    var matchingItems = events.Items.Where(x => x.Attachments != null 
                        && x.Attachments.Any() 
                        && x.Start.DateTime.HasValue 
                        && x.Start.DateTime.Value < tomorrow);

                    foreach (var eventItem in matchingItems)
                    {
                        string when = eventItem.Start.DateTime.ToString();
                        if (String.IsNullOrEmpty(when))
                        {
                            when = eventItem.Start.Date;
                        }
                        var attachmentIndex = 0;

                        foreach(var attachment in eventItem.Attachments)
                        {
                            var data = string.Format("{1}, {2}, {3}, {4}, {5}, {6}", 
                                index, 
                                attachmentIndex == 0 ? eventItem.Summary.Safeproof() : "", 
                                attachmentIndex == 0 ? when.Safeproof() : "",
                                eventItem.Attachments[attachmentIndex].Title.Safeproof(),
                                eventItem.Attachments[attachmentIndex].MimeType.Safeproof(),
                                eventItem.Attachments[attachmentIndex].FileUrl.Safeproof(),
                                eventItem.Description.Safeproof());
                            Console.WriteLine(data);
                            Console.WriteLine();
                            File.AppendAllLines("calendar-recordings.csv",
                                new [] { data });
                            attachmentIndex++;
                        }
                        index++;
                    }
                }
                else
                {
                    Console.WriteLine("No upcoming events found.");
                }

            }while(events.Items.Count == 100 && events.Items.Last().Start.DateTime.Value.Date < DateTime.Now.Date);

            Console.Read();
        }
    }

    public static class ClassExtensions
    {
        public static string Safeproof(this string info)
        {
            return info?.Replace(",", ";")
                .Replace(Environment.NewLine, " ")
                .Replace("\n", " ")
                .Replace("\n\r", " ")
                .Replace("\r\n", " ")
                .Replace("\r", " ")
                 ?? "";
        }
    }
}