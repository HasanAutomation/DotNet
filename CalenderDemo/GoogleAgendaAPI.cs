using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Drive.v3;
using Google.Apis.Upload;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace CalenderDemo
{
    class GoogleAgendaAPI
    {
        static string[] Scopes = { CalendarService.Scope.Calendar, CalendarService.Scope.CalendarEvents, DriveService.Scope.DriveFile };
        static string ApplicationName = "Google Calendar API to create event";
        private static string _fileName = "index.html";
        private static string _filePath = @"D:\WebDevelopment\Project\index.html";
       // private static string _filePath = @"D:\custom.jpg";

        static void Main(string[] args)
        {
            new GoogleAgendaAPI().CreateEvent("Holiday", "Kolkata", "No work today", "hasan.ali@bassetti-group.com");

        }

        public void CreateEvent(string summary, string location, string description, string email)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("Calender_credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            //Create calender service api
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });


            // Create Drive API service.
            var drive_service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            //calling the upload method to upload the file in drive
            string fileId = new GoogleAgendaAPI().uploadFile(drive_service, _fileName,_filePath);

            DateTime start = DateTime.Now;
            DateTime end = start + TimeSpan.FromMinutes(30);

            DateTime initiate = DateTime.Now;
            DateTime ending = start + TimeSpan.FromMinutes(30);
            start = DateTime.SpecifyKind(start, DateTimeKind.Local);
            end = DateTime.SpecifyKind(end, DateTimeKind.Local);


            Event eve = new Event
            {
                Summary = summary,
                Location = location,
                Description = description,
                Start = new EventDateTime()
                {
                    DateTime = DateTime.Parse("2020-12-12T09:00:00-07:00"),
                    TimeZone = "America/Los_Angeles",
                },
                End = new EventDateTime()
                {
                    DateTime = DateTime.Parse("2020-12-13T17:00:00-07:00"),
                    TimeZone = "America/Los_Angeles",
                },
                Recurrence = new String[] { "RRULE:FREQ=WEEKLY;BYDAY=MO" },
                Attendees = new List<EventAttendee>
                 {
                     new EventAttendee { Email = email }
                 },
                Attachments = new List<EventAttachment>
                {
                    new EventAttachment{FileId=fileId}
                }
            };


            String calendarId = "primary";
            eve = service.Events.Insert(eve, calendarId).Execute();



            Console.WriteLine("Event created " + eve.HtmlLink);

        }

        //To upload a file in the drive
        public string uploadFile(DriveService _service, string fileName, string filePath)
        {

            //To create a folder in the drive
            var folderMetaData = new Google.Apis.Drive.v3.Data.File();
            folderMetaData.Name = "CustomFolder";
            folderMetaData.MimeType = "application/vnd.google-apps.folder";

            _service.Files.Create(folderMetaData).Fields = "Id";
            var fileId = _service.Files.Create(folderMetaData).Execute();
            Console.WriteLine(fileId.Id);


            //To create an empty file
            var fileData = new Google.Apis.Drive.v3.Data.File();
            fileData.Name = fileName;
            fileData.MimeType = "application/octet-stream";

            fileData.Parents = new List<string> { fileId.Id };

            FilesResource.CreateMediaUpload request;

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                request = _service.Files.Create(fileData, stream, "application/octet-stream");
                request.Upload();

            }
            Console.WriteLine("FIle id is " + request.ResponseBody.Id);
            return request.ResponseBody.Id;

        }

    }
}
