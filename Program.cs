using System;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;


namespace GoogleSheetsAndCsharp
{
    class Program
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "Legislators";
        static readonly string SpreadsheetId = "1Z4qmu8ErL6RVq3kqEzTbSp8KZWizJGSpfTO0mRB_qjk"; //from URL
        static readonly string sheet = "congress";
        static SheetsService service;
        static void Main(string[] args)
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            ReadEntries();
        }

        static void ReadEntries()
        {
            var range = $"{sheet}!A1:F10";
            var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            var values = response.Values;
            if(values != null && values.Count > 0)
            {
                foreach(var row in values)
                {
                    Console.WriteLine("{0} {1} | {2} | {3}", row[5], row[4], row[3], row[1]);
                }
            }
            else
            {
                Console.WriteLine("No Data found.");
            }

        }
    }
}
