using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

namespace GoogleSheetsAndCsharp
{
    class Program
    {
        // Created following
        // https://www.youtube.com/watch?v=afTiNU6EoA8&feature=youtu.be

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
            CreateEntry();
            UpdateEntry();
            DeleteEntry();
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

        static void CreateEntry()
        {
            var range = $"{sheet}!A:F";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Hello!", "This", "was", "inserted", "via", "C#" };
            valueRange.Values = new List<IList<object>> { objectList };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, range);
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            var appendResponse = appendRequest.Execute();
        }

        static void UpdateEntry()
        {
            var range = $"{sheet}!A2";
            var valueRange = new ValueRange();

            var objectList = new List<object>() { "Updated!" };
            valueRange.Values = new List<IList<object>> { objectList };

            var updateRequest = service.Spreadsheets.Values.Update(valueRange, SpreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            var updateResponse = updateRequest.Execute();

        }

        static void DeleteEntry()
        {
            var range = $"{sheet}!A543:F";
            var requestBody = new ClearValuesRequest();

            var deleteRequest = service.Spreadsheets.Values.Clear(requestBody, SpreadsheetId, range);
            var deleteResponse = deleteRequest.Execute();
        }
    }
}
