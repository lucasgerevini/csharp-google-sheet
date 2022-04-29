using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;

namespace csharp.google.sheet
{
    public class GoogleSheetsHelper
    {
        public SheetsService Service { get; set; }
        const string APPLICATION_NAME = "MASTERCARD";
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public GoogleSheetsHelper()
        {
            InitializeService();
        }
        private void InitializeService()
        {
            var credential = GetCredentialsFromFile();
            Service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = APPLICATION_NAME
            });
        }
        private GoogleCredential GetCredentialsFromFile()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("mastercard-344422-60f4220a84b5.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }
            return credential;
        }
    }
}
//private string[] _scopes = { SheetsService.Scope.Spreadsheets }; // Change this if you're accessing Drive or Docs
        //private string _applicationName = "My Application Name from Google API Project ";
        //private string _spreadsheetId = "xdMsqBc3wblahblahblahblahkeygoeshere";
        //private SheetsService _sheetsService;

        //private void ConnectToGoogle()
        //{
        //    GoogleCredential credential;

        //    // Put your credentials json file in the root of the solution and make sure copy to output dir property is set to always copy 
        //    using (var stream = new FileStream(Path.Combine(HttpRuntime.BinDirectory, "credentials.json"),
        //        FileMode.Open, FileAccess.Read))
        //    {
        //        credential = GoogleCredential.FromStream(stream).CreateScoped(_scopes);
        //    }

        //    // Create Google Sheets API service.
        //    _sheetsService = new SheetsService(new BaseClientService.Initializer()
        //    {
        //        HttpClientInitializer = credential,
        //        ApplicationName = _applicationName
        //    });
        //}

        //// Pass in your data as a list of a list (2-D lists are equivalent to the 2-D spreadsheet structure)
        //public string UpdateData(List<IList<object>> data)
        //{
        //    String range = "My Tab Name!A1:Y";
        //    string valueInputOption = "USER_ENTERED";

        //    // The new values to apply to the spreadsheet.
        //    List<Data.ValueRange> updateData = new List<Data.ValueRange>();
        //    var dataValueRange = new Data.ValueRange();
        //    dataValueRange.Range = range;
        //    dataValueRange.Values = data;
        //    updateData.Add(dataValueRange);

        //    Data.BatchUpdateValuesRequest requestBody = new Data.BatchUpdateValuesRequest();
        //    requestBody.ValueInputOption = valueInputOption;
        //    requestBody.Data = updateData;

        //    var request = _sheetsService.Spreadsheets.Values.BatchUpdate(requestBody, _spreadsheetId);

        //    Data.BatchUpdateValuesResponse response = request.Execute();
        //    // Data.BatchUpdateValuesResponse response = await request.ExecuteAsync(); // For async 

        //    return JsonConvert.SerializeObject(response);
        //}
        // }
        //}
