using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace CoachTimetableEditorApp.Authentication
{
    public class GoogleAuthentication : IGoogleAuthentication
    {

        private string _pathToServiceAccountKeyFile = Path.Combine("secretKey", "client_secret.json");
        public DriveService GetDriveService()
        {
            var credential = GoogleCredential.FromFile(_pathToServiceAccountKeyFile)
                    .CreateScoped(DriveService.ScopeConstants.Drive);
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential
            });
            return service;
        }

        public SheetsService GetSheetsService()
        {
            var credential = GoogleCredential.FromFile(_pathToServiceAccountKeyFile)
                .CreateScoped(SheetsService.Scope.Spreadsheets);
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential
            });
            return service;
        }
    }
}