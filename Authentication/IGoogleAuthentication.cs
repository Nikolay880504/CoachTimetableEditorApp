using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;

namespace CoachTimetableEditorApp.Authentication
{
    public interface IGoogleAuthentication
    {
        public DriveService GetDriveService();

        public SheetsService GetSheetsService();
    }
}
