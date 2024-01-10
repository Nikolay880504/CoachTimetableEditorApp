using CoachTimetableEditorApp.Authentication;
using CoachTimetableEditorApp.GoogleSheetlManager;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Globalization;
using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;

namespace CoachTimetableEditorApp.GoogleDriveExcelManager
{
    public class GoogleSheetHandler : IGoogleSheetHandler
    {
        private readonly IGoogleAuthentication _googleAuthentication;

        private DriveService _driveService;
        private readonly SheetsService _sheetsService;
        private DateTime _currentDateTime;
        private CultureInfo _cultureInfo;

        private const string ClearSheetRangeString = "!C4:AG56";
        private const string MonthNamesRangeString = "!1:1";
        private const string DayOfWeekRangeString = "!C2:AG2";
        private const string NumbersDayOfMonthRangeString = "!C3:AG3";

        public GoogleSheetHandler(IGoogleAuthentication googleAuthentication)
        {
            _googleAuthentication = googleAuthentication;
            _driveService = _googleAuthentication.GetDriveService();
            _sheetsService = _googleAuthentication.GetSheetsService();
        }

        public async Task UpdateSheetCellValueAsync()
        {
            var fileList = await GetFilesListAsync();
            SetCurrentDate();
            SetCultureInfoForUkraine();

            if (fileList != null && fileList.Files.Count > 0)
            {
                foreach (var file in fileList.Files)
                {
                    var spreadsheetListRequest = _sheetsService.Spreadsheets.Get(file.Id);
                    var spreadsheet = await spreadsheetListRequest.ExecuteAsync();

                    await UpdateFileNameToCurrentMonth(spreadsheet.SpreadsheetId, file.Name);

                    if (spreadsheet.Sheets != null && spreadsheet.Sheets.Count > 0)
                    {
                        foreach (var sheet in spreadsheet.Sheets)
                        {
                            if (sheet.Properties != null)
                            {
                                await ClearSheetRange(spreadsheet.SpreadsheetId, sheet.Properties.Title);                              
                                await UpdateMonthNames(spreadsheet.SpreadsheetId, sheet.Properties.Title);
                                await MapDatesToDaysOfWeek(spreadsheet.SpreadsheetId, sheet.Properties.Title);                              
                            }
                        }
                    }
                }
            }
        }
        private async Task<Google.Apis.Drive.v3.Data.FileList> GetFilesListAsync()
        {
            var request = _driveService.Files.List();
            return await request.ExecuteAsync();
        }
        private async Task ClearSheetRange(string spreadsheetId, string nameSheet)
        {
            var range = $"{nameSheet}{ClearSheetRangeString}";
            var clearRequestBody = new ClearValuesRequest();

            var clearRequest = _sheetsService.Spreadsheets.Values.Clear(clearRequestBody, spreadsheetId, range);
            await clearRequest.ExecuteAsync();
        }     
        private async Task UpdateMonthNames(string spreadsheetId, string nameSheet)
        {
            var curentMonthAndYear = $"{GetCurrentYear()} {_currentDateTime.ToString("MMMM", _cultureInfo).ToUpper()}";
            var range = $"{nameSheet}{MonthNamesRangeString}";

            var valueRange = new ValueRange();
            valueRange.Values = new List<IList<object>>() { new List<object> { curentMonthAndYear } };

            var updateRequest = _sheetsService.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;

            await updateRequest.ExecuteAsync();
        }
        private async Task MapDatesToDaysOfWeek(string spreadsheetId, string nameSheet)
        {
            var rangeForDayOfWeek = $"{nameSheet}{DayOfWeekRangeString}";
            var rangeNumbersDayOfMonth = $"{nameSheet}{NumbersDayOfMonthRangeString}";
            var calendar = _cultureInfo.Calendar;
            var daysInMonth = calendar.GetDaysInMonth(GetCurrentYear(), GetCurrentMonth());

            var listOfNumbersDayOfMonth = new List<object>();
            var listOfDayOfWeek = new List<object>();

            for (int day = 1; day <= 31; day++)
            {
                if (day > daysInMonth)
                {
                    listOfNumbersDayOfMonth.Add(string.Empty);
                    listOfDayOfWeek.Add(string.Empty);
                }
                else
                {
                    listOfNumbersDayOfMonth.Add(day);
                    var specifiedDate = new DateTime(GetCurrentYear(), GetCurrentMonth(), day);
                    listOfDayOfWeek.Add(specifiedDate.ToString("ddd", _cultureInfo));
                }
            }

            var valueRangeForDayOfWeek = new ValueRange { Values = new List<IList<object>> { listOfDayOfWeek } };
            var valueRangeForNumbers = new ValueRange { Values = new List<IList<object>> { listOfNumbersDayOfMonth } };

            var updateRequestForDayOfWeek = _sheetsService.Spreadsheets.Values.Update(valueRangeForDayOfWeek, spreadsheetId, rangeForDayOfWeek);
            updateRequestForDayOfWeek.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;

            var updateRequestForNumbers = _sheetsService.Spreadsheets.Values.Update(valueRangeForNumbers, spreadsheetId, rangeNumbersDayOfMonth);
            updateRequestForNumbers.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;

            await updateRequestForDayOfWeek.ExecuteAsync();
            await updateRequestForNumbers.ExecuteAsync();
        }
        private async Task UpdateFileNameToCurrentMonth(string spreadsheetId, string fileName)
        {
            var ukranianMonth = new string[] { "СІЧЕНЬ","ЛЮТИЙ","БЕРЕЗЕНЬ","КВІТЕНЬ","ТРАВЕНЬ","ЛИПЕНЬ","СІЧЕНЬ","СЕРПЕНЬ",
            "ВЕРЕСЕНЬ","ЖОВТЕНЬ","ЛИСТОПАД","ГРУДЕНЬ"
        };
            var newFileNameParts = new List<string>();
            var splitFileName = fileName.Split(' ');
            foreach (var item in splitFileName)
            {
                if (ukranianMonth.Contains(item))
                {
                    newFileNameParts.Add(_currentDateTime.ToString("MMMM", _cultureInfo).ToUpper());
                }
                else if (int.TryParse(item, out _))
                {
                    newFileNameParts.Add(GetCurrentYear().ToString());
                }
                else
                {
                    newFileNameParts.Add(item);
                }
            }
            var newFileName = string.Join(" ", newFileNameParts);

            await CreatingRequestToChangeTheFileName(newFileName, spreadsheetId);
        }
        private async Task CreatingRequestToChangeTheFileName(string newFileName, string spreadsheetId)
        {
            var request = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
        {
            new Request
            {
                UpdateSpreadsheetProperties = new UpdateSpreadsheetPropertiesRequest
                {
                    Properties = new SpreadsheetProperties
                    {
                        Title = newFileName
                    },
                    Fields = "title"
                }
            }
        }
            };
            await _sheetsService.Spreadsheets.BatchUpdate(request, spreadsheetId).ExecuteAsync();
        }
        private void SetCurrentDate()
        {
            _currentDateTime = DateTime.Now.Date;
        }
        private int GetCurrentMonth()
        {
            return _currentDateTime.Month;
        }

        private int GetCurrentYear()
        {
            return _currentDateTime.Year % 100;
        }

        private void SetCultureInfoForUkraine()
        {
            _cultureInfo = new CultureInfo("uk-UA");
        }
    }
}
