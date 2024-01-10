using CoachTimetableEditorApp.GoogleSheetlManager;
using Quartz;

namespace CoachTimetableEditorApp.QuatzManager
{
    public class UpdateGoogleSheetJob : IJob
    {
        private readonly IGoogleSheetHandler _googleSheetHandler;

        public UpdateGoogleSheetJob(IGoogleSheetHandler googleSheetHandler)
        {
            _googleSheetHandler = googleSheetHandler;
        }
        public async Task Execute(IJobExecutionContext context)
        {
            await _googleSheetHandler.UpdateSheetCellValueAsync();
        }
    }
}
