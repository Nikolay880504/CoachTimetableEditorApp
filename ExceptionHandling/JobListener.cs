using CoachTimetableEditorApp.QuatzManager;
using Quartz;

namespace CoachTimetableEditorApp.ExceptionHandling
{
    public class JobListener : IJobListener
    {
        public string Name => "JobListener";
        private readonly ILogger<UpdateGoogleSheetJob> _logger;
        public JobListener(ILogger<UpdateGoogleSheetJob> logger)
        {
            _logger = logger;
        }

        public Task JobExecutionVetoed(IJobExecutionContext context, CancellationToken cancellationToken = default)
        {
            throw new NotImplementedException();
        }

        public Task JobToBeExecuted(IJobExecutionContext context, CancellationToken cancellationToken = default)
        {
            _logger.LogInformation($"Job {context.JobDetail.Key} starts execution");

            return Task.CompletedTask;
        }

        public Task JobWasExecuted(IJobExecutionContext context, JobExecutionException? jobException, CancellationToken cancellationToken = default)
        {
            if (jobException != null)
            {
                _logger.LogError(jobException, $"Error executing job {context.JobDetail.Key}");
            }
            return Task.CompletedTask;
        }
    }
}
