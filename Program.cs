using CoachTimetableEditorApp.Authentication;
using CoachTimetableEditorApp.DataBase;
using CoachTimetableEditorApp.ExceptionHandling;
using CoachTimetableEditorApp.GoogleDriveExcelManager;
using CoachTimetableEditorApp.GoogleSheetlManager;
using CoachTimetableEditorApp.QuatzManager;
using Quartz;
using Serilog;


namespace CoachTimetableEditorApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var dbFilePath = "jobs.sqlite";
            var connectionString = $"Data Source={dbFilePath};Version=3;";
            var logFilePath = Path.Combine("Logs", "log.txt");
   
            var builder = WebApplication.CreateBuilder(args);

            builder.Services.AddScoped<IGoogleAuthentication, GoogleAuthentication>();
            builder.Services.AddScoped<IGoogleSheetHandler, GoogleSheetHandler>();
            builder.Logging.AddSerilog();

            Log.Logger = new LoggerConfiguration()
                .WriteTo.File(logFilePath, rollingInterval: RollingInterval.Month)
                .CreateLogger();

            if (!File.Exists(dbFilePath))
            {
                new QuartzTableCreator().CreateQuartzTables(connectionString);
            }
            // for test "0 0/59 * * * ?" "0 0 0 1 * ? *"
            builder.Services.AddQuartz(configure =>
            {
                configure.UsePersistentStore(c =>
                {
                    c.UseSQLite(@"URI=file:jobs.sqlite;Version=3;");
                    c.UseNewtonsoftJsonSerializer();
                });
                var jobKey = new JobKey(nameof(UpdateGoogleSheetJob));
                configure
                    .AddJob<UpdateGoogleSheetJob>(jobKey); 
                configure.AddTrigger(trigger => trigger.ForJob(jobKey).WithCronSchedule("0 0 0 1 * ? *",
                 x => x.WithMisfireHandlingInstructionFireAndProceed()));
                configure.AddJobListener<JobListener>();

            });

            builder.Services.AddQuartzHostedService(options =>
            {
                options.WaitForJobsToComplete = true;
            });

            var app = builder.Build();

            app.Run();
        }
    }
}
