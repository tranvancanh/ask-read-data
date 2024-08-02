using Microsoft.Extensions.Hosting;
using System.Threading.Tasks;
using System.Threading;
using System;
using System.Diagnostics;
using Microsoft.Extensions.Logging;
using System.IO;
using System.Linq;

namespace ask_read_data.Commons
{
    public class StartupTaskService : IHostedService
    {
        private int MostRecentDays = 365; // 365days

        private readonly ILogger<StartupTaskService> _logger;
        public StartupTaskService(ILogger<StartupTaskService> logger)
        {
            _logger = logger;
            _logger.LogInformation("Nlog is started to StartupTask Service");
        }

        public Task StartAsync(CancellationToken cancellationToken)
        {
            // Place code to run when the application starts
            Debug.WriteLine("Startup Task is started");
            RunStartupTask(cancellationToken);

            return Task.CompletedTask;
        }

        private Task RunStartupTask(CancellationToken cancellationToken)
        {
            // Perform your startup task here
            Debug.WriteLine("Startup Task Executing!");
            // Use NLog to log messages
            _logger.LogInformation("ImmedStartupiately Task Executing!");

            // Example: Read configuration, initialize data, etc.
            var rootPath = Directory.GetCurrentDirectory();
            var uploadPath = Path.Combine(rootPath, @"wwwroot\ImportFiles");
            var downloadPath = Path.Combine(rootPath, @"wwwroot\DownloadFiles");
            var logPath = Path.Combine(rootPath, @"wwwroot\Logs");

            try
            {
                var dirUpFiles = new DirectoryInfo(uploadPath); //Assuming Test is your Folder
                var uploadFiles = dirUpFiles.GetFiles("*.*")
                                .Where(f => f.Extension.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) //Getting log files .txt, .TXT, hoặc .Txt
                                .ToArray();
                //var txtFiles = Directory.GetFiles(uploadPath, "*.*")
                //               .Where(file => file.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                //               .ToArray();


                var dirDownloadPath = new DirectoryInfo(downloadPath);
                var downloadFilesXlsx = dirDownloadPath.GetFiles("*.xlsx");
                var downloadFilesCsv = dirDownloadPath.GetFiles("*.csv");

                var dirLogPath = new DirectoryInfo(logPath);
                var logFiles = dirLogPath.GetFiles("*.log");

                var totalFiles = uploadFiles.Concat(downloadFilesXlsx).Concat(downloadFilesCsv).Concat(logFiles).ToArray();
                var dtCreateLimit = DateTime.Now.AddDays(-MostRecentDays);
                foreach (var file in totalFiles)
                {
                    try
                    {
                        var fFirstTime = file.CreationTime;
                        var fLastTime = file.LastWriteTime;
                        if (fLastTime < dtCreateLimit)
                        {
                            file.Delete();
                            _logger.LogInformation($"File {file.Name} is deleted!!");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError($"File {file.Name} cannot be deleted!");
                        _logger.LogError($"Exception Message : {ex.Message}!");
                    }
                }
                Debug.WriteLine("Startup Task Executed!");
                _logger.LogInformation("Startup Task is Executed!");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Startup Task Error!");
                _logger.LogError("Startup Task Error!, Exception Message : " + ex.Message);
            }

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            // Place code to run when the application stops, if any
            Debug.WriteLine("Startup Stoped.");
            return Task.CompletedTask;
        }
    }
}
