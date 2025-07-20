namespace ReportPolicy.Services.Report
{
    public class AdReportService : BackgroundService
    {
        private readonly IServiceProvider _services;
        private readonly ILogger<AdReportService> _logger;

        public AdReportService(IServiceProvider services, ILogger<AdReportService> logger)
        {
            _services = services;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    using var scope = _services.CreateScope();
                    var logic = scope.ServiceProvider.GetRequiredService<AdReportLogic>();
                    await logic.RunAsync();
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in AdReportService");
                }

                await Task.Delay(TimeSpan.FromMinutes(10), stoppingToken);
            }
        }
    }

}
