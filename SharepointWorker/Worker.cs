using Cronos;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace SharepointWorker
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {

            // Authentication
            var app = ConfidentialClientApplicationBuilder.Create("694e9844-f1f6-4ea5-ab2b-43e2e7bddd6b")
                .WithClientSecret("UWW8Q~WYrY3QzNgRI55wXmOebSVsBinItsC_iaTP")
                .WithAuthority(AzureCloudInstance.AzurePublic, "0ff63586-3b7a-4f81-a766-7409f1ba5ae7")
                .Build();
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);

                try
                {
                    List<string> scopes = new List<string>
                    {
                        "https://graph.microsoft.com/.default"
                    };

                    var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

                    var accessToken = result.AccessToken;

                    string siteUrl = "https://graph.microsoft.com/v1.0/sites/avanade.sharepoint.com:/sites/Fibrasil-files";
                    HttpClient client = new HttpClient();
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, siteUrl);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    HttpResponseMessage response = await client.SendAsync(request);
                    string responseBody = await response.Content.ReadAsStringAsync();
                    var site = JsonConvert.DeserializeObject<dynamic>(responseBody);
                    string siteId = site.id;

                    string driveUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives";
                    request = new HttpRequestMessage(HttpMethod.Get, driveUrl);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    response = await client.SendAsync(request);
                    responseBody = await response.Content.ReadAsStringAsync();
                    var drives = JsonConvert.DeserializeObject<dynamic>(responseBody);
                    string driveId = drives.value[0].id;  // supondo que o arquivo está no primeiro drive listado

                    string filePath = "/Teste/Book.xlsx";  // caminho relativo ao arquivo no SharePoint
                    string itemUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/root:/{filePath}";
                    request = new HttpRequestMessage(HttpMethod.Get, itemUrl);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    response = await client.SendAsync(request);
                    responseBody = await response.Content.ReadAsStringAsync();
                    var item = JsonConvert.DeserializeObject<dynamic>(responseBody);
                    string itemId = item.id;


                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error uploading file: {ex.Message}");
                }
                await Task.Delay(CalculateNextDelay(), stoppingToken);

            }
        }
        private TimeSpan CalculateNextDelay()
        {
            var cronExpression = CronExpression.Parse("*/5 * * * *");  // Executa a cada 5 minutos
            var next = cronExpression.GetNextOccurrence(DateTime.Now);

            if (next.HasValue)
            {
                return next.Value - DateTime.Now;
            }

            return TimeSpan.FromMinutes(30);
        }

    }
}