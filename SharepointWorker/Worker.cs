using Cronos;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net.Http.Headers;
using System.Text;
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
            //var app = ConfidentialClientApplicationBuilder.Create("694e9844-f1f6-4ea5-ab2b-43e2e7bddd6b")
            //    .WithClientSecret("UWW8Q~WYrY3QzNgRI55wXmOebSVsBinItsC_iaTP")
            //    .WithAuthority(AzureCloudInstance.AzurePublic, "0ff63586-3b7a-4f81-a766-7409f1ba5ae7")
            //    .Build();

            var app = ConfidentialClientApplicationBuilder.Create("a06c2c16-7079-463a-9c44-3bba0c62016c")
    .WithClientSecret("STh8Q~AB3C00lwL_D6A5DQrSzmiBuju.035MMbaE")
    .WithAuthority(AzureCloudInstance.AzurePublic, "dcc867ff-f3cc-40f3-8e63-c33b8872a692")
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

                    string siteUrl = "https://graph.microsoft.com/v1.0/sites/vx65k.sharepoint.com:/sites/fibrasil";
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

                    string sheetName = "Sheet1";  // nome da planilha
                    string rangeAddress = "A1:B2";  // intervalo que deseja ler
                    string rangeUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/workbook/worksheets/{sheetName}/range(address='{rangeAddress}')";
                    request = new HttpRequestMessage(HttpMethod.Get, rangeUrl);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    response = await client.SendAsync(request);
                    responseBody = await response.Content.ReadAsStringAsync();
                    var range = JsonConvert.DeserializeObject<dynamic>(responseBody);
                    JArray rangeValues = range.text;
                    foreach (var value in rangeValues)
                    {
                        Console.WriteLine(value);
                    }
                    // valores do intervalo

                    var updateValues = new
                    {
                        values = new string[][]
    {
        new string[] { "New data 4", "New data 3" },
        new string[] { "New data 2", "New data 1" }
    }
                    };
                    request = new HttpRequestMessage(HttpMethod.Patch, rangeUrl)
                    {
                        Content = new StringContent(JsonConvert.SerializeObject(updateValues), Encoding.UTF8, "application/json")
                    };
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    response = await client.SendAsync(request);

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