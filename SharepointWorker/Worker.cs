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

            
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);

                try
                {

                    var app = ConfidentialClientApplicationBuilder.Create("a06c2c16-7079-463a-9c44-3bba0c62016c")
                        .WithClientSecret("STh8Q~AB3C00lwL_D6A5DQrSzmiBuju.035MMbaE")
                        .WithAuthority(AzureCloudInstance.AzurePublic, "dcc867ff-f3cc-40f3-8e63-c33b8872a692")
                        .Build();

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
                    string rangeUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/workbook/worksheets/{sheetName}/usedRange";
                    request = new HttpRequestMessage(HttpMethod.Get, rangeUrl);
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    response = await client.SendAsync(request);
                    responseBody = await response.Content.ReadAsStringAsync();
                    var range = JsonConvert.DeserializeObject<dynamic>(responseBody);
                    JArray rangeValues = range.text;
                    int statusColumnIndex = -1;
                    for (int i = 0; i < ((JArray)rangeValues[0]).Count; i++)
                    {
                        if (rangeValues[0][i].ToString() == "status")
                        {
                            statusColumnIndex = i;
                            break;
                        }
                    }

                    if (statusColumnIndex == -1)
                    {
                        throw new Exception("Coluna 'status' não encontrada");
                    }

                    List<int> rowsToProcess = new List<int>();
                    for (int i = 1; i < rangeValues.Count; i++)  // começa em 1 para pular a linha de cabeçalho
                    {
                        if (rangeValues[i][statusColumnIndex].ToString() == "A PROCESSAR")
                        {
                            rowsToProcess.Add(i);
                        }
                    }

                    int batchSize = 20;  // definir o tamanho do lote como 20
                    int totalBatches = (int)Math.Ceiling((double)rowsToProcess.Count / batchSize);  // calculando o total de lotes

                    for (int i = 0; i < totalBatches; i++)
                    {
                        List<string> requests = new List<string>();
                        for (int j = i * batchSize; j < Math.Min(rowsToProcess.Count, (i + 1) * batchSize); j++)
                        {
                            int row = rowsToProcess[j];
                            string columnLetter = ColumnLetter(statusColumnIndex);
                            string cellAddress = $"{columnLetter}{row + 1}";
                            string updateUrl = $"/sites/{siteId}/drives/{driveId}/items/{itemId}/workbook/worksheets/{sheetName}/range(address='{cellAddress}')";

                            var updateValues = new
                            {
                                values = new string[][]
                                {
                                    new string[] { "EM PROCESSAMENTO" }
                                }
                            };
                            var updateValuesJson = JsonConvert.SerializeObject(updateValues);
                            string requestContent = $"{{\"id\":\"{j}\",\"method\":\"PATCH\",\"url\":\"{updateUrl}\",\"headers\":{{\"Content-Type\":\"application/json\"}},\"body\":{updateValuesJson}}}";
                            requests.Add(requestContent);
                        }
                        string batchContent = $"{{\"requests\":[{string.Join(",", requests)}]}}";
                        HttpRequestMessage batchRequest = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/$batch")
                        {
                            Content = new StringContent(batchContent, Encoding.UTF8, "application/json")
                        };
                        batchRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                        HttpResponseMessage batchResponse = await client.SendAsync(batchRequest);
                        string batchResponseContent = await batchResponse.Content.ReadAsStringAsync();
                        var batchResponseJson = JsonConvert.DeserializeObject<dynamic>(batchResponseContent);
                        foreach (var batchResponset in batchResponseJson.responses)
                        {
                            if (batchResponset.status >= 400)
                            {
                                Console.WriteLine($"Request failed: {batchResponset.status} {batchResponset.body.error.message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error uploading file: {ex.Message}");
                }
                await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken);

            }
        }

        public static string ColumnLetter(int intCol)
        {
            intCol++;
            var intFirstLetter = ((intCol) / 676) + 64;
            var intSecondLetter = ((intCol % 676) / 26) + 64;
            var intThirdLetter = (intCol % 26) + 65;

            var firstLetter = (intFirstLetter > 64) ? (char)intFirstLetter : ' ';
            var secondLetter = (intSecondLetter > 64) ? (char)intSecondLetter : ' ';
            var thirdLetter = (char)intThirdLetter;

            return string.Concat(firstLetter, secondLetter, thirdLetter).Trim();
        }


    }
}