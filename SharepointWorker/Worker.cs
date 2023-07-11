using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using static Microsoft.Graph.Constants;

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
            while (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);

                try
                {
                    var app = ConfidentialClientApplicationBuilder.Create("a06c2c16-7079-463a-9c44-3bba0c62016c")
                        .WithClientSecret("STh8Q~AB3C00lwL_D6A5DQrSzmiBuju.035MMbaE")
                        .WithAuthority(AzureCloudInstance.AzurePublic, "dcc867ff-f3cc-40f3-8e63-c33b8872a692")
                        .Build();

                    var scopes = new List<string> { "https://graph.microsoft.com/.default" };
                    var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
                    var accessToken = result.AccessToken;

                    string siteId = await GetSiteId(accessToken, "https://graph.microsoft.com/v1.0/sites/vx65k.sharepoint.com:/sites/fibrasil");
                    string driveId = await GetDriveId(accessToken, $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives");
                    string itemId = await GetItemId(accessToken, $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/root:/Teste/Book.xlsx");
                    var rangeValues = await GetRangeValues(accessToken, $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/workbook/worksheets/Sheet1/usedRange");

                    int statusColumnIndex = GetStatusColumnIndex(rangeValues);

                    if (statusColumnIndex == -1)
                    {
                        throw new Exception("Coluna 'status' não encontrada");
                    }

                    var rowsToProcess = GetRowsToProcess(rangeValues, statusColumnIndex);
                    await UpdateCellsInBatches(accessToken, rowsToProcess, statusColumnIndex, siteId, driveId, itemId, stoppingToken);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error uploading file: {ex.Message}");
                }

                await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken);
            }
        }

        private async Task<string> GetSiteId(string accessToken, string siteUrl)
        {
            return await GetIdFromUrl(accessToken, siteUrl);
        }

        private async Task<string> GetDriveId(string accessToken, string driveUrl)
        {
            var response = await GetResponseFromUrl(accessToken, driveUrl);
            var jsonObject = JsonConvert.DeserializeObject<dynamic>(response);
            return jsonObject.value[0].id;
        }

        private async Task<string> GetItemId(string accessToken, string itemUrl)
        {
            return await GetIdFromUrl(accessToken, itemUrl);
        }

        private async Task<string> GetIdFromUrl(string accessToken, string url)
        {
            var response = await GetResponseFromUrl(accessToken, url);
            var jsonObject = JsonConvert.DeserializeObject<dynamic>(response);
            return jsonObject.id;
        }

        private async Task<string> GetResponseFromUrl(string accessToken, string url)
        {
            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            var response = await client.SendAsync(request);
            return await response.Content.ReadAsStringAsync();
        }

        private async Task<JArray> GetRangeValues(string accessToken, string rangeUrl)
        {
            var response = await GetResponseFromUrl(accessToken, rangeUrl);
            var range = JsonConvert.DeserializeObject<dynamic>(response);
            return range.text;
        }

        private int GetStatusColumnIndex(JArray rangeValues)
        {
            for (int i = 0; i < ((JArray)rangeValues[0]).Count; i++)
            {
                if (rangeValues[0][i].ToString().ToLower() == "status")
                {
                    return i;
                }
            }
            return -1;
        }

        private List<int> GetRowsToProcess(JArray rangeValues, int statusColumnIndex)
        {
            var rowsToProcess = new List<int>();
            for (int i = 1; i < rangeValues.Count; i++)

            {
                if (rangeValues[i][statusColumnIndex].ToString().ToUpper() == "A PROCESSAR")
                {
                    rowsToProcess.Add(i);
                }
            }
            return rowsToProcess;
        }

        private async Task UpdateCellsInBatches(string accessToken, List<int> rowsToProcess, int statusColumnIndex, string siteId, string driveId, string itemId, CancellationToken stoppingToken)
        {
            string sheetName = "Sheet1";
            var client = new HttpClient();

            foreach (List<int> batchRows in rowsToProcess.Batch(20))
            {
                string boundary = $"batch_{Guid.NewGuid()}";
                MultipartContent batchContent = new MultipartContent("mixed", boundary);

                foreach (int row in batchRows)
                {
                    string columnLetter = ColumnLetter(statusColumnIndex);
                    string cellAddress = $"{columnLetter}{row + 1}";
                    string updateUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/items/{itemId}/workbook/worksheets/{sheetName}/range(address='{cellAddress}')";
                    var updateValues = new
                    {
                        values = new string[][]
                        {
                    new string[] { "EM PROCESSAMENTO" }
                        }
                    };
                    HttpRequestMessage updateRequestMessage = new HttpRequestMessage(HttpMethod.Patch, updateUrl)
                    {
                        Content = new StringContent(JsonConvert.SerializeObject(updateValues), Encoding.UTF8, "application/json")
                    };
                    updateRequestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // Cria um novo conteúdo HttpMessageContent a partir da solicitação
                    HttpMessageContent content = new HttpMessageContent(updateRequestMessage);
                    // Adiciona o conteúdo da mensagem ao conteúdo do lote
                    batchContent.Add(content);
                }

                HttpRequestMessage batchRequest = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/$batch")
                {
                    Content = batchContent
                };
                batchRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                HttpResponseMessage batchResponse = await client.SendAsync(batchRequest);
                string batchResponseContent = await batchResponse.Content.ReadAsStringAsync();
                var batchResult = JsonConvert.DeserializeObject<BatchResponseContent>(batchResponseContent);

                // Verifique os resultados da solicitação de lote para garantir que todas as atualizações foram bem-sucedidas.

                await Task.Delay(TimeSpan.FromSeconds(60), stoppingToken);  // aguarde 60 segundos antes de enviar o próximo lote
            }
        }



        private static string ColumnLetter(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}
