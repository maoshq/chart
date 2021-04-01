using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SampleTabularApiApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread.Sleep(1000);
            String[] DataSetArr = new String[] {"ActiveUsage", "Pen", "AppUsage", "Power", "Cabs", "Cabs"
                , "Firmware", "Performance", "OSAdoption", "Reliability", "TouchPad","Feedback","Hello"};
            string tenantId = Constants.TenantId;
            string clientId = Constants.ClientId;
            string clientSecret = Constants.ClientSecret;

            string scope = "https://manage.devcenter.microsoft.com";

            string accessToken = GetClientCredentialAccessToken(
                    tenantId,
                    clientId,
                    clientSecret,
                    scope).Result;

            var dataQueryService = new DataQueryService();

            try
            {
                var files = Directory.GetFiles(System.Environment.CurrentDirectory + @"\JSONFiles\", "*.json");
                foreach (String fileName in files)
                {                  
                    foreach (String i in DataSetArr)
                    {
                        if (fileName.Contains(i))
                        {
                            String query = File.ReadAllText(fileName);
                            string filename = Path.GetFileNameWithoutExtension(fileName);
                            Console.WriteLine("JSON文件内容----------------------- \n" + query);
                            Console.WriteLine("当前目标数据集 : " + i);
                            dataQueryService.QueryDataAsync(i, query, accessToken, filename).Wait();
                        }
                    }
                }
                   
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadLine();
            }
        }

        public static async Task<string> GetClientCredentialAccessToken(string tenantId,
            string clientId,
            string clientSecret,
            string scope)
        {
            string tokenEndpointFormat = "https://login.microsoftonline.com/{0}/oauth2/token";
            string tokenEndpoint = string.Format(tokenEndpointFormat, tenantId);

            dynamic result;
            using (HttpClient client = new HttpClient())
            {
                string tokenUrl = tokenEndpoint;
                using (
                    HttpRequestMessage request = new HttpRequestMessage(
                        HttpMethod.Post,
                        tokenUrl))
                {
                    string requestContent =
                        $"grant_type=client_credentials&client_id={clientId}&client_secret={clientSecret}&resource={scope}";

                    request.Content = new StringContent(requestContent, Encoding.UTF8, "application/x-www-form-urlencoded");

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new Exception(
                                $"Could not authenticate. ResponseCode: {response.StatusCode}, Contents: {responseContent}");
                        }

                        result = JsonConvert.DeserializeObject(responseContent);
                    }
                }
            }

            return result.access_token;
        }
    }
}
