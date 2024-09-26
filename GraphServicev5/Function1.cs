using System;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Azure.Identity;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Net.Http;
using Microsoft.Identity.Client;

namespace GraphServicev5
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task Run([TimerTrigger("*/1 * * * *")] TimerInfo myTimer, ILogger log)
        {
            try
            {
                var scopes = new[] { "https://graph.microsoft.com/.default" };
                var clientId = "123123";
                var tenantId = "123123";
                var clientSecret = "123123";

                //get token
                var token = await GetTokenAsync(clientId, tenantId, clientSecret);
                log.LogInformation($"Token: {token}");
                var options = new ClientSecretCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                };

                var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret, options);
                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

                // Fetch the list of users in the directory
                var users = await graphClient.Users.GetAsync();
                await GetEmailsAsync(token);
                foreach (var user in users.Value)
                {
                    log.LogInformation($"User: {user.DisplayName}, UPN: {user.UserPrincipalName}");
                }
            }
            catch (Exception ex)
            {
                log.LogError($"Error: {ex.Message}");
            }
        }

        private static async Task<string> GetTokenAsync(string clientid, string tenantid, string clientsecret)
        {
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var app = ConfidentialClientApplicationBuilder.Create(clientid)
                .WithClientSecret(clientsecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantid}"))
                .Build();

            var result = await app.AcquireTokenForClient(scopes)
                .ExecuteAsync();

            return result.AccessToken;
        }

        private static async Task GetEmailsAsync(string accessToken)
        {
            var graphUrl = "https://graph.microsoft.com/v1.0/me/messages";  // You can also use /users/{user-id}/messages

            using (var httpClient = new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                var response = await httpClient.GetAsync(graphUrl);
                var content = await response.Content.ReadAsStringAsync();

                Console.WriteLine(content);
            }
        }

    }
}
