using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

namespace graph_console
{
    public class GraphHelper
    {
        public static DeviceCodeCredential tokenCredential;

        public static GraphServiceClient graphClient;

        public static void Initialize(
            string clientId,
            string[] scopes,
            Func<DeviceCodeInfo, CancellationToken, Task> callBack
        )
        {
            tokenCredential = new DeviceCodeCredential(callBack, clientId);
            graphClient = new GraphServiceClient(tokenCredential, scopes);
        }

        public static async Task<string> GetAccessTokenAsync(string[] scopes)
        {
            var context = new TokenRequestContext(scopes);
            var response = await tokenCredential.GetTokenAsync(context);
            return response.Token;
        }

        public static async Task<User> GetMeAsync()
        {
            try 
            {
                return await graphClient.Me.Request().GetAsync();
            }
            catch(ServiceException ex) 
            {
                Console.WriteLine($"Error getting user's profile: {ex.Message}");
                return null;
            }
            
        }


    }
}