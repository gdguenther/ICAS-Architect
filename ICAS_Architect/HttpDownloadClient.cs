using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace ICAS_Architect
{
    // if ADAL is selected this client will be used to download metadata
    // if browser control is used this client will NOT be used. Download is done by the browser control on frmMain.
    // source https://github.com/microsoft/PowerApps-Samples/tree/master/cds/webapi/C%23/ADALV3WhoAmI

    internal class HttpDownloadClient
    {
        // The URL to the CDS environment you want to connect with
        // i.e. https://yourOrg.crm.dynamics.com/
        private string baseUrl;

        // Azure Active Directory registered app clientid for Microsoft samples
        private const string clientId = "51f81489-12ee-4a9e-aaae-a2591f45987d";
        private const string redirectUri = "app://58145B91-0C36-4500-8554-080854F2AC97";

        public string accessToken;
        public  HttpClient client;

        internal HttpDownloadClient(string baseUrl)
        {
            this.baseUrl = baseUrl;
        }

        internal int Connect(string whoamiUrl)
 //       internal Task<Guid> Connect(string whoamiUrl)
        {
//            const string authority = "https://login.microsoftonline.com/common";
            const string authority = "https://login.microsoftonline.com/common";
            var context = new AuthenticationContext(authority, false);

            //GG: going to try auto for a while
            //var platformParameters = new PlatformParameters(PromptBehavior.SelectAccount);
            var platformParameters = new PlatformParameters(PromptBehavior.Auto);

            accessToken = context.AcquireTokenAsync(baseUrl, clientId, new Uri(redirectUri), platformParameters, UserIdentifier.AnyUser).Result.AccessToken;



            //GG: Security Protocol must be set to Tls12 which VSTO tries to override
            System.Net.ServicePointManager.SecurityProtocol = (System.Net.SecurityProtocolType)3072;

            client = new HttpClient();
            client.BaseAddress = new Uri(baseUrl);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            client.Timeout = new TimeSpan(0, 2, 0);
            client.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
            client.DefaultRequestHeaders.Add("OData-Version", "4.0");
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            /*            // Use the WhoAmI function for testing
                        var response = client.GetAsync(whoamiUrl).Result;
            //            var response = Task.Run(() => client.GetAsync(whoamiUrl)).Result;

                        //Get the response content and parse it.  
                        response.EnsureSuccessStatusCode();
                        JObject body = JObject
                            .Parse(response.Content.ReadAsStringAsync().Result);
                        Guid userId = (Guid)body["UserId"];*/
            return 0;
        }

        internal string Fetch(string url)
        {
            try
            {
                var response = client.GetAsync(url).Result;
                response.EnsureSuccessStatusCode();
                return response.Content.ReadAsStringAsync().Result;
            }
            catch (Exception ex)
            {
                throw new HttpDownloadClientException($"Failed to download {url}. Please make sure you can open this URL.", ex);
            }
        }
    }

    internal class HttpDownloadClientException : Exception
    {
        internal HttpDownloadClientException(string message, Exception innerException):base(message, innerException)
        {

        }
    }
}
