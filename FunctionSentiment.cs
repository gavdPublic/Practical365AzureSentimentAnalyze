using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;

using System;
using System.Collections.Generic;
using System.Security;
using Newtonsoft.Json;
using Microsoft.SharePoint.Client;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics;
using Microsoft.Azure.CognitiveServices.Language.TextAnalytics.Models;
using Microsoft.Rest;
using System.Threading;

namespace PracticalSentimentFunctionApp
{
    public static class FunctionSentiment
    {
        private static readonly string serviceKey = "60dbc25b910f4780xxx";
        private static readonly string serviceEndpoint = "https://[domain].cognitiveservices.azure.com/";

        [FunctionName("FunctionSentiment")]
        public static async Task<HttpResponseMessage> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, 
            TraceWriter log)
        {
            // Registration
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;
            if (validationToken != null)
            {
                var myResponse = req.CreateResponse(HttpStatusCode.OK);
                myResponse.Content = new StringContent(validationToken);
                return myResponse;
            }

            // Changes
            var myContent = await req.Content.ReadAsStringAsync();
            var allNotifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(myContent).Value;

            if (allNotifications.Count > 0)
            {
                foreach (var oneNotification in allNotifications)
                {
                    // Login in SharePoint
                    string baseUrl = "https://somewhere.sharepoint.com/";
                    string myUserName = "someone@somewhere.onmicrosoft.com";
                    string myPassword = "VerySecurePw"; // This is a proof-of-concept, don't do this in production

                    SecureString securePassword = new SecureString();
                    foreach (char oneChar in myPassword) securePassword.AppendChar(oneChar);
                    SharePointOnlineCredentials myCredentials = new SharePointOnlineCredentials(myUserName, securePassword);

                    ClientContext SPClientContext = new ClientContext(baseUrl + oneNotification.SiteUrl);
                    SPClientContext.Credentials = myCredentials;

                    // Get the Changes
                    GetChanges(SPClientContext, oneNotification.Resource, log);
                }
            }

            return new HttpResponseMessage(HttpStatusCode.OK);
        }

        static void GetChanges(ClientContext SPClientContext, string ListId, TraceWriter log)
        {
            // Get the List
            Web spWeb = SPClientContext.Site.RootWeb;
            List changedList = spWeb.Lists.GetById(new Guid(ListId));
            SPClientContext.Load(changedList);
            SPClientContext.ExecuteQuery();

            // Create the ChangeToken and Change Query
            ChangeToken lastChangeToken = new ChangeToken();
            lastChangeToken.StringValue = string.Format("1;3;{0};{1};-1", ListId, DateTime.Now.AddMinutes(-1).ToUniversalTime().Ticks.ToString());
            ChangeToken newChangeToken = new ChangeToken();
            newChangeToken.StringValue = string.Format("1;3;{0};{1};-1", ListId, DateTime.Now.ToUniversalTime().Ticks.ToString());
            ChangeQuery myChangeQuery = new ChangeQuery(false, false);
            myChangeQuery.Item = true;  // Get only Item changes
            myChangeQuery.Add = true;   // Get only the new Items
            myChangeQuery.ChangeTokenStart = lastChangeToken;
            myChangeQuery.ChangeTokenEnd = newChangeToken;

            // Get all Changes
            var allChanges = changedList.GetChanges(myChangeQuery);
            SPClientContext.Load(allChanges);
            SPClientContext.ExecuteQuery();

            foreach (Change oneChange in allChanges)
            {
                if (oneChange is ChangeItem)
                {
                    // Get what is changed
                    ListItem changedListItem = changedList.GetItemById((oneChange as ChangeItem).ItemId);
                    SPClientContext.Load(changedListItem);
                    SPClientContext.ExecuteQuery();

                    // Create a Text Analytics client
                    ApiKeyServiceClientCredentials myCredentials = new ApiKeyServiceClientCredentials(serviceKey);
                    TextAnalyticsClient myClient = new TextAnalyticsClient(myCredentials)
                    {
                        Endpoint = serviceEndpoint
                    };

                    // Call the service
                    SentimentResult myScore = myClient.Sentiment(changedListItem["Comments"].ToString(), "en");

                    // Insert the values back in the Item
                    changedListItem["Sentiment"] = myScore.Score.ToString();
                    changedListItem.Update();
                    SPClientContext.ExecuteQuery();
                }
            }
        }
    }

    class ApiKeyServiceClientCredentials : ServiceClientCredentials
    {
        private readonly string apiKey;

        public ApiKeyServiceClientCredentials(string apiKey)
        {
            this.apiKey = apiKey;
        }

        public override Task ProcessHttpRequestAsync(HttpRequestMessage request, 
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            request.Headers.Add("Ocp-Apim-Subscription-Key", this.apiKey);
            return base.ProcessHttpRequestAsync(request, cancellationToken);
        }
    }

    public class ResponseModel<T>
    {
        [JsonProperty(PropertyName = "value")]
        public List<T> Value { get; set; }
    }

    public class NotificationModel
    {
        [JsonProperty(PropertyName = "subscriptionId")]
        public string SubscriptionId { get; set; }

        [JsonProperty(PropertyName = "clientState")]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "resource")]
        public string Resource { get; set; }

        [JsonProperty(PropertyName = "tenantId")]
        public string TenantId { get; set; }

        [JsonProperty(PropertyName = "siteUrl")]
        public string SiteUrl { get; set; }

        [JsonProperty(PropertyName = "webId")]
        public string WebId { get; set; }
    }

    public class SubscriptionModel
    {
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public string Id { get; set; }

        [JsonProperty(PropertyName = "clientState", NullValueHandling = NullValueHandling.Ignore)]
        public string ClientState { get; set; }

        [JsonProperty(PropertyName = "expirationDateTime")]
        public DateTime ExpirationDateTime { get; set; }

        [JsonProperty(PropertyName = "notificationUrl")]
        public string NotificationUrl { get; set; }

        [JsonProperty(PropertyName = "resource", NullValueHandling = NullValueHandling.Ignore)]
        public string Resource { get; set; }
    }
}
