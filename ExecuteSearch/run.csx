#r "Newtonsoft.Json"
using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using Microsoft.Azure.WebJobs;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Pages;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;

public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]HttpRequestMessage req, TraceWriter log)
{
    log.Info("C# HTTP trigger function processed a request.");
    string ADMIN_USER_CONFIG_KEY = "SharePointUser";
    string ADMIN_PASSWORD_CONFIG_KEY = "SharePointUserPwd";
    string adminUserName = System.Environment.GetEnvironmentVariable(ADMIN_USER_CONFIG_KEY, EnvironmentVariableTarget.Process);
    string adminPassword = System.Environment.GetEnvironmentVariable(ADMIN_PASSWORD_CONFIG_KEY, EnvironmentVariableTarget.Process); 
    //

    // collect site/page details from request body..
    dynamic dataX = await req.Content.ReadAsAsync<object>();
    string siteUrl = dataX.SiteUrl;
    string queryText = dataX.QueryText;
    int maxItems = 1;
    if(dataX.MaxItems!=null)
    {
        maxItems = dataX.MaxItems;   
    }
    
    log.Info($"Received siteUrl={siteUrl}");
    log.Info($"Using keyword '{queryText}' for the search");

    log.Info($"Will attempt to authenticate to SharePoint with username {adminUserName}");

    // auth to SharePoint and get ClientContext..
    ClientContext siteContext = new OfficeDevPnP.Core.AuthenticationManager().GetSharePointOnlineAuthenticatedContextTenant(siteUrl, adminUserName, adminPassword);
    Site site = siteContext.Site;
    siteContext.Load(site);
    siteContext.ExecuteQueryRetry();

    log.Info($"Successfully authenticated to site {siteContext.Url}..");

    KeywordQuery keywordQuery = new KeywordQuery(siteContext);
    keywordQuery.QueryText = queryText;
    keywordQuery.RowLimit = maxItems;
    keywordQuery.SelectProperties.Add("Title");
    keywordQuery.SelectProperties.Add("PictureUrl");
    keywordQuery.SelectProperties.Add("SipAddress");
    keywordQuery.SourceId = new Guid("b09a7990-05ea-4af9-81ef-edfab16c4e31");
    keywordQuery.RankingModelId = "D9BFB1A1-9036-4627-83B2-BBD9983AC8A1";
    SearchExecutor searchExecutor = new SearchExecutor(siteContext);
    ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
    siteContext.ExecuteQuery();
 
    var jsonToReturn = JsonConvert.SerializeObject(results);   
    //var jsonToReturn = "Success";

    if(jsonToReturn == null)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, "Failed to parse json.");
    }
    else{
        
         return new HttpResponseMessage(HttpStatusCode.OK) {
          Content = new StringContent(jsonToReturn, Encoding.UTF8, "application/json")
        };
    }
        
}
