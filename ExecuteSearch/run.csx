#r "Newtonsoft.Json"
using System;
using System.Net;
using System.Text;
using Newtonsoft.Json;
using OfficeDevPnP.Core;
using OfficeDevPnP.Core.Pages;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Newtonsoft.Json;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
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
    log.Info($"Received siteUrl={siteUrl}");

    log.Info($"Will attempt to authenticate to SharePoint with username {adminUserName}");

    // auth to SharePoint and get ClientContext..
    ClientContext siteContext = new OfficeDevPnP.Core.AuthenticationManager().GetSharePointOnlineAuthenticatedContextTenant(siteUrl, adminUserName, adminPassword);
    Site site = siteContext.Site;
    siteContext.Load(site);
    siteContext.ExecuteQueryRetry();

    log.Info($"Successfully authenticated to site {siteContext.Url}..");

    KeywordQuery keywordQuery = new KeywordQuery(siteContext);
    keywordQuery.QueryText = "SharePoint";
    SearchExecutor searchExecutor = new SearchExecutor(siteContext);
    ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
    siteContext.ExecuteQuery();
 
    //var jsonToReturn = JsonConvert.SerializeObject(results);   
    var jsonToReturn = "Success";

    if(jsonToReturn == null)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, "Failed to parse json.")
    }
    else{
        
         return new HttpResponseMessage(HttpStatusCode.OK) {
          Content = new StringContent(jsonToReturn, Encoding.UTF8, "application/json")
        };
    }
        
}
