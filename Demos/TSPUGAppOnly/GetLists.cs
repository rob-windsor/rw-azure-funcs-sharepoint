using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Linq;
using System.Collections.Generic;

namespace TSPUG.TSPUGAppOnly
{
    public static class GetLists
    {
        [FunctionName("GetLists")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string siteUrlKey = "siteUrl";
            string siteUrl = null;

            if (req.Query.ContainsKey(siteUrlKey)) 
            {
                siteUrl = req.Query[siteUrlKey].First();
            }

            if (siteUrl == null)
            {
                return new BadRequestObjectResult(
                    "Please pass a site URL on the query string");
            }

            var clientId = "c582d811-690b-4443-85bf-4a9c16f68006";
            var tenantName = "robwindsor2";
            var certName = "AzureFunctionAppTest.pfx";
            var certPassword = "pass@word1";

            var home = Environment.GetEnvironmentVariable("HOME");
            var certPath = home != null ?
                Path.Combine(home, @"site\wwwroot", certName) :
                Path.Combine(@"E:\Certs\" + certName);

            siteUrl = $"https://{tenantName}.sharepoint.com" + siteUrl;
            var listNames = new List<string>();

            var authManager = new PnP.Framework.AuthenticationManager(clientId, certPath, certPassword, $"{tenantName}.onmicrosoft.com");
            using (var context = authManager.GetContext(siteUrl))
            {
                var query = context.Web.Lists.Where(l => l.Hidden == false)
                    .OrderBy(l => l.Title);

                var lists = context.LoadQuery(query);
                context.ExecuteQuery();

                foreach (var list in lists)
                {
                    listNames.Add(list.Title);
                }
            }
            
            var responseContent = JsonConvert.SerializeObject(
                listNames, Formatting.Indented);
            return new OkObjectResult(responseContent);
        }
    }
}
