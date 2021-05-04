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
using Microsoft.Identity.Client;
using System.Collections.Generic;

namespace TSPUG.TSPUGDelegated
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
                    "Please pass a site URL on the query string " +
                    "or in the request body");
            }

            var clientId = "1d582c6c-9936-4393-9220-36823a23500b";
            var clientSecret = "X3cMt-00Ju-WnX9n0e8wY.d7uLd4E~y9xy";
            var tenantName = "robwindsor2";
            var token = req.Headers["Authorization"].ToString().Remove(0, "Bearer ".Length);
            var userAssertion = new UserAssertion(token);

            siteUrl = $"https://{tenantName}.sharepoint.com" + siteUrl;
            var listNames = new List<string>();

            var authManager = new PnP.Framework.AuthenticationManager(clientId, clientSecret, userAssertion);
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
