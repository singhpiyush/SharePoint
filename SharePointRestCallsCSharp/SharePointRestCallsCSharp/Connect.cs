// ===========================================
// Copyright (c) 2017. All rights reserved.
// Author:: Piyush Kumar Singh
// Purpose:: SharePoint REST calls from C#
// ===========================================
// Change History
// 01/01/2017		=> First version
// ===========================================


using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace SharePointRestCallsCSharp
{
    class ConnectPk
    {
        private string webUrl;
        private SP.SharePointOnlineCredentials credential;

        /// <summary>
        /// Default Contructor
        /// </summary>
        /// <param name="webUrl">SharePoint web url</param>
        /// <param name="userId">userID</param>
        /// <param name="pwd">password</param>
        internal ConnectPk(string webUrl, string userId, string pwd)
        {
            this.webUrl = webUrl;

            SecureString password = new SecureString();
            pwd.ToList().ForEach(password.AppendChar);
            this.credential = new SP.SharePointOnlineCredentials(userId, password);
        }

        /// <summary>
        /// Get ListItems based on the given parameter
        /// </summary>
        /// <param name="restUrl">ListItem fetch request url</param>
        /// <returns>returns the json output</returns>
        internal async Task<string> GetListItems(string restUrl)
        {
            return await Connect(restUrl, "ListItems");
        }

        /// <summary>
        /// Get List based on the given parameter
        /// </summary>
        /// <param name="restUrl">List fetch request url</param>
        /// <returns>returns the json output</returns>
        internal async Task<string> GetList(string restUrl)
        {
            return await Connect(restUrl, "List");
        }

        /// <summary>
        /// Get Web properties based on the given parameter
        /// </summary>
        /// <param name="restUrl">Web fetch request url</param>
        /// <returns>returns the json output</returns>
        internal async Task<string> GetWeb(string restUrl)
        {
            return await Connect(restUrl, "Web");
        }

        /// <summary>
        /// Invokes SharePoint REST call
        /// </summary>
        /// <param name="restUrlSuffix">REST url query</param>
        /// <param name="fileName">(optional) Name of the file where the json output will be dumped.</param>
        /// <returns>returns the json output</returns>
        private async Task<string> Connect(string restUrlSuffix, string fileName)
        {
            using (var handler = new HttpClientHandler() { Credentials = credential })
            {
                //Getting authentication cookies 
                Uri uri = new Uri(webUrl);
                handler.CookieContainer.SetCookies(uri, credential.GetAuthenticationCookie(uri));

                //Invoking REST API 
                using (var client = new HttpClient(handler))
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    HttpResponseMessage response = await client.GetAsync(string.Format(restUrlSuffix, webUrl)).ConfigureAwait(false);
                    response.EnsureSuccessStatusCode();

                    string jsonData = await response.Content.ReadAsStringAsync();

                    if (!String.IsNullOrWhiteSpace(fileName))
                    {
                        //Debug.Write(jsonData);
                        File.WriteAllText(fileName, jsonData);
                    }

                    return jsonData;
                }
            }
        }
    }
}
