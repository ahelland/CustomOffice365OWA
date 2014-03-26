using Newtonsoft.Json.Linq;
using CustomOffice365OWA.Controllers;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CustomOffice365OWA.Models
{
    public abstract class Office365ServiceInfo
    {
        /// <summary>
        /// The API endpoint, with no trailing slash.
        /// </summary>
        public string ApiEndpoint { get; set; }

        /// <summary>
        /// Access token (or null), used for authenticating to the API endpoint. 
        /// </summary>
        public string AccessToken { get; set; }

        /// <summary>
        /// The Resource ID for the service, used to obtain a new access token or to retrieve an existing token from cache.
        /// </summary>
        public string ResourceId { get; set; }

        /// <summary>
        /// Checks whether the access token is valid (e.g., non-empty). If not valid, the caller 
        /// should issue an OAuth request using the Office365CommonController.GetAuthorizationUrl function
        /// </summary>
        public Boolean HasValidAccessToken
        {
            get
            {
                return (AccessToken != null);
            }
        }

        /// <summary>
        /// Extracts a human-readable error message from the response string. If the format is not recognized,
        /// or if the path is not of the expected form, an exception can be thrown.
        /// </summary>
        internal abstract string ParseErrorMessage(string responseString);

        /// <summary>
        /// Returns a URL through which a user can be authorized to access Office 365 APIs.
        /// After the authorization is complete, the user will be redirected back to the URL 
        ///     defined by the redirectTo parameter. This can be the same URL as the caller's URL
        ///     (e.g., Request.Url), or it can contain additional query-string parameters
        ///     (e.g., to restore state).
        /// </summary>
        public virtual string GetAuthorizationUrl(Uri redirectTo)
        {
            return Office365CommonController.GetAuthorizationUrl(ResourceId, redirectTo);
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////
        /////////////////////////// FACTORY METHODS FOR RETRIEVING SERVICES ///////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////

        /// <summary>
        /// Returns information about the Active Directory service, including its cached access token.
        /// If the access token is null, the caller must do an OAuth redirect to the URL 
        ///     obtained using Office365CommonController's GetAuthorizationUrl function.
        /// </summary>
        public static Office365ServiceInfo GetActiveDirectoryServiceInfo()
        {
            return new ActiveDirectoryServiceInfo();
        }

        /// <summary>
        /// Returns information about the Exchange service, including its cached access token.
        /// If the access token is null, the caller must do an OAuth redirect to the URL 
        ///     obtained using Office365CommonController's GetAuthorizationUrl function.
        /// </summary>
        public static Office365ServiceInfo GetExchangeServiceInfo()
        {
            return new ExchangeServiceInfo();
        }

        /// <summary>
        /// Returns information about the SharePoint service, including its cached access token.
        /// Note that for SharePoint, the resource ID and API endpoint will be different for each tenant,
        ///     so that information must be discovered via a Discovery Service before it can be cached.
        /// Because of this, this method is async.
        /// If the access token is null, the caller must do an OAuth redirect to the URL 
        ///     obtained using Office365CommonController's GetAuthorizationUrl function.
        /// </summary>
        public static async Task<Office365ServiceInfo> GetSharePointOneDriveServiceInfoAsync()
        {
            return await SharePointOneDriveServiceInfo.CreateAsync();
        }

        #region Private helper methods

        /// <summary>
        /// Determines the format of the response string, and extract a human-readable error message
        /// from a response string.
        /// </summary>
        private static string GetErrorMessage(string responseString, string[] jsonErrorPath, string[] xmlErrorPath)
        {
            switch (responseString.TrimStart().FirstOrDefault())
            {
                case '{':
                    return ParseJsonErrorMessage(jsonErrorPath, responseString);
                case '<':
                    return ParseXmlErrorMessage(xmlErrorPath, responseString);
                default:
                    throw new ArgumentException("Unrecognized format for the response.");
            }
        }

        private static string ParseJsonErrorMessage(string[] path, string responseString)
        {
            JToken currentJsonNode = JObject.Parse(responseString);
            foreach (string nodeName in path)
            {
                currentJsonNode = currentJsonNode[nodeName];
            }
            return currentJsonNode.Value<string>();
        }

        private static string ParseXmlErrorMessage(string[] path, string responseString)
        {
            using (StringReader reader = new StringReader(responseString))
            {
                XDocument xmlDoc = XDocument.Load(reader);
                XNamespace xmlNamespace = xmlDoc.Root.Name.Namespace;
                XElement currentXmlNode = xmlDoc.Root;
                if (xmlDoc.Root.Name.LocalName != path.First())
                {
                    throw new Exception("Unexpected root node name: " + xmlDoc.Root.Name.LocalName);
                }
                foreach (string nodeName in path.Skip(1))
                {
                    currentXmlNode = currentXmlNode.Element(xmlNamespace + nodeName);
                }
                return currentXmlNode.Value;
            }
        }

        #endregion

        #region Private classes

        private class ActiveDirectoryServiceInfo : Office365ServiceInfo
        {
            internal ActiveDirectoryServiceInfo()
            {
                // For Active Directory, the resource ID and API Endpoint are static for the public O365 cloud.
                ResourceId = "https://graph.windows.net/";
                ApiEndpoint = "https://graph.windows.net";
                AccessToken = Office365CommonController.GetAccessToken(ResourceId);
            }

            internal override string ParseErrorMessage(string responseString)
            {
                string[] jsonErrorPath = { "odata.error", "message", "value" };
                string[] xmlErrorPath = { "error", "message" };
                return GetErrorMessage(responseString, jsonErrorPath, xmlErrorPath);
            }
        }

        private class ExchangeServiceInfo : Office365ServiceInfo
        {
            internal ExchangeServiceInfo()
            {
                // For Exchange, the resource ID and API Endpoint are static for the public O365 cloud.
                ResourceId = "https://outlook.office365.com/";
                ApiEndpoint = "https://outlook.office365.com/ews/odata";
                AccessToken = Office365CommonController.GetAccessToken(ResourceId);
            }

            internal override string ParseErrorMessage(string responseString)
            {
                string[] jsonErrorPath = { "error", "message" };
                string[] xmlErrorPath = { "error", "message" };
                return GetErrorMessage(responseString, jsonErrorPath, xmlErrorPath);
            }
        }

        private class SharePointOneDriveServiceInfo : Office365ServiceInfo
        {
            /// <summary>
            /// This constructor is intentionally private, and should not be used.
            /// Instead, callers should create a new instance by calling the static CreateAsync() method
            /// </summary>
            private SharePointOneDriveServiceInfo() {}

            /// <summary>
            /// Discovery service, if needed. If the Resource ID and API Endpoint are already known,
            ///     this field may remain set to null.
            /// </summary>
            private DiscoveryServiceInfo _discoveryServiceInfo;

            internal static async Task<Office365ServiceInfo> CreateAsync()
            {
                // Attempt to build an Office365ServiceInfo object based on cached API endpoint & resource ID information:
                SharePointOneDriveServiceInfo info = new SharePointOneDriveServiceInfo
                {
                    ResourceId = (string) Office365CommonController.GetFromCache("SharePointOneDriveResourceId"),
                    ApiEndpoint = (string) Office365CommonController.GetFromCache("SharePointOneDriveApiEndpoint")
                };

                // If the Resource ID and API Endpoint are not empty, then the cached information is sufficient:
                if (info.ResourceId != null && info.ApiEndpoint != null)
                {
                    info.AccessToken = Office365CommonController.GetAccessToken(info.ResourceId);
                    return info;
                }

                // If did not return above, invoke the Discovery Service to obtain the resource ID and API endpoint:
                info._discoveryServiceInfo = new DiscoveryServiceInfo();

                // If no auth header is available for Discovery, return the info as is, with the missing 
                //     access token (and possibly a missing ResourceId and ApiEndpoint as well). The caller will need
                //     to do an OAuth redirect anyway.
                if (!info._discoveryServiceInfo.HasValidAccessToken)
                {
                    return info;
                }

                // If still here, discovery has enough information to obtain the SharePoint OneDrive endpoints:
                DiscoveryResult[] results = await info._discoveryServiceInfo.DiscoverServicesAsync();
                DiscoveryResult myFilesEndpoint = results.First(result => result.Capability == "MyFiles");

                // Update and cache the resource ID and API endpoint:
                info.ResourceId = myFilesEndpoint.ServiceResourceId;
                // NOTE: In the initial Preview release of Service Discovery, the "MyFiles" endpoint URL will always
                //     start with something like "https://contoso-my.sharepoint.com/personal/<username>_contoso_com/_api",
                //     but the path following "/_api" may change over time.  For consistency, it is safer to manually
                //     extract the root path, and then append a call for the location of the Documents folder:
                info.ApiEndpoint = myFilesEndpoint.ServiceEndpointUri.Substring(
                    0, myFilesEndpoint.ServiceEndpointUri.IndexOf("/_api", StringComparison.Ordinal)) +
                    "/_api/web/getfolderbyserverrelativeurl('Documents')";
                Office365CommonController.SaveInCache("SharePointOneDriveResourceId", info.ResourceId);
                Office365CommonController.SaveInCache("SharePointOneDriveApiEndpoint", info.ApiEndpoint);
                info.AccessToken = Office365CommonController.GetAccessToken(info.ResourceId);
                return info;
            }

            public override string GetAuthorizationUrl(Uri redirectTo)
            {
                // If the Resource ID is known, use it. Otherwise, get authorized to the discovery service, first.
                string resourceIdForAuthRequest = ResourceId ?? _discoveryServiceInfo.ResourceId;
                return Office365CommonController.GetAuthorizationUrl(resourceIdForAuthRequest, redirectTo);
            }

            internal override string ParseErrorMessage(string responseString)
            {
                string[] jsonErrorPath = { "error", "message", "value" };
                string[] xmlErrorPath = { "error", "message" };
                return GetErrorMessage(responseString, jsonErrorPath, xmlErrorPath);
            }
        }

        private class DiscoveryServiceInfo : Office365ServiceInfo
        {
            internal DiscoveryServiceInfo()
            {
                // In this initial Preview release, you must use a temporary Resource ID for Service Discovery ("Microsoft.SharePoint").
                // TODO: If this Resource ID ceases to work, check for an updated value at http://go.microsoft.com/fwlink/?LinkID=392944
                ResourceId = "Microsoft.SharePoint";

                ApiEndpoint = "https://api.office.com/discovery/me";
                AccessToken = Office365CommonController.GetAccessToken(ResourceId);
            }

            internal override string ParseErrorMessage(string responseString)
            {
                // Discovery is not a user-facing service, and should not be returning an error message. 
                return "An error occurred during service discovery.";
            }

            /// <summary>
            /// Returns information obtained via Discovery. Will throw an exception on error.
            /// </summary>
            internal async Task<DiscoveryResult[]> DiscoverServicesAsync()
            {
                // Create a URL for retrieving the data:
                string requestUrl = String.Format(CultureInfo.InvariantCulture,
                    "{0}/services",
                    ApiEndpoint);

                // Prepare the HTTP request:
                using (HttpClient client = new HttpClient())
                {
                    Func<HttpRequestMessage> requestCreator = () =>
                    {
                        HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                        request.Headers.Add("Accept", "application/json;odata=verbose");
                        return request;
                    };

                    // Send the request using a helper method, which will add an authorization header to the request,
                    // and automatically retry with a new token if the existing one has expired.
                    using (HttpResponseMessage response = await Office365CommonController.SendRequestAsync(
                        this, client, requestCreator))
                    {
                        // Read the response and deserialize the data:
                        string responseString = await response.Content.ReadAsStringAsync();
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new Exception("Could not obtain discovery information. Service returned " +
                                response.StatusCode + ":\n\n" + responseString);
                        }

                        // If successful, return the discovery results
                        return JObject.Parse(responseString)["d"]["results"].ToObject<DiscoveryResult[]>();
                    }
                }
            }
        }

        /// <summary>
        /// A private class for de-serializing service entries returned by the Discovery Service
        /// </summary>
        private class DiscoveryResult
        {
            public string Capability { get; set; }
            public string ServiceEndpointUri { get; set; }
            public string ServiceResourceId { get; set; }
        }

        #endregion
    }
}