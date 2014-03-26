using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using CustomOffice365OWA.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace CustomOffice365OWA.Controllers
{
    /// <summary>
    /// A common controller to handle authentication, SharePoint Service Discovery, 
    ///     and error-handling for Office 365 APIs.
    /// </summary>
    public class Office365CommonController : Controller
    {
        private static readonly string AppPrincipalId = ConfigurationManager.AppSettings["ida:ClientID"];
        private static readonly string AppKey = ConfigurationManager.AppSettings["ida:Password"];

        private const string OAuthUrl = "https://login.windows.net/{0}";
        private static readonly string AuthorizeUrl = string.Format(CultureInfo.InvariantCulture,
            OAuthUrl,
            "common/oauth2/authorize?response_type=code&client_id={0}&resource={1}&redirect_uri={2}&state={3}");

        /// <summary>
        /// Returns a URL through which a user can be authorized to access Office 365 APIs.
        /// After the authorization is complete, the user will be redirected back to the URL 
        ///     defined by the redirectTo parameter. This can be the same URL as the caller's URL
        ///     (e.g., Request.Url), or it can contain additional query-string parameters
        ///     (e.g., to restore state).
        /// </summary>
        internal static string GetAuthorizationUrl(string resourceId, Uri redirectTo)
        {
            HttpContext context = System.Web.HttpContext.Current;

            // To prevent Cross-Site Request Forgery attacks (http://tools.ietf.org/html/rfc6749 section 4.2.1),
            //     it is important to send a randomly-generated value as a state parameter.
            // This state parameter is saved in a cookie, so it can later be compared with the state
            //     parameter that we receive from the Authorization Server along with the Authorization Code.
            // The state cookie will also capture information about the resource ID and redirect-to URL,
            //     for use in the Index method (after the login page redirects back to this controller).
            Office365StateCookieInfo stateCookieInfo = new Office365StateCookieInfo
            {
                UniqueId = Guid.NewGuid().ToString(),
                ResourceId = resourceId,
                RedirectTo = redirectTo.ToString()
            };
            HttpCookie stateCookie = new HttpCookie(OAuthRequestStateCookiePrefix + stateCookieInfo.UniqueId)
            {
                HttpOnly = true,
                Secure = !context.Request.Url.IsLoopback,
                Value = JsonConvert.SerializeObject(stateCookieInfo),
                Expires = DateTime.Now.AddMinutes(10)
            };
            context.Response.Cookies.Add(stateCookie);

            // Create an OAuth request URL. To avoid introducing auth-related complexity into 
            //     individual controllers, the Office365CommonController will handle the entire auth flow,
            //     and only redirect to the original caller on completion.
            Uri redirectToThisController = new Uri(context.Request.Url, "/Office365Common");

            return String.Format(CultureInfo.InvariantCulture,
                AuthorizeUrl,
                Uri.EscapeDataString(AppPrincipalId),
                Uri.EscapeDataString(resourceId),
                Uri.EscapeDataString(redirectToThisController.ToString()),
                Uri.EscapeDataString(stateCookieInfo.UniqueId));
        }

        /// <summary>
        /// This method will be invoked as a call-back from an authentication service (e.g., https://login.windows.net/).
        /// It is not intended to be called directly, or to be called without first invoking the "GetAuthorizationUrl" method.
        /// On completion, the method will cache the refresh token and access tokens, and redirect to the URL
        ///     specified in the state cookie (created by the "GetAuthorizationUrl" method, with its unique ID 
        ///     included in the "state" of this method).
        /// </summary>
        public ActionResult Index(string code, string error, string error_description, string state)
        {
            // NOTE: In production, OAuth must be done over a secure HTTPS connection.
            if (Request.Url.Scheme != "https" && !Request.Url.IsLoopback)
            {
                const string message = "Invalid URL. Please run the app over a secure HTTPS connection.";
                return ShowErrorMessage(message, message + " URL: " + Request.Url.ToString());
            }

            // Ensure that there is a state cookie on the the request.
            HttpCookie stateCookie = Request.Cookies[OAuthRequestStateCookiePrefix + state];
            if (stateCookie == null)
            {
                Office365Cache.RemoveAllFromCache();
                const string message = "An authentication error has occurred. Please return to the previous page and try again.";
                string errorDetail = "Missing OAuth state cookie." + " URL: " + Request.Url.ToString();
                return ShowErrorMessage(message, errorDetail);
            }

            const string genericAuthenticationErrorMessage = "An authentication error has occurred.";

            // Retrieve the unique ID from the saved cookie, and compare it with the state parameter returned by 
            //     the Azure Active Directory Authorization endpoint:
            Office365StateCookieInfo stateCookieInfo = JsonConvert.DeserializeObject<Office365StateCookieInfo>(stateCookie.Value);
            if (stateCookieInfo.UniqueId != state)
            {
                // State is mismatched, error
                Office365Cache.RemoveAllFromCache();
                string errorDetail = "OAuth state cookie mismatch." + " URL: " + Request.Url.ToString();
                return ShowErrorMessage(genericAuthenticationErrorMessage, errorDetail);
            }

            // State check complete, clear the cookie:
            stateCookie.Expires = DateTime.Now.AddDays(-1);
            Response.Cookies.Set(stateCookie);

            // Handle error codes returned from the Authorization Server, if any:
            if (error != null)
            {
                Office365Cache.RemoveAllFromCache();
                return ShowErrorMessage(genericAuthenticationErrorMessage,
                    error + ": " + error_description + " URL: " + Request.Url.ToString());
            }

            // If still here, redeem the authorization code for an access token:
            try
            {
                ClientCredential credential = new ClientCredential(AppPrincipalId, AppKey);
                string authority = string.Format(CultureInfo.InvariantCulture, OAuthUrl, "common");
                AuthenticationContext authContext = new AuthenticationContext(authority);
                AuthenticationResult result = authContext.AcquireTokenByAuthorizationCode(
                    code, new Uri(Request.Url.GetLeftPart(UriPartial.Path)), credential);

                // Cache the access token and refresh token
                Office365Cache.GetAccessToken(stateCookieInfo.ResourceId).Value = result.AccessToken;
                Office365Cache.GetRefreshToken().Value = result.RefreshToken;

                // Also save the Tenant ID and User ID
                SaveInCache("TenantId", result.TenantId);
                SaveInCache("UserId", result.UserInfo.UserId);

                return Redirect(stateCookieInfo.RedirectTo);
            }
            catch (ActiveDirectoryAuthenticationException ex)
            {
                return ShowErrorMessage(genericAuthenticationErrorMessage,
                    "URL: " + Request.Url.ToString() + " Exception: " + ex.ToString());
            }
        }

        /// <summary>
        /// Send an HTTP request, with authorization. If the request fails due to an unauthorized exception,
        ///     this method will try to renew the access token in serviceInfo and try again.
        /// </summary>
        public static async Task<HttpResponseMessage> SendRequestAsync(
            Office365ServiceInfo serviceInfo, HttpClient client, Func<HttpRequestMessage> requestCreator)
        {
            using (HttpRequestMessage request = requestCreator.Invoke())
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", serviceInfo.AccessToken);
                request.Headers.UserAgent.Add(new ProductInfoHeaderValue(AppPrincipalId, String.Empty));
                HttpResponseMessage response = await client.SendAsync(request);

                // Check if the server responded with "Unauthorized". If so, it might be a real authorization issue, or 
                //     it might be due to an expired access token. To be sure, renew the token and try one more time:
                if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    Office365Cache.GetAccessToken(serviceInfo.ResourceId).RemoveFromCache();
                    serviceInfo.AccessToken = GetAccessTokenFromRefreshToken(serviceInfo.ResourceId);

                    // Create and send a new request:
                    using (HttpRequestMessage retryRequest = requestCreator.Invoke())
                    {
                        retryRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", serviceInfo.AccessToken);
                        retryRequest.Headers.UserAgent.Add(new ProductInfoHeaderValue(AppPrincipalId, String.Empty));
                        response = await client.SendAsync(retryRequest);
                    }
                }

                // Return either the original response, or the response from the second attempt:
                return response;
            }
        }

        /// <summary>
        /// A static method that routes errors to a single centralized error-handler.
        /// This method will attempt to extract a human-readable error from the response string,
        /// based on the the format of the data and the error handling scheme of the service.
        /// </summary>
        public static ActionResult ShowErrorMessage(Office365ServiceInfo serviceInfo, string responseString)
        {
            string message, errorDetails;
            try
            {
                message = serviceInfo.ParseErrorMessage(responseString);
                errorDetails = responseString;
            }
            catch (Exception e)
            {
                message = "An unexpected error has occurred.";
                errorDetails = "Exception when parsing response string: " + e.ToString() +
                    "\n\nResponse string was " + responseString;
            }
            return ShowErrorMessage(message, errorDetails);
        }

        /// <summary>
        /// A static method that routes errors to a single centralized error-handler. 
        /// Message is expected to be a human-readable string that can be displayed to the user.
        /// The error details can contains any additional details (e.g., for logging purposes).
        /// </summary>
        public static ActionResult ShowErrorMessage(string message, string errorDetails)
        {
            IController commonController = new DefaultControllerFactory().CreateController(
                System.Web.HttpContext.Current.Request.RequestContext,
                "Office365Common");
            return ((Office365CommonController)commonController).ShowErrorMessageHelper(message, errorDetails);
        }

        /// <summary>
        /// A common error handler, used by the two static "ShowErrorMessage" methods above.
        /// </summary>
        private ActionResult ShowErrorMessageHelper(string message, string errorDetails)
        {
            // TODO: You can customize this method to write the error detail to a database, 
            //       and/or to return a different error view.
            ViewBag.ErrorMessage = message;
            return View("Office365Error");
        }

        /// <summary>
        /// Clears any OAuth-related data, such as access and refresh tokens and state cookies.
        /// This method should be called as part of your application's logout routine.
        /// </summary>
        public static void ClearSession()
        {
            Office365Cache.RemoveAllFromCache();

            // Also remove the cookies used to store the OAuth request state:
            foreach (string cookieName in System.Web.HttpContext.Current.Request.Cookies.AllKeys)
            {
                if (cookieName.StartsWith(OAuthRequestStateCookiePrefix, StringComparison.Ordinal))
                {
                    System.Web.HttpContext.Current.Response.Cookies[cookieName].Expires = DateTime.Now.AddDays(-1);
                }
            }
        }

        /// <summary>
        /// Returns the Tenant ID of the signed-in user.  If not signed in, returns null.
        /// </summary>
        public static string TenantId
        {
            get { return (string)GetFromCache("TenantId"); }
        }

        /// <summary>
        /// Returns the User ID of the signed-in user.  If not signed in, returns null.
        /// </summary>
        public static string UserId
        {
            get { return (string)GetFromCache("UserId"); }
        }

        /// <summary>
        /// Obtains an access token for the specified resource, using a cached access token or a refresh token. 
        /// If successful, the token will be cached for future use.
        /// On failure, this method will return null to signify that the caller must do an OAuth redirect instead.
        /// </summary>
        internal static string GetAccessToken(string resourceId)
        {
            // Try the cache first:
            string accessToken = Office365Cache.GetAccessToken(resourceId).Value;
            if (accessToken != null)
            {
                return accessToken;
            }

            // If there is no Access Token in the cache for this resource, check if there is a refresh token 
            //    in the cache that can be used to get a new access token.
            accessToken = GetAccessTokenFromRefreshToken(resourceId);
            if (accessToken != null)
            {
                return accessToken;
            }

            // If neither succeeded, return null to signal a need for an OAuth redirect.
            return null;
        }

        internal static void SaveInCache(string name, object value)
        {
            Office365Cache.SaveInCache(name, value);
        }

        /// <summary>
        /// If the item exists, returns the saved value; otherwise, returns null.
        /// </summary>
        internal static object GetFromCache(string name)
        {
            return Office365Cache.GetFromCache(name);
        }

        #region Private helpers

        private const string OAuthRequestStateCookiePrefix = "WindowsAzureActiveDirectoryOAuthRequestState#";

        /// <summary>
        /// Try to get a new access token for this resource using a refresh token.
        /// If successful, this method will cache the access token for future use.
        /// If this fails, return null, signaling the caller to do the OAuth redirect.
        /// </summary>
        private static string GetAccessTokenFromRefreshToken(string resourceId)
        {
            string refreshToken = Office365Cache.GetRefreshToken().Value;
            if (refreshToken == null)
            {
                // If no refresh token, the caller will need to send the user to do an OAuth redirect.
                return null;
            }

            // Redeem the refresh token for an access token:
            try
            {
                ClientCredential credential = new ClientCredential(AppPrincipalId, AppKey);
                string authority = string.Format(CultureInfo.InvariantCulture, OAuthUrl, "common");
                AuthenticationContext authContext = new AuthenticationContext(authority);
                AuthenticationResult result = authContext.AcquireTokenByRefreshToken(
                    refreshToken, AppPrincipalId, credential, resourceId);

                // Cache the access token and update the refresh token:
                Office365Cache.GetAccessToken(resourceId).Value = result.AccessToken;
                Office365Cache.GetRefreshToken().Value = result.RefreshToken;

                return result.AccessToken;
            }
            catch (ActiveDirectoryAuthenticationException)
            {
                // Forget the refresh token and return null, so as to start the OAuth redirect from scratch.
                Office365Cache.GetRefreshToken().RemoveFromCache();
                return null;
            }
        }

        #endregion

        #region Private classes

        /// <summary>
        /// Data structure for holding the Office365 state in a cookie during an Authentication request.
        /// </summary>
        private class Office365StateCookieInfo
        {
            public string UniqueId { get; set; }
            public string ResourceId { get; set; }
            public string RedirectTo { get; set; }
        }

        /// <summary>
        /// A default cache implementation that uses Session for storing and retrieving data related to Office 365 APIs,
        ///     such as access and refresh tokens, and dynamically-discovered items like the tenant ID or API endpoints.
        /// If needed, you can save the data to more persistent storage, such as a database.
        /// </summary>
        private class Office365Cache
        {
            private const string CachePrefix = "Office365Cache#";

            private Office365Cache() { }

            public static Office365CacheEntry GetAccessToken(string resourceId)
            {
                return new Office365CacheEntry("AccessToken#" + resourceId);
            }

            public static Office365CacheEntry GetRefreshToken()
            {
                return new Office365CacheEntry("RefreshToken");
            }

            internal static void SaveInCache(string name, object value)
            {
                System.Web.HttpContext.Current.Session[CachePrefix + name] = value;
            }

            internal static object GetFromCache(string name)
            {
                return System.Web.HttpContext.Current.Session[CachePrefix + name];
            }

            internal static void RemoveFromCache(string name)
            {
                System.Web.HttpContext.Current.Session.Remove(CachePrefix + name);
            }

            internal static void RemoveAllFromCache()
            {
                List<string> keysToRemove = new List<string>();
                foreach (string key in System.Web.HttpContext.Current.Session.Keys)
                {
                    if (key.StartsWith(CachePrefix, StringComparison.Ordinal))
                    {
                        keysToRemove.Add(key);
                    }
                }

                foreach (string key in keysToRemove)
                {
                    RemoveFromCache(key);
                }
            }
        }

        private class Office365CacheEntry
        {
            private readonly string _name;

            public Office365CacheEntry(string name)
            {
                _name = name;
            }

            public string Value
            {
                get
                {
                    return (string)Office365Cache.GetFromCache(_name);
                }
                set
                {
                    Office365Cache.SaveInCache(_name, value);
                }
            }

            public void RemoveFromCache()
            {
                Office365Cache.RemoveFromCache(_name);
            }
        }

        #endregion
    }
}