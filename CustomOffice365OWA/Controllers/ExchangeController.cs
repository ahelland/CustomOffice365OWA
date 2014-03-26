using CustomOffice365OWA.Models;
using System;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace CustomOffice365OWA.Controllers
{
    [RoutePrefix("Mail")]
    public class ExchangeController : Controller
    {
        [HttpGet]
        [Route("Index")]
        public async Task<ActionResult> Index()
        {
            // Obtain information for communicating with the service:
            Office365ServiceInfo serviceInfo = Office365ServiceInfo.GetExchangeServiceInfo();
            if (!serviceInfo.HasValidAccessToken)
            {
                return Redirect(serviceInfo.GetAuthorizationUrl(Request.Url));
            }

            string requestUrl = String.Format(CultureInfo.InvariantCulture,
                "{0}/Me/Inbox/Messages",
                serviceInfo.ApiEndpoint
                );

            // Prepare the HTTP request:
            using (HttpClient client = new HttpClient())
            {
                Func<HttpRequestMessage> requestCreator = () =>
                {
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, requestUrl);
                    request.Headers.Add("Accept", "application/json;odata.metadata=full");
                    return request;
                };

                // Send the request using a helper method, which will add an authorization header to the request,
                // and automatically retry with a new token if the existing one has expired.
                using (HttpResponseMessage response = await Office365CommonController.SendRequestAsync(
                    serviceInfo, client, requestCreator))
                {
                    // Read the response and deserialize the data:
                    string responseString = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode)
                    {
                        return Office365CommonController.ShowErrorMessage(serviceInfo, responseString);
                    }

                    var messageContext = Newtonsoft.Json.JsonConvert.DeserializeObject<Message_odataContext>(responseString);

                    var messages = messageContext.Messages;

                    return View(messages);
                }
            }
        }
	}
}