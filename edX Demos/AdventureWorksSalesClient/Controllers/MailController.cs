using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using AdventureWorksSalesClient.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using SalesClient.Utils;

namespace AdventureWorksSalesClient.Controllers
{
    public class MailController : Controller
    {
        // GET: Mail
        public ActionResult SendMessage(string email, string code)
        {
            ViewMessage message = new ViewMessage();
			message.Email = email;
			ViewBag.Code = code;
			return View(message);
        }

		[HttpPost()]
		[ValidateAntiForgeryToken()]
		public async Task<ActionResult> SendMessage(string code, ViewMessage message)
		{
			#region Get OutlookServicesClient

			AuthenticationContext authContext = new AuthenticationContext(
												 ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
												 true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			// Check the cache for the client
			// This is the client that will give us access to all of the "Outlook" information
			OutlookServicesClient outlookClient = Helpers.GetFromCache("OutlookClient") as OutlookServicesClient;

			//Get an authorization code if needed
			if (outlookClient == null && cdr != null && code == null) {
				Uri redirectUri = authContext.GetAuthorizationRequestURL(
					cdr.ServiceResourceId,
					creds.ClientId,
					new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
					UserIdentifier.AnyUser,
					string.Empty);

				return Redirect(redirectUri.ToString());
			}

			//Create the OutlookServicesClient
			if (outlookClient == null && cdr != null && code != null) {
				outlookClient = new OutlookServicesClient(cdr.ServiceEndpointUri, async () => {
					var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
						code,
						new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
						creds);
					return authResult.AccessToken;
				});

				Helpers.SaveInCache("OutlookClient", outlookClient);
			}

			#endregion

			Message outlookMessage = new Message();
			outlookMessage.Body = new ItemBody() {
										ContentType = BodyType.Text,
										Content = message.Body
									};
			outlookMessage.ToRecipients.Add(new Recipient() 
										{ 
											EmailAddress = new EmailAddress() {
												 Address = message.Email,
												 Name = message.Email
											}
										});
			await outlookClient.Me.SendMailAsync(outlookMessage, true);

			TempData["MessageEmail"] = message.Email;

			return RedirectToAction("Index", "Contacts", new {code = code});
		}
    }
}