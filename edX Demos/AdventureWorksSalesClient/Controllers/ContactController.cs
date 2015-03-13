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
    public class ContactController : Controller
    {
        // GET: Contact
        public async Task<ActionResult> Index(string code)
        {
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

			//Get the contacts
			var contactsResults = await outlookClient.Me.Contacts.ExecuteAsync();
			List<ViewContact> contacts = new List<ViewContact>();

			foreach (var contact in contactsResults.CurrentPage.OrderBy(c => c.Surname)) {
				contacts.Add(new ViewContact {
					Id = contact.Id,
					GivenName = contact.GivenName,
					Surname = contact.Surname,
					DisplayName = contact.Surname + ", " + contact.GivenName,
					Email = (contact.EmailAddresses.FirstOrDefault() == null) ? 
								String.Empty : contact.EmailAddresses.FirstOrDefault().Address
				});
			}

			ViewBag.Code = code;

			//Show the contacts
			return View(contacts);
		}

		// Displays create form
		public ActionResult Create(string code)
		{
			ViewBag.Code = code;
			return View();
		}

		// Take the user data and create a new contact
		[HttpPost()]
		[ValidateAntiForgeryToken()]
		public async Task<ActionResult> Create(string code, ViewContact contact)
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

			// Create Outlook Contact
			Contact outlookContact = new Contact() {
				GivenName = contact.GivenName,
				Surname = contact.Surname,
				DisplayName = contact.DisplayName,
				EmailAddresses = new List<EmailAddress> () {
					new EmailAddress() { Address = contact.Email, Name = contact.DisplayName }
				}
			};

			// Save to Office 365
			await outlookClient.Me.Contacts.AddContactAsync(outlookContact);

			TempData["DisplayName"] = contact.DisplayName;
			return RedirectToAction("Index");
		}
    }
}