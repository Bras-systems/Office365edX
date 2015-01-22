using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using SalesClient.Models;
using SalesClient.Utils;

namespace SalesClient.Controllers
{
	public class HomeController : Controller
	{
		private const string spSite = "https://geektrainerdev.sharepoint.com";
		private const string discoResource = "https://api.office.com/discovery/";
		private const string discoEndpoint = "https://api.office.com/discovery/v1.0/me/";		// Default

		// Add in a code parameter. The code will store the user's token
		// Mark method as async and return Task<ActionResult>
		// This allows us to use the await keyword and make async calls
		public async Task<ActionResult> Index(String code)
		{
			AuthenticationContext authContext = new AuthenticationContext(
			   ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
			   true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			// check if the DiscoveryClient is already in the Session object
			DiscoveryClient disco = Helpers.GetFromCache("DiscoveryClient") as DiscoveryClient;

			// If we don't have the discovery client and a code, then we know the user has not authenticated
			// Redirect the user to the authentication page
			if (disco == null && code == null) {
				Uri redirectUri = authContext.GetAuthorizationRequestURL(
					discoResource, // Where the discovery service is
					creds.ClientId, // This identifies the application
					new Uri(Request.Url.AbsoluteUri.Split('?')[0]), // Redirect URL
					UserIdentifier.AnyUser,
					string.Empty);

				return Redirect(redirectUri.ToString()); // Redirect to the login page
			}

			if(code == null) code = await FilesRepository.GetAccessToken();

			// Token, but no discovery object
			// User has authenticated, but we haven't connected to Office 365 to see what the user can do
			if (disco == null && code != null) {
				disco = new DiscoveryClient(new Uri(discoEndpoint), async () => {
					var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
						code, // User Token
						new Uri(Request.Url.AbsoluteUri.Split('?')[0]), // Where to route the user back to after the request
						creds // credentials for the application
					);

					return authResult.AccessToken;
				});
			}

			//Discover required capabilities
			CapabilityDiscoveryResult contactsDisco =
				await disco.DiscoverCapabilityAsync("Contacts"); // Contacts
			CapabilityDiscoveryResult filesDisco =
				await disco.DiscoverCapabilityAsync("MyFiles"); // OneDrive for Business

			Helpers.SaveInCache("ContactsDiscoveryResult", contactsDisco);
			Helpers.SaveInCache("FilesDiscoveryResult", filesDisco);

			List<ViewDiscovery> discoveries = new List<ViewDiscovery>(){
                new ViewDiscovery(){
                    Capability = "Contacts",
                    EndpointUri = contactsDisco.ServiceEndpointUri.OriginalString,
                    ResourceId = contactsDisco.ServiceResourceId,
                    Version = contactsDisco.ServiceApiVersion
                },
                new ViewDiscovery(){
                    Capability = "My Files",
                    EndpointUri = filesDisco.ServiceEndpointUri.OriginalString,
                    ResourceId = filesDisco.ServiceResourceId,
                    Version = filesDisco.ServiceApiVersion
                }
            };

			return View(discoveries);
		}

		public async Task<ActionResult> Contact(String code)
		{
			AuthenticationContext authContext = new AuthenticationContext(
						 ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
						 true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
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
					CompanyName = contact.CompanyName,
					EmailAddress = contact.EmailAddresses.FirstOrDefault().Address,
					BusinessPhone = contact.BusinessPhones.FirstOrDefault(),
					HomePhone = contact.HomePhones.FirstOrDefault()
				});
			}

			//Show the contacts
			return View(contacts);
		}

		[HttpGet()]
		public ActionResult SendEmail(String code)
		{
			AuthenticationContext authContext = new AuthenticationContext(
				ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
				true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
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

			return View();
		}
		
		[HttpPost()]
		public async Task<ActionResult> SendEmail(String code, String toAddress, String body)
		{
			AuthenticationContext authContext = new AuthenticationContext(
							 ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
							 true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
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

			Recipient recipient = new Recipient();
			recipient.EmailAddress = new EmailAddress() { Address = toAddress};

			List<Recipient> toRecipients = new List<Recipient>();
			toRecipients.Add(recipient);

			ItemBody itemBody = new ItemBody() { Content = body };

			Message message = new Message() {
				ToRecipients = toRecipients,
				Body = itemBody
			};

			await outlookClient.Me.SendMailAsync(message, true);

			ViewBag.Message = "Message Sent!!";
			return View();
		}

		[HttpGet()]
		public ActionResult CreateAppointment(String code)
		{
			AuthenticationContext authContext = new AuthenticationContext(
				ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
				true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
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

			return View();
		}

		[HttpPost()]
		public async Task<ActionResult> CreateAppointment(String code, String toAddress, String body)
		{
			AuthenticationContext authContext = new AuthenticationContext(
							 ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
							 true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
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

			Event appointment = new Event();
			
			List<Attendee> attendees = new List<Attendee>();
			attendees.Add(new Attendee() { 
								EmailAddress = new EmailAddress() { Address = toAddress},
								Type = AttendeeType.Required
			});
			appointment.Attendees = attendees;

			appointment.Body = new ItemBody() { Content = body };

			appointment.Start = new DateTimeOffset(DateTime.Now.AddHours(1));
			appointment.End = new DateTimeOffset(DateTime.Now.AddHours(2));

			await outlookClient.Me.Calendar.Events.AddEventAsync(appointment);

			ViewBag.Message = "Appointment Created!!";
			return View();
		}

	}
}