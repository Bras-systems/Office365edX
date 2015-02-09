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
using SalesManager.Utils;

namespace SalesManager.Filters
{
	public class TokenFilterAttribute : FilterAttribute, IActionFilter
	{
		String discoveryEndpoint = SettingsHelper.DiscoveryEndpoint;
		AuthenticationContext authenticationContext = SettingsHelper.AuthenticationContext;
		ClientCredential clientCredential = SettingsHelper.ClientCredential;
		String discoResource = ConfigurationManager.AppSettings["discoResource"];


		// This will examine the query string to ensure a code exists for authentication
		// If the code does not exist, it will redirect to Azure Active Directory to handle the authentication
		public async void OnActionExecuting(ActionExecutingContext filterContext)
		{
			var request = filterContext.HttpContext.Request;
			
			String code = request.QueryString["code"];

			if(String.IsNullOrEmpty(code)) {
				RedirectToAuthentication(filterContext);
				return;
			}

			var discoveryClient = LoadDiscoveryClientAsync(filterContext, code);
			var contactsDiscoveryResult = 
				await LoadContactsDiscoveryResultAsync(filterContext, discoveryClient);
			LoadOutlookServicesClient(filterContext, contactsDiscoveryResult, code);
		}

		private void RedirectToAuthentication(ActionExecutingContext filterContext)
		{
			Uri redirectUri =
				SettingsHelper.AuthenticationContext.GetAuthorizationRequestURL(
					discoResource, // Where the discovery service is
					SettingsHelper.ClientCredential.ClientId, // This identifies the application
					new Uri(filterContext.HttpContext.Request.Url.AbsoluteUri.Split('?')[0]), // Redirect URL
					UserIdentifier.AnyUser,
					string.Empty
				);

			filterContext.Result = new RedirectResult(redirectUri.ToString()); // Redirect to the login page
		}

		private DiscoveryClient LoadDiscoveryClientAsync(ActionExecutingContext filterContext, String code)
		{
			var request = filterContext.HttpContext.Request;
			var session = filterContext.HttpContext.Session;

			DiscoveryClient discoveryClient = session["DiscoveryClient"] as DiscoveryClient;

			// see if the discovery client has already been loaded
			if (discoveryClient == null) {
				discoveryClient = new DiscoveryClient(new Uri(discoveryEndpoint), async () => {
					var authResult = await authenticationContext.AcquireTokenByAuthorizationCodeAsync(
						code, // User Token
						new Uri(request.Url.AbsoluteUri.Split('?')[0]), // Where to route the user back to after the request
						clientCredential // credentials for the application
					);

					return authResult.AccessToken;
				});

				session["DiscoveryClient"] = discoveryClient;
			}

			return discoveryClient;
		}

		private async Task<CapabilityDiscoveryResult> LoadContactsDiscoveryResultAsync(ActionExecutingContext filterContext, DiscoveryClient discoveryClient)
		{
			var contactsDiscoveryResult = filterContext.HttpContext.Session["ContactsDiscoveryResult"] as CapabilityDiscoveryResult;
			if (contactsDiscoveryResult == null) {
				contactsDiscoveryResult = await discoveryClient.DiscoverCapabilityAsync("Contacts");
				filterContext.HttpContext.Session["ContactsDiscoveryResult"] = contactsDiscoveryResult;
			}
			return contactsDiscoveryResult;
		}

		private OutlookServicesClient LoadOutlookServicesClient(ActionExecutingContext filterContext, CapabilityDiscoveryResult contactsDiscoveryResult, String code)
		{
			var outlookServicesClient = filterContext.HttpContext.Session["OutlookServicesClient"] as OutlookServicesClient;
			if (outlookServicesClient == null) {
				outlookServicesClient = new OutlookServicesClient(contactsDiscoveryResult.ServiceEndpointUri, async () => {
					var authResult = await authenticationContext.AcquireTokenByAuthorizationCodeAsync(
						code,
						new Uri(filterContext.HttpContext.Request.Url.AbsoluteUri.Split('?')[0]),
						clientCredential);
					return authResult.AccessToken;
				});
				filterContext.HttpContext.Session["OutlookServicesClient"] = outlookServicesClient;
			}

			return outlookServicesClient;
		}

		public void OnActionExecuted(ActionExecutedContext filterContext)
		{
			// no work needed here
		}
	}
}