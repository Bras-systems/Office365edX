using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using CalendarClient.Utils;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;

namespace CalendarClient.Controllers
{
    public class CalendarController : Controller
    {
		public async Task<ActionResult> Index()
		{
			// fetch user information from claims
			var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
			var userObjectUniqueId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;

			// discover contact endpoint
			var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret);
			var userIdentifier = new UserIdentifier(userObjectUniqueId, UserIdentifierType.UniqueId);

			// create authentication context
			AuthenticationContext authenticationContext = new AuthenticationContext(SettingsHelper.AzureADAuthority, new EFADALTokenCache(signInUserId));

			// create O365 discovery client 
			DiscoveryClient discovery = new DiscoveryClient(new Uri(SettingsHelper.O365DiscoveryServiceEndpoint),
			  async () => {
				  var authenticationResult = await authenticationContext.AcquireTokenSilentAsync(SettingsHelper.O365DiscoveryResourceId, clientCredential, userIdentifier);

				  return authenticationResult.AccessToken;
			  });

			// obtain the user's calendar api endpoint from the DiscoveryClient
			var discoveryCapabilityResult = await discovery.DiscoverCapabilityAsync("Calendar");

			// create Outlook client using the calendar api endpoint
			OutlookServicesClient client = new OutlookServicesClient(discoveryCapabilityResult.ServiceEndpointUri,
			  async () => {
				  var authResult = await authenticationContext.AcquireTokenSilentAsync(discoveryCapabilityResult.ServiceResourceId, clientCredential,
				  userIdentifier);

				  return authResult.AccessToken;
			  });

			// get the first 20 events for the user
			var results = await client.Me.Events.Take(20).ExecuteAsync();

			// pass the events into the ViewBag to send to the View
			ViewBag.Events = results.CurrentPage.OrderBy(c => c.Start);

			return View();
		}
    }
}