using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Office365.OutlookServices;
using CalendarDemo.Utils;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;

namespace CalendarDemo.Models
{
	public class CalenderItemsRepository
	{

		private async Task<OutlookServicesClient> EnsureClientCreated()
		{
			// fetch from stuff user claims
			var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
			var userObjectId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;

			// discover contact endpoint
			var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret);
			var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);

			// create auth context
			AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.AzureADAuthority, new EFADALTokenCache(signInUserId));

			// create O365 discovery client 
			DiscoveryClient discovery = new DiscoveryClient(new Uri(SettingsHelper.O365DiscoveryServiceEndpoint),
			  async () => {
				  var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.O365DiscoveryResourceId, clientCredential, userIdentifier);

				  return authResult.AccessToken;
			  });

			// query discovery service for endpoint for 'calendar' endpoint
			CapabilityDiscoveryResult dcr = await discovery.DiscoverCapabilityAsync("Calendar");

			// create an OutlookServicesclient
			return new OutlookServicesClient(dcr.ServiceEndpointUri,
			  async () => {
				  var authResult =
					await
					  authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId, clientCredential, userIdentifier);
				  return authResult.AccessToken;
			  });
		}

		internal async Task<List<CalendarItem>> GetItems()
		{
			var context = await EnsureClientCreated();
			return null;
		}
	}
}