using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using SalesClient.Utils;

namespace SalesClient.Models
{
	public class FilesRepository
	{
		public static async Task<string> GetAccessToken()
		{
			// fetch from stuff user claims
			var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
			var userObjectId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;

			// discover contact endpoint
			var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret);
			var userIdentifier = new UserAssertion(userObjectId);

			// create auth context
			AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.AzureADAuthority, new EFADALTokenCache(signInUserId));

			// authenticate
			var authResult = await authContext.AcquireTokenAsync(SettingsHelper.SharePointServiceResourceId, clientCredential, userIdentifier);

			// obtain access token
			return authResult.AccessToken;
		}
	}
}
