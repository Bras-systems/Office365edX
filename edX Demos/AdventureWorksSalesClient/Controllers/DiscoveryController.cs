using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using SalesClient.Models;
using SalesClient.Utils;

namespace AdventureWorksSalesClient.Controllers
{
    public class DiscoveryController : Controller
    {
		private const string discoResource = "https://api.office.com/discovery/";
		private const string discoEndpoint = "https://api.office.com/discovery/v1.0/me/";		// Default

        // GET: Discovery
        public async Task<ActionResult> Index(string code)
        {
			// Where do we send auth requests to?
			AuthenticationContext authContext = new AuthenticationContext(
			   ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
			   true);

			// wrapper for the id and password for this application
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

			ViewBag.Code = code;
			return View(discoveries);
		}
    }
}