﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.SessionState;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;

namespace SalesManager.Utils
{
	public class ServicesHelper
	{
		readonly String discoveryEndpoint = SettingsHelper.DiscoveryEndpoint;
		readonly AuthenticationContext authenticationContext = SettingsHelper.AuthenticationContext;
		readonly ClientCredential clientCredential = SettingsHelper.ClientCredential;
		readonly String discoResource = ConfigurationManager.AppSettings["discoResource"];
		readonly HttpContextBase httpContext = null;
		readonly HttpSessionStateBase session = null;
		readonly HttpRequestBase request = null;
		readonly String code = null;
		Uri redirectUri;

		public ServicesHelper(HttpContextBase httpContext)
		{
			this.httpContext = httpContext;
			session = httpContext.Session;
			request = httpContext.Request;
			code = request.QueryString["code"];
		}

		public async Task<OutlookServicesClient> LoadOutlookServicesClient()
		{
			if (String.IsNullOrEmpty(code)) {
				CreateDiscoveryAuthorizationUri();
				return null;
			}

			var discoveryClient = LoadDiscoveryClient();
			var contactsDiscoveryResult =
				await LoadContactsDiscoveryResultAsync(discoveryClient);
			return LoadOutlookServicesClient(contactsDiscoveryResult);
		}

		private void CreateDiscoveryAuthorizationUri()
		{
			redirectUri =
				SettingsHelper.AuthenticationContext.GetAuthorizationRequestURL(
					discoResource, // Where the discovery service is
					SettingsHelper.ClientCredential.ClientId, // This identifies the application
					new Uri(request.Url.AbsoluteUri.Split('?')[0]), // Redirect URL
					UserIdentifier.AnyUser,
					string.Empty
				);
		}

		public ActionResult RedirectToAuthentication()
		{
			return new RedirectResult(redirectUri.ToString());
		}

		public DiscoveryClient LoadDiscoveryClient()
		{
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

		public async Task<CapabilityDiscoveryResult> LoadContactsDiscoveryResultAsync(DiscoveryClient discoveryClient)
		{
			var contactsDiscoveryResult = session["ContactsDiscoveryResult"] as CapabilityDiscoveryResult;

			if (contactsDiscoveryResult == null) {
				contactsDiscoveryResult = await discoveryClient.DiscoverCapabilityAsync("Contacts");
				session["ContactsDiscoveryResult"] = contactsDiscoveryResult;
			}

			return contactsDiscoveryResult;
		}

		public OutlookServicesClient LoadOutlookServicesClient(CapabilityDiscoveryResult contactsDiscoveryResult)
		{
			var outlookServicesClient = session["OutlookServicesClient"] as OutlookServicesClient;
			if (outlookServicesClient == null) {
				if (session["CreatedOutlookClientCode"] == null) {
					CreateOutlookAuthorizationUri(contactsDiscoveryResult);
					session["CreatedOutlookClientCode"] = true;
				} else {
					outlookServicesClient = new OutlookServicesClient(contactsDiscoveryResult.ServiceEndpointUri, async () => {
						var authResult = await authenticationContext.AcquireTokenByAuthorizationCodeAsync(
							code,
							new Uri(request.Url.AbsoluteUri.Split('?')[0]),
							clientCredential);
						return authResult.AccessToken;
					});
					session["OutlookServicesClient"] = outlookServicesClient;
				}
			}

			return outlookServicesClient;
		}

		private void CreateOutlookAuthorizationUri(CapabilityDiscoveryResult contactsDiscoveryResult)
		{
			redirectUri = authenticationContext.GetAuthorizationRequestURL(
					contactsDiscoveryResult.ServiceResourceId,
					clientCredential.ClientId,
					new Uri(request.Url.AbsoluteUri.Split('?')[0]),
					UserIdentifier.AnyUser,
					string.Empty);
		}
	}
}