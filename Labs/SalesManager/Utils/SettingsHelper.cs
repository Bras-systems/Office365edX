using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace SalesManager.Utils
{
	public static class SettingsHelper
	{
		public static readonly string SPSite = ConfigurationManager.AppSettings["spSite"];
		public static readonly string DiscoResource = ConfigurationManager.AppSettings["discoResource"];
		public static readonly string DiscoveryEndpoint = ConfigurationManager.AppSettings["discoEndpoint"];		// Default

		public static AuthenticationContext AuthenticationContext
		{
			get
			{
				return new AuthenticationContext(
						   ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
						   true);
			}
		}

		public static ClientCredential ClientCredential
		{
			get
			{
				return new ClientCredential(
							ConfigurationManager.AppSettings["ida:ClientID"],
							ConfigurationManager.AppSettings["ida:Password"]);
			}
		}
	}
}