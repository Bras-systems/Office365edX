using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using System.Configuration;
using System.Threading.Tasks;
using SalesClient.Utils;
using SalesClient.Models;
using Microsoft.Office365.SharePoint.CoreServices;
using System.Text;
using System.Xml.Linq;
using System.Net.Http;
using System.Net.Http.Headers;

namespace SalesClient.Controllers
{
	public class HomeController : Controller
	{
		private const string spSite = "https://geektrainerdev.sharepoint.com";
        private const string discoResource = "https://api.office.com/discovery/";
        private const string discoEndpoint = "https://api.office.com/discovery/v1.0/me/";
		
		public ActionResult Index(string code)
		{
            AuthenticationContext authContext = new AuthenticationContext(
               ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
               true);

            ClientCredential creds = new ClientCredential(
                ConfigurationManager.AppSettings["ida:ClientID"],
                ConfigurationManager.AppSettings["ida:Password"]);

			return View();
		}

		public ActionResult About()
		{
			ViewBag.Message = "Your application description page.";

			return View();
		}

		public ActionResult Contact()
		{
			ViewBag.Message = "Your contact page.";

			return View();
		}
	}
}