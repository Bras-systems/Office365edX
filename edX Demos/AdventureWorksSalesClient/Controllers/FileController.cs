using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using AdventureWorksSalesClient.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;
using FileServices = Microsoft.Office365.SharePoint.FileServices;
using SalesClient.Utils;

namespace AdventureWorksSalesClient.Controllers
{
    public class FileController : Controller
    {
        // GET: File
        public async Task<ActionResult> Index(string code)
        {
			#region Authentication
			
			AuthenticationContext authContext = new AuthenticationContext(
						   ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
						   true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("FilesDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
			SharePointClient sharepointClient = Helpers.GetFromCache("SharePointClient") as SharePointClient;

			//Get an authorization code, if needed
			if (sharepointClient == null && cdr != null && code == null) {
				Uri redirectUri = authContext.GetAuthorizationRequestURL(
					cdr.ServiceResourceId,
					creds.ClientId,
					new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
					UserIdentifier.AnyUser,
					string.Empty);

				return Redirect(redirectUri.ToString());
			}

			//Create the SharePointClient
			if (sharepointClient == null && cdr != null && code != null) {

				sharepointClient = new SharePointClient(cdr.ServiceEndpointUri, async () => {

					var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
						code,
						new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
						creds);

					return authResult.AccessToken;
				});

				Helpers.SaveInCache("SharePointClient", sharepointClient);
			}
			#endregion

			//Get the files
			var filesResults = await sharepointClient.Files.ExecuteAsync();

			var fileList = new List<ViewFile>();

			foreach (var file in filesResults.CurrentPage.Where(f => f.Name != "Shared with Everyone").OrderBy(e => e.Name)) {
				fileList.Add(new ViewFile {
					Id = file.Id,
					Name = file.Name,
					Url = file.WebUrl
				});
			}

			//Show the files
			return View(fileList);
		}

		public ActionResult Create(string code)
		{
			ViewBag.Code = code;
			return View();
		}

		[HttpPost()]
		public async Task<ActionResult> Create(string code, ViewFileCreate file)
		{
			#region Authentication

			AuthenticationContext authContext = new AuthenticationContext(
						   ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
						   true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("FilesDiscoveryResult") as CapabilityDiscoveryResult;

			//Get a client, if this page was already visited
			SharePointClient sharePointClient = Helpers.GetFromCache("SharePointClient") as SharePointClient;

			//Get an authorization code, if needed
			if (sharePointClient == null && cdr != null && code == null) {
				Uri redirectUri = authContext.GetAuthorizationRequestURL(
					cdr.ServiceResourceId,
					creds.ClientId,
					new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
					UserIdentifier.AnyUser,
					string.Empty);

				return Redirect(redirectUri.ToString());
			}

			//Create the SharePointClient
			if (sharePointClient == null && cdr != null && code != null) {

				sharePointClient = new SharePointClient(cdr.ServiceEndpointUri, async () => {

					var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
						code,
						new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
						creds);

					return authResult.AccessToken;
				});

				Helpers.SaveInCache("SharePointClient", sharePointClient);
			}
			#endregion

			MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(file.Content));
			var oneDriveFile = new FileServices.File() { Name = file.FileName + ".txt" };

			// Create the item
			await sharePointClient.Files.AddItemAsync(oneDriveFile);

			// upload the data into the item
			await sharePointClient.Files.GetById(oneDriveFile.Id).ToFile().UploadAsync(stream);

			// Redirect back to the index
			TempData["FileName"] = oneDriveFile.Name;
			return RedirectToAction("Index", new { code = code });
		}
    }
}