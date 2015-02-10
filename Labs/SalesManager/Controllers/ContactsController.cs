using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using SalesManager.Models;
using SalesManager.Utils;

namespace SalesManager.Controllers
{
    public class ContactsController : Controller
    {
        // GET: Contacts
        public async Task<ActionResult> Index(String code)
        {
			ServicesHelper servicesHelper = new ServicesHelper(HttpContext);
			var outlookServicesClient = await servicesHelper.LoadOutlookServicesClient();
			if (outlookServicesClient == null) return servicesHelper.RedirectToAuthentication();

			var contactsResults = await outlookServicesClient.Me.Contacts.ExecuteAsync();
			List<ViewContact> contacts = new List<ViewContact>();

			foreach (var contact in contactsResults.CurrentPage.OrderBy(c => c.Surname)) {
				contacts.Add(new ViewContact {
					Id = contact.Id,
					GivenName = contact.GivenName,
					Surname = contact.Surname,
					//DisplayName = contact.Surname + ", " + contact.GivenName,
					//CompanyName = contact.CompanyName,
					//EmailAddress = contact.EmailAddresses.FirstOrDefault().Address,
					//BusinessPhone = contact.BusinessPhones.FirstOrDefault(),
					//HomePhone = contact.HomePhones.FirstOrDefault()
				});
			}

			return View(contacts);
        }
    }
}