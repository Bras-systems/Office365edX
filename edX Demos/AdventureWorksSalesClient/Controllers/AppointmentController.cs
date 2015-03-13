using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using AdventureWorksSalesClient.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using SalesClient.Utils;

namespace AdventureWorksSalesClient.Controllers
{
    public class AppointmentController : Controller
    {
        // GET: Appointment
        public ActionResult Create(string code, string email)
        {
            ViewAppointment appointment = new ViewAppointment();
			appointment.Email = email;
			ViewBag.Code = code;
			return View(appointment);
        }

		[HttpPost()]
		[ValidateAntiForgeryToken()]
		public async Task<ActionResult> Create(string code, ViewAppointment appointment)
		{
			#region Get OutlookServicesClient

			AuthenticationContext authContext = new AuthenticationContext(
												 ConfigurationManager.AppSettings["ida:AuthorizationUri"] + "/common",
												 true);

			ClientCredential creds = new ClientCredential(
				ConfigurationManager.AppSettings["ida:ClientID"],
				ConfigurationManager.AppSettings["ida:Password"]);

			//Get the discovery information that was saved earlier
			CapabilityDiscoveryResult cdr = Helpers.GetFromCache("ContactsDiscoveryResult") as CapabilityDiscoveryResult;

			// Check the cache for the client
			// This is the client that will give us access to all of the "Outlook" information
			OutlookServicesClient outlookClient = Helpers.GetFromCache("OutlookClient") as OutlookServicesClient;

			//Get an authorization code if needed
			if (outlookClient == null && cdr != null && code == null) {
				Uri redirectUri = authContext.GetAuthorizationRequestURL(
					cdr.ServiceResourceId,
					creds.ClientId,
					new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
					UserIdentifier.AnyUser,
					string.Empty);

				return Redirect(redirectUri.ToString());
			}

			//Create the OutlookServicesClient
			if (outlookClient == null && cdr != null && code != null) {
				outlookClient = new OutlookServicesClient(cdr.ServiceEndpointUri, async () => {
					var authResult = await authContext.AcquireTokenByAuthorizationCodeAsync(
						code,
						new Uri(Request.Url.AbsoluteUri.Split('?')[0]),
						creds);
					return authResult.AccessToken;
				});

				Helpers.SaveInCache("OutlookClient", outlookClient);
			}

			#endregion
			
			Event outlookAppointment = new Event();
			outlookAppointment.Body = new ItemBody() { ContentType = BodyType.Text, Content = appointment.Body };
			outlookAppointment.Attendees.Add(
				new Attendee() {
					Type = AttendeeType.Required,
					EmailAddress = new EmailAddress() {
						Address = appointment.Email,
						Name = appointment.Email
					}
				}
			);
			outlookAppointment.Start = appointment.StartTime;
			outlookAppointment.End = appointment.EndTime;
			outlookAppointment.Subject = appointment.Subject;

			await outlookClient.Me.Calendar.Events.AddEventAsync(outlookAppointment);

			TempData["AppointmentEmail"] = appointment.Email;
			return RedirectToAction("Index", "Contact", new { code = code });
		}
    }
}