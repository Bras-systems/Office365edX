using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SalesClient.Controllers
{
    public class FilesController : Controller
    {
		private const string discoResource = "https://api.office.com/discovery/";
		private const string discoEndpoint = "https://api.office.com/discovery/v1.0/me/";		// Default

        // GET: Files
        public ActionResult Index()
        {
			return null;
        }
    }
}