using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SalesClient.Models
{
	public class ViewContact
	{
		public String Id { get; set; }
		public String GivenName { get; set; }
		public String Surname { get; set; }
		public String CompanyName { get; set; }
		public String EmailAddress { get; set; }
		public String BusinessPhone { get; set; }
		public String HomePhone { get; set; }
		public String DisplayName { get; set; }
	}
}
