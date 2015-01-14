using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace CalendarDemo.Models
{
	public class CalendarItem
	{
		[DisplayName("Subject")]
		public String Subject { get; set; }

		[DisplayName("Start Time")]
		[DisplayFormat(DataFormatString="{0:MM/dd/yyyy hh:mm tt}", ApplyFormatInEditMode=true)]
		public DateTimeOffset? Start { get; set; }

		[DisplayName("End Time")]
		[DisplayFormat(DataFormatString = "{0:MM/dd/yyyy hh:mm tt}", ApplyFormatInEditMode = true)]
		public DateTimeOffset? End { get; set; }

		[DisplayName("Location")]
		public String Location { get; set; }

		[DisplayName("Body")]
		public String Body { get; set; }

	}
}