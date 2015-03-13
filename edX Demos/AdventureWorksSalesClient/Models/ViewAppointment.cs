using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdventureWorksSalesClient.Models
{
	public class ViewAppointment
	{
		[DataType(DataType.EmailAddress)]
		public string Email { get; set; }

		[DataType(DataType.DateTime)]
		public DateTime StartTime { get; set; }

		[DataType(DataType.DateTime)]
		public DateTime EndTime { get; set; }

		public string Subject { get; set; }

		[DataType(DataType.MultilineText)]
		public string Body { get; set; }
	}
}
