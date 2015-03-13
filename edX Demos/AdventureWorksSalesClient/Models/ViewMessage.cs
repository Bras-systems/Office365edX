using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AdventureWorksSalesClient.Models
{
	public class ViewMessage
	{
		[DataType(DataType.EmailAddress)]
		public string Email { get; set; }

		[DataType(DataType.MultilineText)]
		public string Body { get; set; }
	}
}
