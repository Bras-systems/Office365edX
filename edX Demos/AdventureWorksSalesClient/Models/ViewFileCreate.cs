using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace AdventureWorksSalesClient.Models
{
	public class ViewFileCreate
	{
		public string FileName { get; set; }
		
		[DataType(DataType.MultilineText)]
		public string Content { get; set; }
	}
}