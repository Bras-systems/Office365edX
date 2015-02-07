using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SalesClient.Models
{
    public class ViewDiscovery
    {
        public string Capability { get; set; }
        public string EndpointUri { get; set; }
        public string ResourceId { get; set; }
        public string Version { get; set; }
    }
}