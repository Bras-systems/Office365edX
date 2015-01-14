using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(CalendarClient.Startup))]
namespace CalendarClient
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
