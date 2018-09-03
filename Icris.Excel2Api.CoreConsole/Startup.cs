using System;
using System.Threading.Tasks;
using System.Web.Http;
using Microsoft.Owin;
using Owin;

[assembly: OwinStartup(typeof(Icris.Excel2Api.CoreConsole.Startup))]

namespace Icris.Excel2Api.CoreConsole
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            HttpConfiguration config = new HttpConfiguration();

            config.Routes.MapHttpRoute(
                name: "Swagger",
                routeTemplate: "api/swagger/{*path}",
                defaults: new { controller = "Swagger" }
            );
            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{*path}",
                defaults: new { controller = "Sheet" }
            );


            app.UseWebApi(config);

            app.UseStaticFiles(new Microsoft.Owin.StaticFiles.StaticFileOptions() { FileSystem = new Microsoft.Owin.FileSystems.PhysicalFileSystem("www") });

        }
    }
}
