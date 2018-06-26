using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace Wordgenerator
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }

        protected void Application_BeginRequest()
        {
            //obsługa CORS - Cross site origin - Klient i Api to oddzielne aplikacje
            if (Request.Headers.AllKeys.Contains("Origin"))
            {
                Response.Headers.Add("Access-Control-Allow-Origin", "*");
                Response.Headers.Add("Access-Control-Allow-Headers", "Origin, Content-Type, X-Auth-Token, Access-Control-Allow-Credentials, X-Requested-With, Content-Disposition");
                Response.Headers.Add("Access-Control-Allow-Methods", "GET, POST, PATCH, PUT, DELETE, OPTIONS");
                Response.Headers.Add("Access-Control-Allow-Credentials", "true");
                Response.Headers.Add("Access-Control-Max-Age", "1728000");
                Response.Headers.Add("Access-Control-Expose-Headers", "Content-Disposition");

                if (Request.HttpMethod == "OPTIONS")
                {
                    Response.End();
                }
            }
        }
    }
}
