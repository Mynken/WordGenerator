using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using Wordgenerator.Logic.Additional;
using Wordgenerator.Models.DAL.Additional;

namespace Wordgenerator.Controllers
{
    public class DocumentAdditionalIeController : ApiController
    {
        public HttpResponseMessage Post([FromBody]AdditionalIE data)
        {
            try
            {
                var generator = new AdditionalIEGenerator();
                var path = System.Web.Hosting.HostingEnvironment.MapPath("~/App_Data/Additional");
                var filePath = generator.CreateIEDocument(data, path);

                var response = new HttpResponseMessage(HttpStatusCode.OK);
                var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open);
                response.Content = new StreamContent(stream);
                response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");

                return response;

            }
            catch (Exception ex)
            {
                var response = new HttpResponseMessage(HttpStatusCode.BadGateway);
                response.Content = new StringContent(ex.ToString());
                return response;
            }
        }
    }
}