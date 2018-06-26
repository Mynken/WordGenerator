using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using Wordgenerator.Logic;
using Wordgenerator.Models.VM;

namespace Wordgenerator.Controllers
{
    public class DocumentLlcController : ApiController
    {
        public HttpResponseMessage PostLLC([FromBody]DocumentLLCModel data)
        {
            try
            {
                var generator = new WordLLCGenerator();
                var path = System.Web.Hosting.HostingEnvironment.MapPath("~/App_Data");
                var filePath = generator.CreateIEDocument(data.Kontrahent, data.Film, data.DataForDoc, data.Trailers, path);

                if (data.DataForDoc.IsPdf)
                {
                    Document document = new Document();
                    document.LoadFromFile(filePath);
                    Paragraph paragraph = document.Sections[0].AddParagraph();

                    DocPicture picture = paragraph.AppendPicture(Image.FromFile(System.Web.Hosting.HostingEnvironment.MapPath("~/Content/stamp.png")));
                    picture.Height = 80;
                    picture.Width = 220;
                    picture.TextWrappingStyle = TextWrappingStyle.InFrontOfText;
                    picture.TextWrappingType = TextWrappingType.Both;

                    picture.VerticalOrigin = VerticalOrigin.OuterMarginArea;
                    picture.VerticalPosition = -60;
                    string output = path + "\\" + System.IO.Path.GetFileNameWithoutExtension(filePath) + ".pdf";
                    document.Protect(ProtectionType.AllowOnlyReading, "kinomania");
                    document.SaveToFile(output, FileFormat.PDF);

                    var response = new HttpResponseMessage(HttpStatusCode.OK);
                    var stream = new System.IO.FileStream(output, System.IO.FileMode.Open);
                    response.Content = new StreamContent(stream);
                    response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf");

                    return response;
                }
                else
                {
                    var response = new HttpResponseMessage(HttpStatusCode.OK);
                    var stream = new System.IO.FileStream(filePath, System.IO.FileMode.Open);
                    response.Content = new StreamContent(stream);
                    response.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-stream");

                    return response;
                }
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