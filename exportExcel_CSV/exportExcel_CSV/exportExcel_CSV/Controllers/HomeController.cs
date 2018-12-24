using exportExcel_CSV.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace exportExcel_CSV.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            #region list user
            var clients = Client.GenerateDummyClientList();
            ViewBag.ClientList = clients;
            #endregion

            return View();
        }

        public void exportToCSV()
        {
            StringWriter sw = new StringWriter();

            sw.WriteLine("\"First Name\",\"Last Name\",Date of Birth\",Email\"");

            Response.Clear();
            Response.ClearHeaders();
            Response.ClearContent();

            Response.AddHeader("content-disposition", "attachment;filename=ExportedClientList.csv");
            Response.ContentType = "text/csv";

            var clients = Client.GenerateDummyClientList();

            foreach (var client in clients)
            {
                sw.WriteLine(string.Format("\"{0}\",\"{1}\",{2}\",{3}\"",
                    client.FirstName,
                    client.LastName,
                    client.DOB,
                    client.Email));
            }

            Response.Write(sw.ToString());
            Response.End();
        }

        public void exportToExcel()
        {
            var grid = new GridView();

            grid.DataSource = from data in Client.GenerateDummyClientList()
                              select new
                              {
                                  FirstName = data.FirstName,
                                  LastName = data.LastName,
                                  DOB = data.DOB,
                                  Email = data.Email
                              };
            grid.DataBind();

            Response.Clear();
            Response.ClearHeaders();
            Response.ClearContent();

            Response.ContentEncoding = System.Text.Encoding.GetEncoding("windows-1254");
            Response.Charset = "windows-1254";
            Response.AddHeader("content-disposition", "attachment;filename=ExportedClientList.xls");
            Response.ContentType = "application/vnd.ms-excel";

            StringWriter sw = new StringWriter();
            HtmlTextWriter htmlTextWriter = new HtmlTextWriter(sw);

            grid.RenderControl(htmlTextWriter);

            Response.Write("<meta http-equiv='Content-Type' content ='text/html; charset = 'windows-1254'/>" + sw.ToString());

            Response.End();
        }

    }
}