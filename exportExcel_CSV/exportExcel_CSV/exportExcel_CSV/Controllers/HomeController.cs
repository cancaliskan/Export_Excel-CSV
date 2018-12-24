using exportExcel_CSV.Models;
using exportExcel_CSV.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using MoreLinq;

namespace exportExcel_CSV.Controllers
{
    // used morelinq for convert list to datatable
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            #region list user
            var clients = GenerateDummyClientList();
            ViewBag.ClientList = clients;
            #endregion

            return View();
        }

        public ActionResult exportToCSV()
        {
            return new CsvResult(GenerateDummyClientList().ToDataTable());
        }

        public ActionResult exportToExcel()
        {
            return new ExcelResult(GenerateDummyClientList().ToDataTable());
        }

        public static List<Client> GenerateDummyClientList()
        {
            var client = new List<Client>
            {
                new Client{FirstName="Can", LastName="Çalışkan", DOB=DateTime.Parse("07/09/1995"), Email="cancaliskan@windowslive.com"},
                new Client{FirstName="Can", LastName="Çalışkan", DOB=DateTime.Parse("07/09/1995"), Email="cancaliskan@windowslive.com"},
                new Client{FirstName="Can", LastName="Çalışkan", DOB=DateTime.Parse("07/09/1995"), Email="cancaliskan@windowslive.com"}
            };

            return client;
        }
    }
}
