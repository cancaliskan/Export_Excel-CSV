using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace exportExcel_CSV.Models
{
    public class Client
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public DateTime DOB { get; set; }
        public string Email { get; set; }

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