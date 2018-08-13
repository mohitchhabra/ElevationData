using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GoogleMapsApi;
using GoogleMapsApi.Entities.Common;
using GoogleMapsApi.Entities.Elevation.Request;
using GoogleMapsApi.Entities.Elevation.Response;
using Newtonsoft.Json.Linq;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            var request = (HttpWebRequest)WebRequest.Create(string.Format("https://maps.googleapis.com/maps/api/elevation/json?locations={0},{1}", 48.40763,4.14798));
            
            var response = (HttpWebResponse)request.GetResponse();
            var sr = new StreamReader(response.GetResponseStream() ?? new MemoryStream()).ReadToEnd();
            var json = JObject.Parse(sr);
            ElevationResponse elevationResponse = json.ToObject<ElevationResponse>();
            var elevation = elevationResponse.Results.First().Elevation;
            var latitude = elevationResponse.Results.First().Location.Latitude;
            var longitude = elevationResponse.Results.First().Location.Longitude;




            //ElevationRequest elevationRequest = new ElevationRequest()
            //{
            //    Locations = new Location[] { new Location(54, 78) },
            //};
            //try
            //{
            //    var request = new ElevationRequest { Locations = new[] { new Location(40.7141289, -73.9614074) } };
            //     result = GoogleMaps.Elevation.Query(elevationRequest);
            //    var final = result.ToString();

            //}
            //catch (Exception exception)
            //{
            //    Console.WriteLine(exception);
            //    throw;
            //}


        }
    }
}
