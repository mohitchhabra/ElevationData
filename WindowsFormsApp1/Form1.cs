using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel;
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

            //foreach (var worksheet in Workbook.Worksheets(@"D:\input\a1.csv"))
            //{
            //    foreach (var row in worksheet.Rows)
            //    {
            //        foreach (var cell in row.Cells)
            //        {
            //            cell.Amount = 4;
            //        }
            //    }
            //}
            string fname = @"D:\input\a813.csv";
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            Microsoft.Office.Interop.Excel.Application xlnewApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlnewApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "X";
            xlWorkSheet.Cells[1, 2] = "Y";
            xlWorkSheet.Cells[1, 3] = "name";
            xlWorkSheet.Cells[1, 4] = "elevation";
            int rowCount = xlRange.Rows.Count;
            
           
           // i=2 bcz first row is header and rowcount-1 bcz last row not req to evaluate
            for (int i = 2; i <= rowCount-1; i++)
            {
               // Thread.Sleep(2000);

                var longi =  xlRange.Cells[i, 1].Value2;
                  var lati = xlRange.Cells[i, 2].Value2;
                var name = xlRange.Cells[i, 3].Value2;
                var request = (HttpWebRequest)WebRequest.Create(string.Format("https://maps.googleapis.com/maps/api/elevation/json?locations={0},{1}&key={2}", lati, longi, "AIzaSyAqfARRG69qmS-5hykBwcyPPEMLon2S1Bo"));
                var response = (HttpWebResponse)request.GetResponse();
                var sr = new StreamReader(response.GetResponseStream() ?? new MemoryStream()).ReadToEnd();
                var json = JObject.Parse(sr);
                ElevationResponse elevationResponse = json.ToObject<ElevationResponse>();
                var elevation = elevationResponse.Results.First().Elevation;
                var latitude = elevationResponse.Results.First().Location.Latitude;
                var longitude = elevationResponse.Results.First().Location.Longitude;

                //Create Excel 
                //var f2Name = @"D:\output\a1.xls";
                //DataTable table = new DataTable();
                //table.Columns.Add("X", typeof(double));
                //table.Columns.Add("Y", typeof(double));
                //table.Columns.Add("name", typeof(string));
                //table.Columns.Add("elevation", typeof(double));
                //table.Rows.Add(i, latitude, longitude, name, elevation);


                //Write to excel 
               
                
                xlWorkSheet.Cells[i, 1] = latitude;
                xlWorkSheet.Cells[i, 2] = longitude;
                xlWorkSheet.Cells[i, 3] = name;
                xlWorkSheet.Cells[i, 4] =  elevation;

                //Here saving the file in xlsx
               

            }
            xlWorkBook.SaveAs("d:\\output\\a813.xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


            xlWorkBook.Close(true, misValue, misValue);
            xlnewApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlnewApp);





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
