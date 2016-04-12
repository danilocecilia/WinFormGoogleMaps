using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GMap.NET;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms.ToolTips;
using System.IO;
using Excel;
using System.Diagnostics;
using System.Globalization;
using System.Xml.Serialization;
using System.Net.Mail;
using System.Net;


namespace GoogleMaps
{
    public partial class Form1 : Form
    {
        #region Attributes
        // marker
        GMapMarker currentMarker;
        readonly GMapOverlay top = new GMapOverlay();
        internal readonly GMapOverlay objects = new GMapOverlay("objects");
        internal readonly GMapOverlay routes = new GMapOverlay("routes");
        List<Location> lstLocation;
        int idxRowSelected = -1;
        Location location;
        bool isMouseDown = false;
        PointLatLng start;
        PointLatLng end;

        //static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
        //static int alarmCounter = 1;
        //static bool exitFlag = false;
        #endregion

        #region Constructor
        public Form1()
        {
            InitializeComponent();

            MainMap.MapProvider = GMapProviders.GoogleMap;
            MainMap.Position = new PointLatLng(-23.6342541, -46.5698397);
            MainMap.MinZoom = 0;
            MainMap.MaxZoom = 24;
            MainMap.Zoom = 9;

            MainMap.Overlays.Add(objects);
            MainMap.Overlays.Add(top);

            // set current marker
            currentMarker = new GMarkerGoogle(MainMap.Position, GMarkerGoogleType.arrow);

            top.Markers.Add(currentMarker);

            if (objects.Markers.Count > 0)
            {
                MainMap.ZoomAndCenterMarkers(null);
            }

            RoutingProvider rp = MainMap.MapProvider as RoutingProvider;
            if (rp == null)
            {
                rp = GMapProviders.OpenStreetMap; // use OpenStreetMap if provider does not implement routing
            }

            LoadGMaps(null, null);

            LoadExcel();

            ///* Adds the event and the event handler for the method that will process the timer event to the timer. */
            //myTimer.Tick += new EventHandler(TimerEventProcessor);

            //// Sets the timer interval to 5 seconds.
            ////myTimer.Interval = 86400000; // 1 day
            //myTimer.Interval = 30000;
            //myTimer.Start();

            //// Runs the timer, and raises the event.
            //while (exitFlag == false)
            //{
            //    // Processes all the events in the queue.
            //    Application.DoEvents();
            //}
        }
        #endregion

        #region Methods
        //// This is the method to run when the timer is raised.
        //private static void TimerEventProcessor(Object myObject, EventArgs myEventArgs)
        //{
        //    myTimer.Stop();

        //    try
        //    {
        //        SmtpClient client = new SmtpClient("smtp.live.com");
        //        client.Port = 587;
        //        client.EnableSsl = true;
        //        client.Timeout = 100000;
        //        client.DeliveryMethod = SmtpDeliveryMethod.Network;
        //        client.UseDefaultCredentials = false;
        //        client.Credentials = new NetworkCredential("danilo_cyber@hotmail.com", "myPass");
        //        MailMessage msg = new MailMessage();
        //        msg.To.Add("dcecilia@msxi.com");
        //        msg.From = new MailAddress("danilo_cyber@hotmail.com");
        //        msg.Subject = "Arquivo Location Lojas";
        //        msg.Body = "Arquivo Location Lojas";
        //        Attachment data1 = new Attachment(Path.GetDirectoryName(Application.ExecutablePath) + @"\SavedLocationData.txt");
        //        Attachment data2 = new Attachment(Path.GetDirectoryName(Application.ExecutablePath) + @"\Backup_SavedLocationData.txt");
        //        msg.Attachments.Add(data1);
        //        msg.Attachments.Add(data2);
        //        client.Send(msg);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }

        //    myTimer.Enabled = true;
        //}
        public static string SerializeToXml<T>(T value)
        {
            StringWriter writer = new StringWriter(CultureInfo.InvariantCulture);
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            serializer.Serialize(writer, value);
            return writer.ToString();
        }
        public object Deserialize(string input, System.Type toType)
        {
            System.Xml.Serialization.XmlSerializer ser = new System.Xml.Serialization.XmlSerializer(toType);

            using (StringReader sr = new StringReader(input))
                return ser.Deserialize(sr);
        }
        private void LoadGMaps(PointLatLng? start, string end)
        {
            GDirections route;
            if (start != null && end != null)
            {
                var statusCode = GMap.NET.MapProviders.GoogleMapProvider.Instance.GetDirections(out route, start.Value, end, true);

                if (statusCode == DirectionsStatusCode.NOT_FOUND)
                    MessageBox.Show("Não Localizado, verifique o endereço.");

                if (route != null)
                {
                    // add route
                    GMapRoute r = new GMapRoute(route.Route, "My Route");
                    r.IsHitTestVisible = false;
                    routes.Routes.Add(r);

                    // add route start/end marks
                    GMapMarker m1 = new GMarkerGoogle(start.Value, GMarkerGoogleType.green_big_go);
                    m1.ToolTipText = "Start: " + route.StartAddress;
                    m1.ToolTipMode = MarkerTooltipMode.Always;

                    GMapMarker m2 = new GMarkerGoogle(route.Route[route.Route.Count() - 1], GMarkerGoogleType.red_big_stop);
                    m2.ToolTipText = "End: " + end.ToString();
                    m2.ToolTipMode = MarkerTooltipMode.Always;

                    objects.Markers.Add(m1);
                    objects.Markers.Add(m2);

                    MainMap.ZoomAndCenterRoute(r);
                }
            }
            else if (end != null)
            {
                GeoCoderStatusCode status = MainMap.SetPositionByKeywords(end);
                if (status != GeoCoderStatusCode.G_GEO_SUCCESS)
                {
                    MessageBox.Show("Opsss... Endereço não localizado: '" + end + "', Razão: " + status.ToString(), "GMap.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }
        private void LoadExcel()
        {
            FileStream stream = File.Open(Path.GetDirectoryName(Application.ExecutablePath) + @"\Lojas_Latitude_Longitude.xlsx", FileMode.Open, FileAccess.Read);

            //1. Reading from a binary Excel file ('97-2003 format; *.xls)
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            //...
            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();
            //...
            //4. DataSet - Create column names from first row
            excelReader.IsFirstRowAsColumnNames = true;

            var data = result.Tables[0].ToGenericList<Location>(GoogleMaps.Location.Converter);

            //Remove primeira linha
            data.RemoveAt(0);

            dataGridView1.DataSource = data;

            //dataGridView1.DataSource = result.Tables[0];

            //5. Data Reader methods
            //while (excelReader.Read())
            //{
            //    excelReader.GetInt32(0);
            //}

            //6. Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
        private string ReadTextFile()
        {
            try
            {   // Open the text file using a stream reader.
                using (System.IO.FileStream stream = new System.IO.FileStream(Path.GetDirectoryName(Application.ExecutablePath) + @"\SavedLocationData.txt", System.IO.FileMode.Open, FileAccess.ReadWrite))
                using (System.IO.StreamReader sr = new System.IO.StreamReader(stream))
                {
                    var content = sr.ReadToEnd();
                    stream.Close();

                    if (!string.IsNullOrEmpty(content)) return content;
                    else return string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("The file could not be read: " + ex.Message);
            }
            return string.Empty;
        }
        private void WriteFileContent()
        {
            try
            {
                File.Delete(Path.GetDirectoryName(Application.ExecutablePath) + @"\Backup_SavedLocationData.txt");
                File.Copy(Path.GetDirectoryName(Application.ExecutablePath) + @"\SavedLocationData.txt", Path.GetDirectoryName(Application.ExecutablePath) + @"\Backup_SavedLocationData.txt");

                using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path.GetDirectoryName(Application.ExecutablePath) + @"\SavedLocationData.txt", false))
                {
                    file.WriteLine(SerializeToXml(lstLocation));
                    file.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("The file could not be written: " + ex.Message);
            }
        }
        public static List<T> ConvertDS<T>(DataSet ds, Converter<DataRow, T> converter)
        {
            return
                (from row in ds.Tables[0].AsEnumerable()
                 select converter(row)).ToList();
        }
        private void ChangeRowColors()
        {
            var contentFile = ReadTextFile();

            if (!string.IsNullOrEmpty(contentFile))
            {
                try
                {
                    // Read the stream to a string, and write the string to the console.
                    lstLocation = new List<GoogleMaps.Location>();
                    lstLocation.AddRange(Deserialize(contentFile, typeof(List<Location>)) as List<Location>);

                    var data = dataGridView1.Rows.Cast<DataGridViewRow>().ToList().Where(w => lstLocation.Any(a => a.ID == w.Cells[2].Value.ToString())).ToList();

                    data.ForEach(w =>
                    {
                        w.DefaultCellStyle.BackColor = Color.Green;
                        w.DefaultCellStyle.ForeColor = Color.White;
                    });
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
        }
        private void UpdateLocation(DataGridViewCellEventArgs e, DataGridView senderGrid)
        {
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && senderGrid.Columns[e.ColumnIndex].Name == "Update")
            {
                if (string.IsNullOrEmpty(txtLatitude.Text) && string.IsNullOrEmpty(txtLongitude.Text)) MessageBox.Show("Lat/Lng devem ser preenchidas.", "ATENÇÂO!", MessageBoxButtons.OK);

                //Get saved data from File
                try
                {
                    var contentFile = ReadTextFile();

                    if (!string.IsNullOrEmpty(contentFile))
                    {
                        lstLocation = new List<GoogleMaps.Location>();
                        lstLocation.AddRange(Deserialize(contentFile, typeof(List<Location>)) as List<Location>);

                        PrepareToSave(e);
                    }
                    else
                    {
                        PrepareToSave(e);
                    }

                    LoadExcel();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("The file could not be read: " + ex.Message);
                }
            }
        }
        private void GetRoute(DataGridViewCellEventArgs e, DataGridView senderGrid)
        {
            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && senderGrid.Columns[e.ColumnIndex].Name == "Map")
            {
                PointLatLng? LatLng = null;

                if (dataGridView1.Rows[e.RowIndex].Cells[4].Value != null)
                {
                    LatLng = new PointLatLng
                    {
                        Lat = double.Parse(dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString().Replace(",", ""), System.Globalization.CultureInfo.InvariantCulture),
                        Lng = double.Parse(dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString().Replace(",", ""), System.Globalization.CultureInfo.InvariantCulture)
                    };
                }

                LoadGMaps(LatLng, dataGridView1.Rows[e.RowIndex].Cells["Endereco"].Value.ToString());

                var id = dataGridView1.Rows[e.RowIndex].Cells[2].Value;
                var contentFile = ReadTextFile();

                if (!string.IsNullOrEmpty(contentFile))
                {
                    try
                    {
                        // Read the stream to a string, and write the string to the console.
                        lstLocation = new List<GoogleMaps.Location>();
                        lstLocation.AddRange(Deserialize(contentFile, typeof(List<Location>)) as List<Location>);

                        var foundItem = lstLocation.Find(w => w.ID == id.ToString());

                        if (foundItem != null)
                        {
                            PointLatLng p = new PointLatLng() { Lat = double.Parse(foundItem.Latitude, CultureInfo.InvariantCulture), Lng = double.Parse(foundItem.Longitude, CultureInfo.InvariantCulture) };
                            GMapMarker myCity = new GMarkerGoogle(p, GMarkerGoogleType.green_small);
                            myCity.ToolTipMode = MarkerTooltipMode.Always;
                            myCity.ToolTipText = foundItem.Loja;
                            objects.Markers.Add(myCity);
                        }
                        else
                        {
                            if (idxRowSelected == -1)
                            {
                                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Black;
                                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;

                                idxRowSelected = e.RowIndex;
                            }
                            else
                            {
                                //Set default color
                                dataGridView1.Rows[idxRowSelected].DefaultCellStyle.BackColor = Color.White;
                                dataGridView1.Rows[idxRowSelected].DefaultCellStyle.ForeColor = Color.Black;

                                //Set selected color
                                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Gray;
                                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.White;

                                idxRowSelected = e.RowIndex;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }



            }
        }
        private void PrepareToSave(DataGridViewCellEventArgs e)
        {
            //Fill the Location object
            location = new GoogleMaps.Location()
            {
                ID = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString(),
                Cidade = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString(),
                Latitude = txtLatitude.Text,
                Loja = dataGridView1.Rows[e.RowIndex].Cells[3].Value != null ? dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString() : string.Empty,
                Longitude = txtLongitude.Text,
                UF = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString()
            };

            //Serialize it
            if (lstLocation == null)
            {
                lstLocation = new List<GoogleMaps.Location>();
                lstLocation.Add(location);
            }
            else
            {
                //Check this item already exists on the File
                var foundItem = lstLocation.Find(w => w.ID == location.ID);

                if (foundItem != null)
                {
                    //If exists, I delete from the actual list and add it with the new item.
                    lstLocation.Remove(foundItem);
                }

                lstLocation.Add(location);
            }

            WriteFileContent();
        }
        #endregion

        #region Events
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            objects.Markers.Clear();

            var senderGrid = (DataGridView)sender;

            //Get route
            GetRoute(e, senderGrid);

            //Update the data
            UpdateLocation(e, senderGrid);
        }

        private void MainMap_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = false;
            }
        }

        private void MainMap_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = true;

                if (currentMarker.IsVisible)
                {
                    currentMarker.Position = MainMap.FromLocalToLatLng(e.X, e.Y);

                    var px = MainMap.MapProvider.Projection.FromLatLngToPixel(currentMarker.Position.Lat, currentMarker.Position.Lng, (int)MainMap.Zoom);
                    var tile = MainMap.MapProvider.Projection.FromPixelToTileXY(px);

                    txtLatitude.Text = currentMarker.Position.Lat.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    txtLongitude.Text = currentMarker.Position.Lng.ToString(System.Globalization.CultureInfo.InvariantCulture);

                    Debug.WriteLine("MouseDown: geo: " + currentMarker.Position + " | px: " + px + " | tile: " + tile);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MainMap.ReloadMap();
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            ChangeRowColors();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (((CheckBox)(sender)).Checked)
                MainMap.MapProvider = GMapProviders.GoogleSatelliteMap;
            else
                MainMap.MapProvider = GMapProviders.GoogleMap;
        }
        #endregion
    }

    #region Classes
    public class Location
    {
        public string ID { get; set; }
        public string Loja { get; set; }
        public string Latitude { get; set; }
        public string Longitude { get; set; }
        public string UF { get; set; }
        public string Cidade { get; set; }
        public string Endereco { get; set; }

        [System.ComponentModel.Browsable(false)]
        public string SerializedObj { get; set; }

        public static Location Converter(DataRow row)
        {
            return new Location
            {
                ID = row[0].ToString(),
                Loja = row[1] as string,
                Latitude = row[2] as string,
                Longitude = row[3] as string,
                Endereco = row[4] as string,
                UF = row[5] as string,
                Cidade = row[6] as string
            };
        }
    }

    public static class DataTableExtensions
    {
        public static List<T> ToGenericList<T>(this DataTable datatable, Func<DataRow, T> converter)
        {
            return (from row in datatable.AsEnumerable()
                    select converter(row)).ToList();
        }
    }
    #endregion
}

