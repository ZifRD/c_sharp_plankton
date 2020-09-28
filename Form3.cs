using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlTypes;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Plankton_Base
{
    public partial class Form3 : Form    {

        private Form1 parent;
        private CultureInfo provider = CultureInfo.InvariantCulture;
        private NumberFormatInfo nfi = new NumberFormatInfo();

        private int stations = 0;
        private string region = "";
        private string ship = "";
        private double lat = 0.0;
        private double lon = 0.0;
        private double depthsea = -1.0;
        private double depthsample = 0.0;
        private DateTime tdate = Convert.ToDateTime("01.01.3000");
        private DateTime time = Convert.ToDateTime("01.01.3000");
        private DateTime datetime = Convert.ToDateTime("01.01.3000");
        private int courseID = 0;
        private string taxon = "";
        private int numcells = 0;
        private int id = 0;

        public Form3(Form1 parent)
        {
            this.parent = parent;            
            NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;
            nfi.NumberDecimalDigits = 4;
            nfi.NumberDecimalSeparator = ".";
            InitializeComponent();
        }

		/// <summary>
		/// Load data from form fields.
		/// </summary>		
        private void button1_Click(object sender, EventArgs e)
        {
            nfi.NumberDecimalSeparator = ".";
            stations = Int32.Parse(stationsTB.Text);
            region = regionTB.Text;
            ship = shipTB.Text;
            lat = Double.Parse(latTB.Text, provider);
            lon = Double.Parse(lonTB.Text, provider);
            depthsea = Double.Parse(depthseaTB.Text, provider);
            depthsample = Double.Parse(depthsampleTB.Text, provider);
            tdate = Convert.ToDateTime(tdateTB.Text);
            string format = "H:mm:ss";
            time = DateTime.ParseExact(timeTB.Text, format, provider);
            string format2 = "dd.MM.yyyy H:mm:ss";
            datetime = DateTime.ParseExact(tdateTB.Text + " " + timeTB.Text, format2, provider);
            courseID = Int32.Parse(courseIDTB.Text);
            taxon = taxonTB.Text;
            numcells = Int32.Parse(numcellsTB.Text);
            id = Int32.Parse(idTB.Text);

            int result = addRecord();
            switch (result)
            {
                case 0: { MessageBox.Show("Record exists! No changes!"); break; }
                case 1: { MessageBox.Show("Record added successfully!"); break; }
                case -1: { MessageBox.Show("Some error obtained! No changes!"); break; }
            }//MessageBox.Show("Done successfully!");
        }
        
		/// <summary>
		/// Try to add record
		/// </summary>		
		/// <returns> 0 - exists; 1 - added;  1 - error</returns>
		private int addRecord()
        {
            //check if record in database
            nfi.NumberDecimalDigits = 4;
            string sql = "SELECT Taxon, Tdate, round(Lat,4) AS Lat, round(Lon,4) AS Lon, Num_cells, Depth_sample, Region FROM `All_Data_Fix` WHERE " +
                "Stations = " + stations.ToString() + " AND Region = '" + region + "' AND Ship = '" + ship + "'" + " AND round(Lat,5) = " + Math.Round(lat, 5).ToString("N",nfi) +
                " AND round(Lon,5) =" + Math.Round(lon, 5).ToString("N",nfi) + " AND round(Depth_sample,2) = " + Math.Round(depthsample,2).ToString("N",nfi) + " AND CourseID = " + courseID.ToString() +
                " AND Tdate = @dt1" + " AND Taxon = '" + taxon + "' AND id = " + id.ToString() + "  ORDER BY `Tdate`";

            OleDbCommand query = new OleDbCommand(sql, parent.Connection);
            query.Parameters.AddWithValue("@dt1", tdate);

            OleDbDataReader reader = query.ExecuteReader();

            System.Data.DataTable data = new System.Data.DataTable();
            data.Load(reader);
            if (data.Rows.Count < 1)
            {
                query.Dispose();
                reader.Close();

                sql = "INSERT INTO All_Data_Fix(Stations, Region, Ship, Lat, Lon, Depth_Sea, Depth_sample,Date_time,Tdate,Ttime,CourseID, Taxon, Num_cells,id) values(" +
                    stations + ",'" + region + "','" + ship + "'," + Math.Round(lat, 5).ToString("N",nfi) + "," + Math.Round(lon, 5).ToString("N",nfi) +
                    "," + depthsea.ToString("N",nfi) + "," + depthsample.ToString("N",nfi) + ",@datetime,@dt1,@dt2," + courseID.ToString() + ",'" + taxon + "'," + numcells.ToString() + "," +
                id.ToString() + ")";

                query = new OleDbCommand(sql, parent.Connection);
                query.Parameters.AddWithValue("@datetime", datetime);
                query.Parameters.AddWithValue("@dt1", tdate.ToShortDateString());
                query.Parameters.AddWithValue("@dt2", time.ToShortTimeString());


                if (query.ExecuteNonQuery() > 0)
                {
                    parent.RecordsNumber++;
                    query.Dispose();
                    reader.Close();
                    //MessageBox.Show("Done successfully!");
                    return 1;
                }
                else
                {
                    query.Dispose();
                    reader.Close();
                    return -1;
                }

            }
            else
            {
                query.Dispose();
                reader.Close();
                //MessageBox.Show("Record exists! No changes!");
                return 0;
            }
        }       

        private async void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Maximum = 100;
            progressBar1.Step = 1;

            var progress = new Progress<int>(v =>
            {
                // This lambda is executed in context of UI thread,
                // so it can safely update form controls
                progressBar1.Value = v;
            });

            // Run operation in another thread
            OpenFileDialog openfileDialog1 = new OpenFileDialog();
            string fileName = "";
            if (openfileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = openfileDialog1.FileName;
            }
            await Task.Run(() => DoWork(progress,fileName));
            progressBar1.Value = 0;
        }

		/// <summary>
		/// Load from single file.
		/// </summary>
		/// <param name="progress"> Progress bar manager</param>
		/// <param name="fileName"> Name </param>		
        public void DoWork(IProgress<int> progress,string fileName)
        {               
               
                OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0");           
                con.Open();
                DataSet myDataSet = new DataSet();
                OleDbDataAdapter myCommand = null;

                int added = 0;
                int exists = 0;
                int errors = 0;
                int total = 0;


                try
                {
                    //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference                    
                    myCommand = new OleDbDataAdapter(" SELECT * FROM [Лист1$]", con);
                    myCommand.Fill(myDataSet);
                    con.Close();

                    int count = myDataSet.Tables[0].Rows.Count;

                    for (int i = 0; i < count;i++)
                    {
                        Object[] cells = myDataSet.Tables[0].Rows[i].ItemArray;

                        
                        {
                            nfi.NumberDecimalSeparator = ",";
                            stations = Int32.Parse(cells[0].ToString());
                            region = cells[1].ToString();
                            ship = cells[2].ToString();
                            lat = Double.Parse(cells[3].ToString());
                            lon = Double.Parse(cells[4].ToString());
                            depthsample = Double.Parse(cells[5].ToString());
                            tdate = Convert.ToDateTime(cells[6].ToString().Substring(0, 10));
                            string format = "H:mm:ss";
                            if(cells[7].ToString() != "") time = DateTime.ParseExact(cells[7].ToString(), format, provider);
                            else time = Convert.ToDateTime("01.01.3000");
                            string format2 = "dd.MM.yyyy H:mm:ss";


                            //string str1 = cells[i + 6].ToString().Substring(0, 10) + " " + time.TimeOfDay.ToString();


                                datetime = DateTime.ParseExact(cells[6].ToString().Substring(0, 10) + " " + time.TimeOfDay.ToString(), format2, provider);

                            taxon = cells[8].ToString();
                            numcells = Int32.Parse(cells[9].ToString());
                            nfi.NumberDecimalSeparator = ".";
                            total++;

                            int result = addRecord();
                            switch (result)
                            {
                                case 0: { exists++; break; }
                                case 1: { added++; break; }
                                case -1: { errors++; break; }
                            }//MessageBox.Show("Done successfully!");
                        }
                        if (progress != null)
                                progress.Report((int)((i+1) * 100.0 / count));

                    }

                    //Thread.Sleep(15000);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    //Thread.Sleep(15000);
                }
                finally
                {
                    con.Close();
                    MessageBox.Show("Done! Total: "+total.ToString()+", added: "+ added.ToString()+", exists: "+exists.ToString()+", errors: "+errors.ToString()+"!");
                   
                    myDataSet.Dispose();
                    myCommand.Dispose();
                }

        }

        private async void button4_Click(object sender, EventArgs e)
        {
            progressBar1.Value = 0;
            progressBar1.Maximum = 100;
            progressBar1.Step = 1;

            var progress = new Progress<int>(v =>
            {
                // This lambda is executed in context of UI thread,
                // so it can safely update form controls
                progressBar1.Value = v;
            });

            // Run operation in another thread
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Custom Description";
            string selectedPath = "";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                selectedPath = fbd.SelectedPath;
            }
            await Task.Run(() => DoWorkSeveral(progress, selectedPath));
            progressBar1.Value = 0;
        }

		/// <summary>
		/// Загрузка нескольких файлов.
		/// </summary>
		/// <param name="progress"> Менеджер для progress bar </param>
		/// <param name="selectedPath"> Путь </param>		
        public void DoWorkSeveral(IProgress<int> progress, string selectedPath)                      

            DirectoryInfo dir = new DirectoryInfo(selectedPath);

            int filesCount = dir.GetFiles().Length;
            int counter0 = 0;
            foreach (FileInfo file in dir.GetFiles())
            {
                string fileName = file.FullName;
                OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 12.0");

                //   Provider = Microsoft.ACE.OLEDB.12.0; Data Source = { 0 }; Extended Properties =\"Excel 12.0;HDR=YES
                con.Open();
                DataSet myDataSet = new DataSet();
                OleDbDataAdapter myCommand = null;

                int added = 0;
                int exists = 0;
                int errors = 0;
                int total = 0;

                try                {
                    //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference

                    myCommand = new OleDbDataAdapter(" SELECT * FROM [Лист1$]", con);
                    myCommand.Fill(myDataSet);
                    con.Close();

                    int count = myDataSet.Tables[0].Rows.Count;

                    for (int i = 0; i < count; i++)
                    {
                        Object[] cells = myDataSet.Tables[0].Rows[i].ItemArray;


                        {
                            nfi.NumberDecimalSeparator = ",";
                            stations = Int32.Parse(cells[0].ToString());
                            region = cells[1].ToString();
                            ship = cells[2].ToString();
                            lat = Double.Parse(cells[3].ToString());
                            lon = Double.Parse(cells[4].ToString());
                            depthsample = Double.Parse(cells[5].ToString());
                            tdate = Convert.ToDateTime(cells[6].ToString().Substring(0, 10));
                            string format = "H:mm:ss";
                            if (cells[7].ToString() != "") time = DateTime.ParseExact(cells[7].ToString(), format, provider);
                            else time = Convert.ToDateTime("01.01.3000");
                            string format2 = "dd.MM.yyyy H:mm:ss";
                            datetime = DateTime.ParseExact(cells[6].ToString().Substring(0, 10) + " " + time.TimeOfDay.ToString(), format2, provider);
                            taxon = cells[8].ToString();
                            numcells = Int32.Parse(cells[9].ToString());
                            nfi.NumberDecimalSeparator = ".";
                            total++;

                            int result = addRecord();
                            switch (result)
                            {
                                case 0: { exists++; break; }
                                case 1: { added++; break; }
                                case -1: { errors++; break; }
                            }//MessageBox.Show("Done successfully!");
                        }                     

                    }

                    //Thread.Sleep(15000);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    //Thread.Sleep(15000);
                }
                finally
                {
                    con.Close();
                    MessageBox.Show((counter0+1).ToString()+ " of "+filesCount.ToString()+" (" + fileName +"): done!\nTotal: " + total.ToString() + ", added: " + added.ToString() + ", exists: " + exists.ToString() + ", errors: " + errors.ToString() + "!");

                    myDataSet.Dispose();
                    myCommand.Dispose();
                }
                counter0++;
                if (progress != null)
                    progress.Report((int)(counter0 * 100.0 / filesCount));
            }

        }
    }
}

