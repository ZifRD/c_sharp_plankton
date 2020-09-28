using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;

namespace Plankton_Base
{
    public partial class Form1 : Form
    {
        private OleDbConnection connection;

        public OleDbConnection Connection { get => connection; set => connection = value; }     

        private int recordsNumber = 0;
        public int RecordsNumber { get => recordsNumber; set => recordsNumber = value; }

        public Form1()
        {
            InitializeComponent();

        }       

        private void Form1_Load(object sender, EventArgs e)
        {
            string path = System.IO.Path.Combine(Application.StartupPath, "data/Polarnightbase.accdb");
            connection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Persist Security Info=False;");
            connection.Open();
            OleDbCommand query = new OleDbCommand("SELECT `Taxon` FROM `Taxon_All`", connection);
            OleDbDataReader reader = query.ExecuteReader();

            listBox1.Items.Clear();
            listBox1.Items.Add("Все");
           
            while (reader.Read())
            {
                listBox1.Items.Add(reader["Taxon"].ToString());
            }

            query.Dispose();
            reader.Close();

            query = new OleDbCommand("SELECT DISTINCT `Tdate` FROM `All_Data_Fix` ORDER BY `Tdate`", connection);
            reader = query.ExecuteReader();

            listBox2.Items.Clear();
            listBox2.Items.Add("Все");
            while (reader.Read())
            {
                listBox2.Items.Add(reader["Tdate"].ToString().Substring(0,10));
            }

            query.Dispose();
            reader.Close();

            query = new OleDbCommand("SELECT DISTINCT `Region` FROM `All_Data_Fix` ORDER BY `Region`", connection);
            reader = query.ExecuteReader();

            listBox3.Items.Clear();
            listBox3.Items.Add("Все");
            while (reader.Read())
            {
                listBox3.Items.Add(reader["Region"].ToString());
            }

            query.Dispose();
            reader.Close();

            query = new OleDbCommand("SELECT Taxon, Tdate, round(Lat,4) AS Lat, round(Lon,4) AS Lon, Num_cells, Depth_sample, Region FROM `All_Data_Fix` ", connection);
            reader = query.ExecuteReader();
            DataTable data = new DataTable();
            data.Load(reader);
            dataGridView1.DataSource = data;
            recordsNumber = dataGridView1.Rows.Count;
            query.Dispose();
            reader.Close();


        }
     
        private void fillByToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.all_Data_FixTableAdapter.FillBy(this.polarnightbaseDataSet2.All_Data_Fix);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void fillBy1ToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.all_Data_FixTableAdapter.FillBy1(this.polarnightbaseDataSet2.All_Data_Fix);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }
        
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           makeQuery();            
        }
        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            makeQuery();        
        }
        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            makeQuery();        
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            connection.Close();
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;

            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);

            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
          

        private void makeQuery()
        {
            if (listBox1.SelectedIndex == -1) listBox1.SelectedIndex = 0;
            if (listBox2.SelectedIndex == -1) listBox2.SelectedIndex = 0;
            if (listBox3.SelectedIndex == -1) listBox3.SelectedIndex = 0;

            int taxonIndex = listBox1.SelectedIndex;
            int dateIndex = listBox2.SelectedIndex;
            int regionIndex = listBox3.SelectedIndex;
            string taxon, date, region;
            if (taxonIndex != 0)
                taxon = "='" + listBox1.Items[taxonIndex].ToString() + "'";
            else
                taxon = "<> ''";

            DateTime dateTime = Convert.ToDateTime("01.01.3000");
            if (dateIndex != 0)
            {
                date = "= @datetime ";
                dateTime = Convert.ToDateTime(listBox2.Items[listBox2.SelectedIndex].ToString());
            }
            else
                date = "< @datetime ";
            if (regionIndex != 0)
                region = "='" + listBox3.Items[regionIndex].ToString() + "'";
            else
                region = "<> ''";
            string sql = "SELECT Taxon, Tdate, round(Lat,4) AS Lat, round(Lon,4) AS Lon, Num_cells, Depth_sample, Region FROM `All_Data_Fix` WHERE `Taxon`"
                 + taxon + " AND Tdate" + date + " AND Region " + region + "  ORDER BY `Tdate`,Taxon";

            OleDbCommand query = new OleDbCommand(sql, connection);
            query.Parameters.AddWithValue("@datetime", dateTime);           

            OleDbDataReader reader = query.ExecuteReader();

            DataTable data = new DataTable();
            data.Load(reader);
            dataGridView1.DataSource = data;
            query.Dispose();
            reader.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            connection.Close();
            Application.Exit();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.google.com");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.yandex.ru");
        }        

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.Show();
            Hide();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Form3 f3 = new Form3(this);
            f3.ShowDialog(this);
        }
    }
    }
