using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelFileMerger
{
    public partial class Form1 : Form
    {
        string dirpath = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Excel Workbook (*.xlsx) | *.xlsx";
            DialogResult res = fd.ShowDialog();
            if (res != DialogResult.OK)
            {
                return;
            }
            textBox1.Text = fd.FileName;
            string file = fd.FileName;

            string[] fpath = file.Split('\\');
            string outfileName = fpath[fpath.Length - 1].Substring(0, fpath[fpath.Length - 1].Length - 5) + "_OUTPUT.csv";
            string savepath = "";
            for (int i = 0; i < fpath.Length - 1; i++)
            {
                savepath += fpath[i] + "\\";
            }
            textBox2.Text = savepath + outfileName;

        }

        private void button2_Click(object sender, EventArgs e)
        { 

            Thread t = new Thread(() =>
            {
                string file = textBox1.Text;
                OpenFileDialog fd = new OpenFileDialog();
                fd.FileName = file;
                if (!fd.CheckFileExists)
                {
                    MessageBox.Show("Invalid file");
                    return;
                }
                string[] fpath = file.Split('\\');
                Console.WriteLine("File Name =" + fpath[fpath.Length - 1].Substring(0, fpath[fpath.Length - 1].Length - 5));
                string outfileName = fpath[fpath.Length - 1].Substring(0, fpath[fpath.Length - 1].Length - 5) + "_OUTPUT.csv";
                string savepath = "";
                for (int i = 0; i < fpath.Length - 1; i++)
                {
                    savepath += fpath[i] + "\\";
                }


                Console.WriteLine("Reading file");
                updateStatusText("Opening workbook...");
                string constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + file + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"";
                OleDbConnection con = new OleDbConnection(constring);
                try
                {
                    con.Open();
                }
                catch (Exception)
                {
                 
                    MessageBox.Show("Error opening " + file + ". Is the file already open?");
                    return;
                }
                DataTable dtSheet = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                List<string> sheelist = new List<string>();
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))
                    {
                        sheelist.Add(drSheet["TABLE_NAME"].ToString());
                        Console.WriteLine(drSheet["TABLE_NAME"].ToString());
                    }
                }
                OleDbCommand cmd;
                string command = "select * from [" + sheelist.ElementAt(0) + "]";
                updateStatusText("Selecting data....");
                for (int i = 1; i < dtSheet.Rows.Count; i++)
                {


                    command += " union select * from [" + sheelist.ElementAt(i) + "]";

                
                }
                Console.WriteLine(command);
                Console.WriteLine("Creating data set");
                updateStatusText("Creating new dataset...");
                cmd = new OleDbCommand(command, con);
                OleDbDataAdapter adpt = new OleDbDataAdapter(cmd);
                DataTable dt = new DataTable();
                adpt.Fill(dt);
                Console.WriteLine("creating workbook");
                updateStatusText("Writing output file...");
                /* dataTable to csv --from stack*/
                StringBuilder sb = new StringBuilder();

                IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                                  Select(column => column.ColumnName);
                sb.AppendLine(string.Join(",", columnNames));

                foreach (DataRow row in dt.Rows)
                {
                    IEnumerable<string> fields = row.ItemArray.Select(field => ("\"" + field.ToString() + "\""));
                    sb.AppendLine(string.Join(",", fields));
                }

                File.WriteAllText(savepath + outfileName, sb.ToString());

                updateStatusText("Done");
                dt.Clear();
                con.Close();
            });
            t.Start();
        }
        void updateStatusText(string text)
        {
            try
            {
                this.Invoke((MethodInvoker)(() =>
                {
                    this.statusText1.Text = text;
                }));
            }
            catch (Exception)
            {

            }
        }
    }
}
