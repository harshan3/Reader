using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Reader
{
    public partial class Form1 : Form
    {

        BackgroundWorker bgw = new BackgroundWorker();
        public Form1()
        {
            InitializeComponent();
            label3.Text = "";
         
        }

        DataTableCollection tableCollection;

        private void button1_Click(object sender, EventArgs e)

        {

        }

        private void txtFilename_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {





        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            // this.progressBar1.Increment(1);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {



        }

        //CARD ISSUER DATE - date correction

        private void CARD_ISSUER_DATE()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;


            foreach (DataRow r in dt.Select("[SET 1] = 'CARD ISSUER DATE'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    if (r[x].ToString() != null && !r[x].ToString().Equals(""))
                    {
                        DateTime dts;
                        String str = r[x].ToString();
                        str = Regex.Replace(str, @"[^\d]", "");
                        String stf = Regex.Replace(str.Replace("-", "/"), @"[^\d/]", "");

                     //   Console.WriteLine("String: " + r[x].ToString() + "\n Removed Spaces: " + str);

                        if (DateTime.TryParseExact(stf, "MMddyyyy", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out dts))
                        {
                            String a = dts.ToString("MM/dd/yyyy");
                            r[x] = a;
                            //Console.WriteLine("String: " + r[x].ToString() + "\n Removed Spaces: " + str + "\n Date formatted: " + a + "\n");
                        }

                    }

                }

        }

        //CARD HOLDER DOB - date correction

        private void CARD_HOLDER_DOB()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;


            foreach (DataRow r in dt.Select("[SET 1] = 'CARD HOLDER DOB'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    if (r[x].ToString() != null && !r[x].ToString().Equals(""))
                    {
                        DateTime dts;
                        String str = r[x].ToString();
                        str = Regex.Replace(str, @"[^\d]", "");
                        String stf = Regex.Replace(str.Replace("-", "/"), @"[^\d/]", "");

                        //   Console.WriteLine("String: " + r[x].ToString() + "\n Removed Spaces: " + str);

                        if (DateTime.TryParseExact(stf, "MMddyyyy", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out dts))
                        {
                            String a = dts.ToString("MM/dd/yyyy");
                            r[x] = a;
                            //Console.WriteLine("String: " + r[x].ToString() + "\n Removed Spaces: " + str + "\n Date formatted: " + a + "\n");
                        }

                    }

                }

        }


        //CARD ISSUER DATE - date correction

        private void BENEFICIARY_DOB()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;


            foreach (DataRow r in dt.Select("[SET 1] = 'BENEFICIARY DOB'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    if (r[x].ToString() != null && !r[x].ToString().Equals(""))
                    {
                        DateTime dts;
                        String str = r[x].ToString();
                        str = Regex.Replace(str, @"[^\d]", "");
                        String stf = Regex.Replace(str.Replace("-", "/"), @"[^\d/]", "");

                        //   Console.WriteLine("String: " + r[x].ToString() + "\n Removed Spaces: " + str);

                        if (DateTime.TryParseExact(stf, "MMddyyyy", System.Globalization.CultureInfo.CurrentCulture, System.Globalization.DateTimeStyles.None, out dts))
                        {
                            String a = dts.ToString("MM/dd/yyyy");
                            r[x] = a;
                            //Console.WriteLine("String: " + r[x].ToString() + "\n Removed Spaces: " + str + "\n Date formatted: " + a + "\n");
                        }

                    }

                }

        }

        //ZIP 1 - Pad left6

        private void Zip_1()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;
            char pad = '0';

            foreach (DataRow r in dt.Select("[SET 1] = 'ZIP 1'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    if (r[x].ToString() != null && !r[x].ToString().Equals(""))
                    {
                        r[x] = r[x].ToString().PadLeft(6, '0');
                        // Console.WriteLine(r[x].ToString().PadLeft(8, '0'));

                    }


                }

        }

        //ZIP 2 - Pad left6

        private void Zip_2()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;
            char pad = '0';

            foreach (DataRow r in dt.Select("[SET 1] = 'ZIP 2'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    if (r[x].ToString() != null && !r[x].ToString().Equals(""))
                    {

                        r[x] = r[x].ToString().PadLeft(6, '0');
                        //  Console.WriteLine(r[x].ToString().PadLeft(6, '0'));
                    }


                }

        }

        //CUSTOMER ID  - Pad left12

        private void CUSTOMER_ID()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;
            char pad = '0';

            foreach (DataRow r in dt.Select("[SET 1] = 'CUSTOMER ID '"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    if (r[x].ToString() != null && !r[x].ToString().Equals(""))
                    {
                        r[x] = r[x].ToString().PadLeft(12, '0');
                        // Console.WriteLine(r[x].ToString().PadLeft(12, '0'));

                    }

                }

        }

        //CARD ISSUE TYPE - Proper Case

        private void CARD_ISSUE_TYPE()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'CARD ISSUE TYPE'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    r[x] = info.ToTitleCase(r[x].ToString().ToLower());
                    // Console.WriteLine(r[x].ToString().ToLower());
                }

        }

        //CARD HOLDER NAME - Proper Case

        private void CARD_HOLDER_NAME()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'CARD ISSUE TYPE'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    r[x] = info.ToTitleCase(r[x].ToString().ToLower());
                    // Console.WriteLine(r[x].ToString().ToLower());
                }

        }

        //CARD TYPE - Proper Case

        private void CARD_TYPE()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'CARD ISSUE TYPE'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    r[x] = info.ToTitleCase(r[x].ToString().ToLower());
                    //  Console.WriteLine(r[x].ToString().ToLower());
                }

        }

        //BENEFICIARY NAME - Proper Case

        private void BENEFICIARY_NAME()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'BENEFICIARY NAME'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    r[x] = info.ToTitleCase(r[x].ToString().ToLower());
                    //Console.WriteLine(r[x].ToString().ToLower());
                }

        }

        //COUNTRY - Proper Case

        private void COUNTRY()
        {
            TextInfo info = CultureInfo.CurrentCulture.TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'COUNTRY'"))
                for (int x = 1; x < r.ItemArray.Length; x++)
                {
                    r[x] = info.ToTitleCase(r[x].ToString().ToLower());
                    //Console.WriteLine(r[x].ToString().ToLower());
                }

        }





        //LIFE INSURANCE UPPER CASE

        private void LIFE_INSURANCE()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'LIFE INSURANCE'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {



                    String str = r[y].ToString();
                    r[y] = str.Replace('0', 'o');
                    r[y] = textInfo.ToUpper(r[y].ToString());

                    

                }
        }

        //REMARKS - UPPER CASE

        private void REMARKS()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'REMARKS'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    r[y] = textInfo.ToUpper(r[y].ToString());
                    //Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
                }
        }

        private void REMARK()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'REMARK'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    r[y] = textInfo.ToUpper(r[y].ToString());
                    //Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
                }
        }

        //SEX 1 - UPPER CASE

        private void SEX_1()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'SEX 1'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    r[y] = textInfo.ToUpper(r[y].ToString());
                    // Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
                }
        }

        //PROVINCE 2 - UPPER CASE

        private void PROVINCE_2()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'PROVINCE 2'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    r[y] = textInfo.ToUpper(r[y].ToString());
                    //Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
                }
        }


        //BLOOD GROUP - UPPER CASE

        private void BLOOD_GROUP()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'BLOOD GROUP'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    r[y] = textInfo.ToUpper(r[y].ToString());
                    //Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
                }
        }


        //CARD LIMIT - Thousands,Decimal format with $ 

        private void CARD_LIMIT()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'CARD LIMIT'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    //r[y] = textInfo.ToUpper(r[y].ToString());
                    r[y] = String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]);
                    //Console.WriteLine(String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]));

                }
        }

        //AVERAGE MONTHLY USAGE - Thousands,Decimal format with $ 

        private void AVERAGE_MONTHLY_USAGE()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'AVERAGE MONTHLY USAGE'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    //r[y] = textInfo.ToUpper(r[y].ToString());
                    r[y] = String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]);
                    //Console.WriteLine(String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]));

                }
        }

        //AVERAGE MONTHLY PAYMENT - Thousands,Decimal format with $ 

        private void AVERAGE_MONTHLY_PAYMENT()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'AVERAGE MONTHLY PAYMENT'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    //r[y] = textInfo.ToUpper(r[y].ToString());
                    r[y] = String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]);
                    //Console.WriteLine(String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]));

                }
        }

        //FICO CREDIT SCORE - Thousands,Decimal format with $ 

        private void FICO_CREDIT_SCORE()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;

            foreach (DataRow r in dt.Select("[SET 1] = 'FICO CREDIT SCORE'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    //r[y] = textInfo.ToUpper(r[y].ToString());
                    r[y] = String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]);
                    //Console.WriteLine(String.Format(new System.Globalization.CultureInfo("en-US"), "{0:C}", r[y]));

                }
        }

        //RATE OF INTEREST - with % 

        private void RATE_OF_INTEREST()
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            DataTable dt = dataGridView1.DataSource as DataTable;


            foreach (DataRow r in dt.Select("[SET 1] = 'RATE OF INTEREST'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    if (r[y].ToString() != null && !r[y].ToString().Equals(""))
                    {
                        if (!r[y].ToString().Contains("%"))
                        {
                            //double d = double.Parse(r[y].t);
                            //r[y] = textInfo.ToUpper(r[y].ToString());
                            //   r[y] = d.ToString("F2", CultureInfo.InvariantCulture) + "%";
                            //    r[y] = String.Format("{0:P2}", r[y]); // formats as 85.26 % (varies by culture)
                            r[y] = r[y].ToString() + "%";
                            // Console.WriteLine(String.Format("{0:P2}", r[y]));


                        }
                     



                    }


                }
        }






        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            int total = 100; //some number (this is your variable to change)!!

            for (int i = 0; i <= total; i++) //some number (total)
            {
                System.Threading.Thread.Sleep(100);
                int percents = (i * 100) / total;
                bgw.ReportProgress(percents, i);
                //2 arguments:
                //1. procenteges (from 0 t0 100) - i do a calcumation 
                //2. some current value!


            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label3.Text = String.Format("Loading... {0} %", e.ProgressPercentage);
            //  label4.Text = String.Format("Total items transfered: {0}", e.UserState);
            label3.ForeColor = Color.FromArgb(255, 255, 255);

            if (e.ProgressPercentage.Equals(100))
            {
                label3.Text = "Done!";
                label3.ForeColor = Color.FromArgb(6, 176, 37);
                progressBar1.Value = 0;
            }

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            DisplayTable();
            //    if (e.ProgressPercentage.Equalse.ProgressPercentage)
        }

        private void button2_Click(object sender, EventArgs e)
        {








        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occurred while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


        void DisplayTable()
        {
            DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            dataGridView1.DataSource = dt;
          //  DataTable dts = dataGridView1.DataSource as DataTable;

            // DataTable dt = dataGridView1.DataSource as DataTable;

/*
            foreach (DataRow r in dts.Select("[SET 1] = 'RATE OF INTEREST'"))
                for (int y = 1; y < r.ItemArray.Length; y++)
                {
                    if (r[y].ToString() != null && !r[y].ToString().Equals(""))
                    {
                        if (r[y].ToString().Contains("%"))
                        {

                           
                            Console.WriteLine("colour gone");

                        }
                        else
                        {
                            Console.WriteLine("colour nottttt gone");
                             FormatACell((DataGridViewCell)r[y]);
                        }
                  



                    }


                } */



            //     this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
            //   this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (cboSheet.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Sheet", "Info");

            }
            else
            {
                /* using (Form2 frm = new Form2(DisplayTable))
                 {
                     frm.ShowDialog(this);

                 }
                */

                bgw.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
                bgw.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
                bgw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
                bgw.WorkerReportsProgress = true;
                bgw.RunWorkerAsync();



            }

        }



        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click_1(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Code to check
            if (txtFilename.TextLength == 0)
            {
                MessageBox.Show("File not selected", "Info");

            }
            else if (cboSheet.SelectedIndex == -1)
            {
                MessageBox.Show("Sheet not selected", "Info");
            }
            else if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Sheet not opened", "Info");

            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure that you want to close the file", "Info", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    txtFilename.Text = "";
                    cboSheet.Items.Clear();
                    cboSheet.Text = "";
                    dataGridView1.DataSource = null;
                    dataGridView1.Refresh();
                    label3.Text = "";
                    progressBar1.Value = 0;

                    button3.Enabled = true;
                    button5.Enabled = true;
                    button6.Enabled = true;
                }
            }


        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {

                using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel Files|*.xls;*.xlsx;*.xlsm" })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        txtFilename.Text = openFileDialog.FileName;
                        using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                                });
                                tableCollection = result.Tables;
                                cboSheet.Items.Clear();
                                foreach (DataTable table in tableCollection)
                                    cboSheet.Items.Add(table.TableName);//add sheet to combo box
                                                                        //     this.dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
                                                                        //     this.dataGridView1.AlternatingRowsDefaultCellStyle.BackColor =
                                                                        // Color.Beige;
                                                                        //  dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Orange;
                                                                        // dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Red;
                                                                        // dataGridView1.Rows.Col.BackColor = Color.Black;
                                                                        //  dataGridView1.EnableHeadersVisualStyles = false;
                                                                        //  dataGridView1.Columns["Fields"].DefaultCellStyle.ForeColor = Color.Gray;

                                //FormatACell();

                            }
                        }
                    }

                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("File Not Found", "Alert");
            }
            catch (IOException)
            {
                txtFilename.Text = "";
                MessageBox.Show("The file is already being used by another instance. Please close the file and try again.", "Alert");
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (txtFilename.TextLength == 0)
            {
                MessageBox.Show("Please select a valid file", "Alert");

            }
            else if (cboSheet.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Sheet", "Info");
            }
            else if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Please open the Sheet", "Info");

            }
            else
            {
                try
                {
                    //Code to check
                    CARD_ISSUE_TYPE();
                    CARD_HOLDER_NAME();
                    CARD_TYPE();
                    BENEFICIARY_NAME();
                    COUNTRY();
                    LIFE_INSURANCE();
                    REMARKS();
                    REMARK();
                    SEX_1();
                    PROVINCE_2();
                    BLOOD_GROUP();
                    CARD_LIMIT();
                    FICO_CREDIT_SCORE();
                    AVERAGE_MONTHLY_PAYMENT();
                    AVERAGE_MONTHLY_USAGE();
                    RATE_OF_INTEREST();
                    Zip_1();
                    Zip_2();
                    CUSTOMER_ID();

                    try {
                        CARD_ISSUER_DATE();
                        CARD_HOLDER_DOB();
                        BENEFICIARY_DOB();
                        Console.WriteLine("date sucess");
                    } catch {
                        Console.WriteLine("date failiure"); 
                    }
                  

                    button3.Enabled = false;
                    button5.Enabled = false;
                    button6.Enabled = false;

                    //  dataGridView1.CellEndEdit += new DataGridViewCellEventHandler(dataGridView1_CellEndEdit);
                    //  dataGridView1.RowsAdded += new DataGridViewRowsAddedEventHandler(dataGridView1_RowsAdded);
                  
                    MessageBox.Show("Changes Applied Successfully", "Info");
                }
                catch (Exception ex)
                {
                    //Code here if an error


                    txtFilename.Text = "";
                    cboSheet.Items.Clear();
                    cboSheet.Text = "";
                    dataGridView1.DataSource = null;
                    dataGridView1.Refresh();
                    label3.Text = "";
                    progressBar1.Value = 0;
                    MessageBox.Show("There was a problem that occurred while correcting data, Please open a valid file or Try Again.", "Alert");
                }

                //Code that I want to run if it's all OK ?? 






            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            try
            {
                //Code to check
                if (txtFilename.TextLength == 0)
                {
                    MessageBox.Show("Please select a valid file", "Info");

                }
                else if (cboSheet.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select a Sheet", "Info");
                }
                else if (dataGridView1.Rows.Count == 0)
                {
                    MessageBox.Show("Please open the Sheet", "Info");

                }
                else
                {

                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Excel Documents (*.xls)|*.xls";
                    sfd.FileName = "Export.xls";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        // Copy DataGridView results to clipboard
                        copyAlltoClipboard();

                        object misValue = System.Reflection.Missing.Value;
                        Excel.Application xlexcel = new Excel.Application();

                        xlexcel.DisplayAlerts = false; // Without this you will get two confirm overwrite prompts
                        Excel.Workbook xlWorkBook = xlexcel.Workbooks.Add(misValue);
                        Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                        // Format column D as text before pasting results, this was required for my data
                        Excel.Range rng = xlWorkSheet.get_Range("D:D").Cells;
                        rng.NumberFormat = "@";

                        // Paste clipboard results to worksheet range
                        Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                        CR.Select();
                        xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);

                        // For some reason column A is always blank in the worksheet. ¯\_(ツ)_/¯
                        // Delete blank column A and select cell A1
                        Excel.Range delRng = xlWorkSheet.get_Range("A:A").Cells;
                        delRng.Delete(Type.Missing);
                        xlWorkSheet.get_Range("A1").Select();
                        xlWorkSheet.Columns.AutoFit();
                        xlWorkSheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                      
                        //xlWorkSheet.ColumnHeadersDefaultCellStyle.BackColor = Color.Orange;
                      //  xlWorkSheet.RowHeadersDefaultCellStyle.BackColor = Color.Red;

                        //  xlWorkSheet.Columns.Range["B5:J4"].Style.Color = Color.Red;
                        // xlWorkSheet.Columns.hea




                     //   rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                         //  var columnHeadingsRange = xlWorkSheet.Range[
                        //     xlWorkSheet.Cells["A1", "G1"]];           
                       //  columnHeadingsRange.Interior.Color = Excel.XlRgbColor.rgbLightGoldenrodYellow;
                      //  columnHeadingsRange.Style.Font.Bold = true;
                     //     dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Orange;
                     // dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.Red;

                        //xlWorkSheet.Range["A1", "G1"].Interior.Color = Excel.XlRgbColor.rgbDarkBlue;
                        //   xlWorkSheet.Range["A1", "G1"].Font.Color = Excel.XlRgbColor.rgbWhite;
                        // xlWorkSheet.Range["A1:D1"].Style.Color = Color.LightSeaGreen;





                        // Save the excel file under the captured location from the SaveFileDialog
                        xlWorkBook.SaveAs(sfd.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlexcel.DisplayAlerts = true;
                        xlWorkBook.Close(true, misValue, misValue);
                        xlexcel.Quit();

                        releaseObject(xlWorkSheet);
                        releaseObject(xlWorkBook);
                        releaseObject(xlexcel);

                        // Clear Clipboard and DataGridView selection
                        Clipboard.Clear();
                        dataGridView1.ClearSelection();

                        // Open the newly saved excel file
                        if (File.Exists(sfd.FileName))
                            System.Diagnostics.Process.Start(sfd.FileName);

                        //reset application
                        MessageBox.Show("File Exported Successfully", "Info");
                        txtFilename.Text = "";
                        cboSheet.Items.Clear();
                        cboSheet.Text = "";
                        dataGridView1.DataSource = null;
                        dataGridView1.Refresh();
                        label3.Text = "";
                        progressBar1.Value = 0;

                        button3.Enabled = true;
                        button5.Enabled = true;
                        button6.Enabled = true;

                    }

                }
            }
            catch (Exception ex)
            {
                //Code here if an error
                MessageBox.Show("There was a problem that occurred while exporting, Please Try Again.", "Alert");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show("Are you sure you want to close this form?", "Confirmation ", MessageBoxButtons.YesNo);
            if (dr == DialogResult.No) e.Cancel = true;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
           
        }

        private void FormatACell(DataGridViewCell cell)
        {
            if (cell.Value != null && cell.Value.ToString().Contains("%"))
            {
                cell.Style.BackColor = Color.Red;
            }
            else
            {
                cell.Style.BackColor = Color.White;
            }
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
            foreach (DataGridViewCell cell in row.Cells)
            {
                FormatACell(cell);
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            FormatACell(cell);
        }
    }
}

