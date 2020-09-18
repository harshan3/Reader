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
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Reader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataTableCollection tableCollection;

        private void button1_Click(object sender, EventArgs e)

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
                            }
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("File Not Found");
            }
            catch (IOException)
            {
                MessageBox.Show("Another user is already using this file.");
            }

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
            if (txtFilename.TextLength == 0)
            {
                MessageBox.Show("Please select a valid file");

            }
            else if (cboSheet.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Sheet");
            }
            else
            {
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
                    Console.WriteLine(r[x].ToString().ToLower());
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
                    Console.WriteLine(r[x].ToString().ToLower());
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
                    Console.WriteLine(r[x].ToString().ToLower());
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
                    Console.WriteLine(r[x].ToString().ToLower());
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
                    Console.WriteLine(r[x].ToString().ToLower());
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
                    r[y] = textInfo.ToUpper(r[y].ToString());
                    Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
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
                    Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
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
                    Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
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
                    Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
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
                    Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
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
                    Console.WriteLine(r[y].ToString()); //WAR AND PEACE}
                }
        }






        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!backgroundWorker1.CancellationPending)
            {

            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            label3.Text = String.Format("Loading...{0}",e.ProgressPercentage);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if(e.Error == null)
            {
                Thread.Sleep(100);
                label3.Text = "Data is loaded sucessfully";
                progressBar1.Update();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (txtFilename.TextLength == 0)
            {
                MessageBox.Show("Please select a valid file");

            }
            else if (cboSheet.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Sheet");
            }
            else if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Please open the Sheet");

            }
            else
            {

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Documents (*.xls)|*.xls";
                sfd.FileName = "Inventory_Adjustment_Export.xls";
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

                    // rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    /*
                                    var columnHeadingsRange = xlWorkSheet.Range[
                                        xlWorkSheet.Cells["1", "A"],
                                        xlWorkSheet.Cells["1", "Z"]];           
                                    columnHeadingsRange.Interior.Color = Excel.XlRgbColor.rgbLightGoldenrodYellow;
                                    columnHeadingsRange.Style.Font.Bold = true;
                    */



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
                }

            }



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
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (cboSheet.SelectedIndex == -1)
            {
                MessageBox.Show("Please select a Sheet");

            }
            else
            {
                /* using (Form2 frm = new Form2(DisplayTable))
                 {
                     frm.ShowDialog(this);

                 }
                */
             
                DisplayTable();
               
            }

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}

