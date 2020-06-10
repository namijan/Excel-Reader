using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            label1.Text = openFileDialog1.FileName;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (label1.Text == "")
                MessageBox.Show("Please select a file first", "Error");
            else
            {
                try
                {
                    //create a instance for the Excel object  
                    Excel.Application oExcel = new Excel.Application();
                    //specify the file name where its actually exist  
                    string filepath = label1.Text;
                    //pass that to workbook object  

                    MessageBox.Show("I will now attempt to read the file after you hit ok");
                    Excel.Workbook WB = oExcel.Workbooks.Open(filepath);

                    MessageBox.Show("Looks like I am able to read the file. I will proceed to give you some relevant details now");

                    //  get the workbookname  
                    string ExcelWorkbookname = WB.Name;

                    //  get the worksheet count  
                    int worksheetcount = WB.Worksheets.Count;

                    Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];

                    //  get the firstworksheetname  
                    string firstworksheetname = wks.Name;

                    //statement get the first cell value  
                    var firstcellvalue = ((Excel.Range)wks.Cells[1, 1]).Value;


                    MessageBox.Show(

                        "Workbook name = " + ExcelWorkbookname + '\n' +
                        "Count of worksheets = " + worksheetcount + '\n' +
                        "Name of the first worksheet = " + firstworksheetname + '\n' +
                        "First cell value = " + firstcellvalue, "READ WAS SUCCESSFUL");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("I caught an exception. Hit ok to get details");
                    string error = ex.Message;
                    MessageBox.Show(error, "READ WAS UNSUCCESSFUL");
                }
            }
        }
    }
}
