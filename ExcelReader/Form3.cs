using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ExcelReader
{
   
    public partial class Form3 : Form
    {

        public Form3()
        {
            InitializeComponent();
            this.CenterToScreen();

        }

     
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Hide();
            Program.form2.Show();
        }

        private void Form3_Load(object sender, EventArgs e)
        {

            using (var package = new ExcelPackage(new FileInfo(Program._excelPath)))
            {

                var worksheet = package.Workbook.Worksheets[Program._sheetName];

                ExcelWorksheet worksheetToDelete = package.Workbook.Worksheets["Success"];
                if (!(worksheetToDelete == null))
                {
                    package.Workbook.Worksheets.Delete(worksheetToDelete);
                    package.Save();

                }
                worksheetToDelete = package.Workbook.Worksheets["Error"];

                if (!(worksheetToDelete == null))
                {
                    package.Workbook.Worksheets.Delete(worksheetToDelete);
                    package.Save();
                }
            }


            var excelcombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["ExcelSheet"];
            excelcombobox.DataSource = Program.Get_Excel_Sheets();
            

            var tablecombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["Table"];
            tablecombobox.DataSource = Program.Get_Tables();


            dataGridView1.DataSource = new List<MyData>(){ new MyData() { ExcelSheetNames="", TableNames="" } };

        }
      

        private void btnNext_Click(object sender, EventArgs e)
        {
            object source = dataGridView1.Rows[0].Cells[0].Value;
            object target = dataGridView1.Rows[0].Cells[1].Value;
            if (source is null || target is null)
            {
                MessageBox.Show("Warning: \n Please select Source and Target");
                return;
            }
            Program._sheetName = source.ToString();
            Program._tableName = target.ToString(); 
            MessageBox.Show("Success! \n Source: " + Program._sheetName + "\n Target: " + Program._tableName);
            this.Hide();
            Program.form4.Show();


        }

        public class MyData
        {
            public string ExcelSheetNames;
            public string TableNames;
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }
    }
}
