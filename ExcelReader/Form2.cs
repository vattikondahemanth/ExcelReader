using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReader
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            this.CenterToScreen();
        }
        
        private void btnNext_Click(object sender, EventArgs e)
        {
            this.Hide();
            Program.form3.Show();
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            Program._sqlconnection = txtSqlString.Text;

            Program.CheckSQLConnection();

            MessageBox.Show("Connectoin Successful!");
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            this.Hide();
            Program.form1.Show();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
