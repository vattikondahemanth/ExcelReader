using OfficeOpenXml.Table;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExcelReader.Form3;
using System.Collections;
using System.IO;
using static System.Windows.Forms.AxHost;
using System.Data.SqlClient;

namespace ExcelReader
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        public void LoadData() 
        {
            var excelcombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["ExcelColumn"];
            excelcombobox.DataSource = Program.Get_Excel_Column_Names();


            var tablecombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["SQLColumn"];
            tablecombobox.DataSource = Program.Get_Table_Column_Names();

            var clearLink = (DataGridViewLinkColumn)dataGridView1.Columns["Clear"];
            clearLink.Name = "Clear";
            clearLink.DataPropertyName= "ClearID";
            clearLink.Text = "Clear";
            clearLink.Visible = true;
            clearLink.ToolTipText = "clear mappings";
            clearLink.UseColumnTextForLinkValue = true;


            List<MappingDetails> mapData = new List<MappingDetails>();

            for (int i = 0; i < Program.Get_Excel_Column_Names().Count; i++)
            {
                mapData.Add(new MappingDetails() { });
            }

            dataGridView1.DataSource = mapData;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells["InsertNull"].Value = "true";
            }

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
            if (dataGridView1.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && e.RowIndex >= 0 && e.ColumnIndex == 0)
            {
                DataGridViewComboBoxCell comboBoxCell = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (comboBoxCell.Value is null)
                {
                    return;
                }
                string selectedValue = comboBoxCell.Value.ToString();

                string colType = Program.GetSQLColumnType(selectedValue);
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Value = colType;

            }
            
            if (dataGridView1.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && e.RowIndex >= 0 && e.ColumnIndex == 2)
            {

                DataGridViewTextBoxCell sqlColTypecell = (DataGridViewTextBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 1];
                DataGridViewComboBoxCell comboBoxCell = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (comboBoxCell.Value is null)
                {
                    return;
                }
                if (sqlColTypecell.Value is null) 
                {
                    comboBoxCell.Value = null;
                    MessageBox.Show("Warning: \n Please select a sql column first");
                    return;
                }
                string sqlColType = sqlColTypecell.Value.ToString();
                string selectedValue = comboBoxCell.Value.ToString();
                               

                string colType = Program.GetColumnType(selectedValue, sqlColType);
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex + 1].Value = colType;
            }
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

     

        private void Form4_VisibleChanged(object sender, EventArgs e)
        {
            var excelcombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["ExcelColumn"];
            excelcombobox.DataSource = null;


            var tablecombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["SQLColumn"];
            tablecombobox.DataSource = null;           
            

            dataGridView1.DataSource = null;
            dataGridView1.Refresh();

            LoadData();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            rdbtnError.Checked = true;
        }

        private void btnFinish_Click(object sender, EventArgs e)
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
            lblMessage.Text = "Processing. Please wait ....";
            lblMessage.Visible = true;
            lblMessage.Enabled = true;
            if(!validate_empty_mapping()) return;
            if(!validate_duplicate_mapping()) return;
            (List<string> sqlColNames, List<string> excelColNames) = get_sql_columns_in_excel_order();
            if (rdbtnError.Checked)
            {
                insert_bulk_data(sqlColNames, excelColNames);
            }
            if (rdbtnIgnore.Checked)
            {
                insert_row_data(sqlColNames, excelColNames);
            }
            lblMessage.Text = "";
            lblMessage.Visible = false;
            lblMessage.Enabled = false;
        }

        private bool validate_empty_mapping()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewComboBoxCell sqlcolumn = (DataGridViewComboBoxCell)row.Cells[0];
                DataGridViewComboBoxCell excelcolumn = (DataGridViewComboBoxCell)row.Cells[2];

                if ((sqlcolumn.Value != null && excelcolumn.Value == null) || (sqlcolumn.Value != null && excelcolumn.Value == null))
                {
                    MessageBox.Show("Warning: \n Please complete all the mappings");
                    return false;
                }
            }
            return true;
        }
        private bool validate_duplicate_mapping()
        {
            List<string> sqlColNames = new List<string>();
            List<string> excelColNames = new List<string>();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewComboBoxCell sqlcolumn = (DataGridViewComboBoxCell)row.Cells[0];
                DataGridViewComboBoxCell excelcolumn = (DataGridViewComboBoxCell)row.Cells[2];
                if (sqlcolumn.Value == null && excelcolumn.Value == null)
                {
                    continue;
                }
                else 
                {
                    sqlColNames.Add(sqlcolumn.Value.ToString());
                    excelColNames.Add(excelcolumn.Value.ToString());

                }
            }
            if ((sqlColNames.GroupBy(x => x).Any(g => g.Count() > 1)) || (excelColNames.GroupBy(x => x).Any(g => g.Count() > 1)))
            {
                MessageBox.Show("Warning: \n Please remove duplicate mappings");
                return false;
            }
            return true;
        }
        private (List<string> , List<string>) get_sql_columns_in_excel_order()
        {
            List<string> sqlColNames = new List<string>();
            List<string> excelColNames = new List<string>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewComboBoxCell sqlcolumn = (DataGridViewComboBoxCell)row.Cells[0];
                DataGridViewComboBoxCell excelcolumn = (DataGridViewComboBoxCell)row.Cells[2];

                if (sqlcolumn.Value == null && excelcolumn.Value == null)
                {
                    continue;
                }
                else
                {
                    sqlColNames.Add(sqlcolumn.Value.ToString());
                    excelColNames.Add(excelcolumn.Value.ToString());

                }
            }
            return (sqlColNames, excelColNames);
        }
     
        private void insert_bulk_data(List<string> sqlColNames, List<string> excelColNames)
        {
            using (SqlConnection connection = new SqlConnection(Program._sqlconnection))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();

                try
                {
                    using (var package = new ExcelPackage(new FileInfo(Program._excelPath)))
                    {
                        var worksheet = package.Workbook.Worksheets[Program._sheetName];
                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            bool IsSuccessful = false;
                            string errorMessage = "";
                            
                            

                            try
                            {
                                List<string> insert_values = new List<string>();
                                string insert_query = get_insert_query(sqlColNames);

                                for (int i = 0; i < excelColNames.Count; i++)
                                {
                                    int idx = worksheet.Cells["1:1"].First(c => c.Value.ToString() == excelColNames[i]).Start.Column;
                                    var x = worksheet.Cells[row, idx].Text;

                                    int rowIndex, colIndex = -1;
                                    (rowIndex, colIndex) = (excelColNames.IndexOf(excelColNames[i]), 4);

                                    bool isNullAllowed = dataGridView1.Rows[rowIndex].Cells[colIndex].Value.ToString() == "true" ? true: false;

                                    if ((string.IsNullOrEmpty(x) | string.IsNullOrWhiteSpace(x)) & (isNullAllowed))
                                    {
                                        insert_values.Add("null");
                                    }
                                    else
                                    {
                                        insert_values.Add("'" + x.Replace("'", "''") + "'");
                                    }
                                }
                                var commaSeparatedString = string.Join(",", insert_values);
                                insert_query = string.Format(insert_query, commaSeparatedString);

                                (IsSuccessful, errorMessage) = Program.run_bulk_insert(insert_query, connection, transaction);
                            }
                            catch (Exception ex)
                            {
                                (IsSuccessful, errorMessage) = (false, ex.Message);
                            }
                            if (!IsSuccessful)
                            {
                                transaction.Rollback();
                                MessageBox.Show(errorMessage);
                                return;
                            }

                        }

                        transaction.Commit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
        }
        private void insert_row_data(List<string> sqlColNames, List<string> excelColNames) 
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

                ExcelWorksheet successWorksheet = package.Workbook.Worksheets.Add("Success");
                ExcelWorksheet errorWorksheet = package.Workbook.Worksheets.Add("Error");

                using (SqlConnection connection = new SqlConnection(Program._sqlconnection))
                {
                    connection.Open();
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        List<string> insert_values = new List<string>();
                        string insert_query = get_insert_query(sqlColNames);

                        bool IsSuccessful = false;
                        string errorMessage = "";

                        for (int i = 0; i < excelColNames.Count; i++)
                        {
                            int idx = worksheet.Cells["1:1"].First(c => c.Value.ToString() == excelColNames[i]).Start.Column;
                            var x = worksheet.Cells[row, idx].Text;

                            int rowIndex, colIndex = -1;
                            (rowIndex, colIndex) = (excelColNames.IndexOf(excelColNames[i]), 4);

                            bool isNullAllowed = dataGridView1.Rows[rowIndex].Cells[colIndex].Value.ToString() == "true" ? true : false;

                            if ((string.IsNullOrEmpty(x) | string.IsNullOrWhiteSpace(x)) & (isNullAllowed))
                            {
                                insert_values.Add("null");
                            }
                            else 
                            {
                                insert_values.Add("'" + x.Replace("'", "''") + "'");
                            }
                            
                        }
                        var commaSeparatedString = string.Join(",", insert_values);
                        insert_query = string.Format(insert_query, commaSeparatedString);

                        try
                        {
                            
                            (IsSuccessful, errorMessage) = Program.run_insert(insert_query, connection);
                        }
                        catch (Exception ex)
                        {
                            (IsSuccessful, errorMessage) = (false, ex.Message);
                        }

                       
                        for (int i = 0; i < excelColNames.Count; i++)
                        {
                            int idx = worksheet.Cells["1:1"].First(c => c.Value.ToString() == excelColNames[i]).Start.Column;
                            if (row == 2)
                            {
                                successWorksheet.Cells[1, i + 1].Value = excelColNames[i];
                                errorWorksheet.Cells[1, i + 1].Value = excelColNames[i];
                            }

                            if (IsSuccessful)
                            {
                                successWorksheet.Cells[row, i + 1].Value = worksheet.Cells[row, idx].Text;
                            }
                            else
                            {
                                errorWorksheet.Cells[row, i + 1].Value = worksheet.Cells[row, idx].Text;
                            }
                        }
                        if (IsSuccessful)
                        {
                            successWorksheet.Cells[row, excelColNames.Count + 1].Value = errorMessage;
                        }
                        else
                        {
                            errorWorksheet.Cells[row, excelColNames.Count + 1].Value = errorMessage;
                        }


                        package.Save();


                    }

                }
                
            }
        }

        private string get_insert_query(List<string> sqlColNames) 
        {
            StringBuilder stringBuilder= new StringBuilder();
            stringBuilder.Append("insert into " + Program._tableName + "(");
            var commaSeparatedString = string.Join(",", sqlColNames);
            stringBuilder.Append(commaSeparatedString);
            stringBuilder.Append(") values({0});");
            return stringBuilder.ToString();

        }
        private void btnBack_Click(object sender, EventArgs e)
        {
            var excelcombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["ExcelColumn"];
            excelcombobox.DataSource = null;


            var tablecombobox = (DataGridViewComboBoxColumn)dataGridView1.Columns["SQLColumn"];
            tablecombobox.DataSource = null;


            dataGridView1.DataSource = null;
            dataGridView1.Refresh();

            this.Hide();
            Program.form3.Show();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["Clear"].Index && e.RowIndex >= 0)
            {
                DataGridViewTextBoxCell excelColTypecell = (DataGridViewTextBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 1];

                DataGridViewComboBoxCell excelColcell = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 2];

                DataGridViewTextBoxCell sqlColTypecell = (DataGridViewTextBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 3];

                DataGridViewComboBoxCell sqlColcell = (DataGridViewComboBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex - 4];

                excelColTypecell.Value = null;
                excelColcell.Value = null;
                sqlColTypecell.Value = null;
                sqlColcell.Value = null;

            }
            
            
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }
    }

    public class MappingDetails
    {
        public string ExcelColumnID = "<Not Found>";
        public string ExcelColumnType = "<None>";

        public string SQLColumnID = "<Not Found>";
        public string SQLColumnType = "<None>";

        public string OnErrorID;
        public string ClearID  = "&Clear";
        public int InsertNull = 1;

    }
}
