namespace ExcelReader
{
    partial class Form4
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lblMappings = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.dataGridViewImageColumn1 = new System.Windows.Forms.DataGridViewImageColumn();
            this.btnFinish = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.rdbtnIgnore = new System.Windows.Forms.RadioButton();
            this.rdbtnError = new System.Windows.Forms.RadioButton();
            this.lblMessage = new System.Windows.Forms.Label();
            this.SQLColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.SQLColumnType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExcelColumn = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.ExcelColumnType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.InsertNull = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Clear = new System.Windows.Forms.DataGridViewLinkColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.SQLColumn,
            this.SQLColumnType,
            this.ExcelColumn,
            this.ExcelColumnType,
            this.InsertNull,
            this.Clear});
            this.dataGridView1.GridColor = System.Drawing.SystemColors.Window;
            this.dataGridView1.Location = new System.Drawing.Point(58, 73);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(693, 505);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
            this.dataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView1_DataError);
            // 
            // lblMappings
            // 
            this.lblMappings.AutoSize = true;
            this.lblMappings.Location = new System.Drawing.Point(64, 43);
            this.lblMappings.Name = "lblMappings";
            this.lblMappings.Size = new System.Drawing.Size(56, 13);
            this.lblMappings.TabIndex = 1;
            this.lblMappings.Text = "Mappings:";
            // 
            // btnBack
            // 
            this.btnBack.Location = new System.Drawing.Point(446, 612);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 2;
            this.btnBack.Text = "< &Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // dataGridViewImageColumn1
            // 
            this.dataGridViewImageColumn1.HeaderText = "Status";
            this.dataGridViewImageColumn1.Name = "dataGridViewImageColumn1";
            // 
            // btnFinish
            // 
            this.btnFinish.Location = new System.Drawing.Point(550, 612);
            this.btnFinish.Name = "btnFinish";
            this.btnFinish.Size = new System.Drawing.Size(75, 23);
            this.btnFinish.TabIndex = 3;
            this.btnFinish.Text = "&Finish >";
            this.btnFinish.UseVisualStyleBackColor = true;
            this.btnFinish.Click += new System.EventHandler(this.btnFinish_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(657, 612);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // rdbtnIgnore
            // 
            this.rdbtnIgnore.AutoSize = true;
            this.rdbtnIgnore.Location = new System.Drawing.Point(318, 39);
            this.rdbtnIgnore.Name = "rdbtnIgnore";
            this.rdbtnIgnore.Size = new System.Drawing.Size(55, 17);
            this.rdbtnIgnore.TabIndex = 5;
            this.rdbtnIgnore.TabStop = true;
            this.rdbtnIgnore.Text = "Ignore";
            this.rdbtnIgnore.UseVisualStyleBackColor = true;
            // 
            // rdbtnError
            // 
            this.rdbtnError.AutoSize = true;
            this.rdbtnError.Location = new System.Drawing.Point(416, 39);
            this.rdbtnError.Name = "rdbtnError";
            this.rdbtnError.Size = new System.Drawing.Size(47, 17);
            this.rdbtnError.TabIndex = 6;
            this.rdbtnError.TabStop = true;
            this.rdbtnError.Text = "Error";
            this.rdbtnError.UseVisualStyleBackColor = true;
            // 
            // lblMessage
            // 
            this.lblMessage.AutoSize = true;
            this.lblMessage.Location = new System.Drawing.Point(64, 612);
            this.lblMessage.Name = "lblMessage";
            this.lblMessage.Size = new System.Drawing.Size(0, 13);
            this.lblMessage.TabIndex = 7;
            // 
            // SQLColumn
            // 
            this.SQLColumn.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.SQLColumn.Frozen = true;
            this.SQLColumn.HeaderText = "SQLColumn";
            this.SQLColumn.Name = "SQLColumn";
            this.SQLColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.SQLColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // SQLColumnType
            // 
            this.SQLColumnType.Frozen = true;
            this.SQLColumnType.HeaderText = "SQLColumnType";
            this.SQLColumnType.Name = "SQLColumnType";
            // 
            // ExcelColumn
            // 
            this.ExcelColumn.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.ExcelColumn.Frozen = true;
            this.ExcelColumn.HeaderText = "ExcelColumn";
            this.ExcelColumn.Name = "ExcelColumn";
            this.ExcelColumn.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.ExcelColumn.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // ExcelColumnType
            // 
            this.ExcelColumnType.Frozen = true;
            this.ExcelColumnType.HeaderText = "ExcelColumnType";
            this.ExcelColumnType.Name = "ExcelColumnType";
            // 
            // InsertNull
            // 
            this.InsertNull.DataPropertyName = "InsertNull";
            this.InsertNull.FalseValue = "false";
            this.InsertNull.Frozen = true;
            this.InsertNull.HeaderText = "InsertNull";
            this.InsertNull.Name = "InsertNull";
            this.InsertNull.TrueValue = "true";
            // 
            // Clear
            // 
            this.Clear.DataPropertyName = "ClearID";
            this.Clear.Frozen = true;
            this.Clear.HeaderText = "Clear";
            this.Clear.Name = "Clear";
            this.Clear.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.Clear.ToolTipText = "clear mappings";
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(945, 667);
            this.Controls.Add(this.lblMessage);
            this.Controls.Add(this.rdbtnError);
            this.Controls.Add(this.rdbtnIgnore);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnFinish);
            this.Controls.Add(this.btnBack);
            this.Controls.Add(this.lblMappings);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form4";
            this.Text = "Form4";
            this.Load += new System.EventHandler(this.Form4_Load);
            this.VisibleChanged += new System.EventHandler(this.Form4_VisibleChanged);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lblMappings;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.DataGridViewImageColumn dataGridViewImageColumn1;
        private System.Windows.Forms.Button btnFinish;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.RadioButton rdbtnIgnore;
        private System.Windows.Forms.RadioButton rdbtnError;
        private System.Windows.Forms.Label lblMessage;
        private System.Windows.Forms.DataGridViewComboBoxColumn SQLColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn SQLColumnType;
        private System.Windows.Forms.DataGridViewComboBoxColumn ExcelColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn ExcelColumnType;
        private System.Windows.Forms.DataGridViewCheckBoxColumn InsertNull;
        private System.Windows.Forms.DataGridViewLinkColumn Clear;
    }
}