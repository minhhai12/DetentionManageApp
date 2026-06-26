namespace DetentionManageApp
{
    partial class FormList
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormList));
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lblTitleDanhSach = new System.Windows.Forms.Label();
            this.txtSearch = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnClearSearch = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnFileExportTrichXuat = new System.Windows.Forms.Button();
            this.btnFileExportTamGiam = new System.Windows.Forms.Button();
            this.btnLocations = new System.Windows.Forms.Button();
            this.btnChooseWordFileTrichXuat = new System.Windows.Forms.Button();
            this.btnChooseWordFileTamGiam = new System.Windows.Forms.Button();
            this.btnChooseFile = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridView1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridView1.EnableHeadersVisualStyles = false;
            this.dataGridView1.Location = new System.Drawing.Point(21, 228);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowTemplate.Height = 30;
            this.dataGridView1.Size = new System.Drawing.Size(1140, 552);
            this.dataGridView1.TabIndex = 6;
            this.dataGridView1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellDoubleClick);
            this.dataGridView1.ColumnHeaderMouseClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_ColumnHeaderMouseClick);
            this.dataGridView1.SelectionChanged += new System.EventHandler(this.dataGridView1_SelectionChanged);
            this.dataGridView1.Sorted += new System.EventHandler(this.dataGridView1_Sorted);
            // 
            // lblTitleDanhSach
            // 
            this.lblTitleDanhSach.AutoSize = true;
            this.lblTitleDanhSach.Font = new System.Drawing.Font("Times New Roman", 26.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitleDanhSach.Location = new System.Drawing.Point(388, 43);
            this.lblTitleDanhSach.Name = "lblTitleDanhSach";
            this.lblTitleDanhSach.Size = new System.Drawing.Size(423, 40);
            this.lblTitleDanhSach.TabIndex = 0;
            this.lblTitleDanhSach.Text = "DANH SÁCH TẠM GIAM";
            // 
            // txtSearch
            // 
            this.txtSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSearch.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSearch.ForeColor = System.Drawing.SystemColors.GrayText;
            this.txtSearch.Location = new System.Drawing.Point(776, 25);
            this.txtSearch.Name = "txtSearch";
            this.txtSearch.Size = new System.Drawing.Size(330, 29);
            this.txtSearch.TabIndex = 3;
            this.txtSearch.Text = "Nhập từ khóa tìm kiếm...";
            this.txtSearch.TextChanged += new System.EventHandler(this.txtSearch_TextChanged);
            this.txtSearch.Enter += new System.EventHandler(this.txtSearch_Enter);
            this.txtSearch.Leave += new System.EventHandler(this.txtSearch_Leave);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.btnCreate);
            this.panel1.Controls.Add(this.btnClearSearch);
            this.panel1.Controls.Add(this.txtSearch);
            this.panel1.Controls.Add(this.btnEdit);
            this.panel1.Controls.Add(this.btnDelete);
            this.panel1.Location = new System.Drawing.Point(12, 133);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1160, 79);
            this.panel1.TabIndex = 4;
            // 
            // btnCreate
            // 
            this.btnCreate.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.btnCreate.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreate.Image = ((System.Drawing.Image)(resources.GetObject("btnCreate.Image")));
            this.btnCreate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnCreate.Location = new System.Drawing.Point(12, 15);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnCreate.Size = new System.Drawing.Size(125, 50);
            this.btnCreate.TabIndex = 0;
            this.btnCreate.Text = "Tạo mới";
            this.btnCreate.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnCreate.UseVisualStyleBackColor = false;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnClearSearch
            // 
            this.btnClearSearch.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClearSearch.Font = new System.Drawing.Font("Times New Roman", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearSearch.Image = ((System.Drawing.Image)(resources.GetObject("btnClearSearch.Image")));
            this.btnClearSearch.Location = new System.Drawing.Point(1100, 23);
            this.btnClearSearch.Margin = new System.Windows.Forms.Padding(0);
            this.btnClearSearch.Name = "btnClearSearch";
            this.btnClearSearch.Size = new System.Drawing.Size(47, 33);
            this.btnClearSearch.TabIndex = 4;
            this.btnClearSearch.UseVisualStyleBackColor = true;
            this.btnClearSearch.Click += new System.EventHandler(this.btnClearSearch_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnEdit.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEdit.Image = global::DetentionManageApp.Properties.Resources.icon_edit;
            this.btnEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnEdit.Location = new System.Drawing.Point(143, 14);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnEdit.Size = new System.Drawing.Size(141, 50);
            this.btnEdit.TabIndex = 1;
            this.btnEdit.Text = "Chỉnh sửa";
            this.btnEdit.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnEdit.UseVisualStyleBackColor = false;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnDelete.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDelete.Image = ((System.Drawing.Image)(resources.GetObject("btnDelete.Image")));
            this.btnDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDelete.Location = new System.Drawing.Point(290, 14);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnDelete.Size = new System.Drawing.Size(100, 50);
            this.btnDelete.TabIndex = 2;
            this.btnDelete.Text = "Xóa";
            this.btnDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(12, 218);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(1160, 572);
            this.flowLayoutPanel1.TabIndex = 5;
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.btnFileExportTrichXuat);
            this.panel2.Controls.Add(this.btnFileExportTamGiam);
            this.panel2.Controls.Add(this.btnLocations);
            this.panel2.Location = new System.Drawing.Point(12, 796);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1160, 73);
            this.panel2.TabIndex = 7;
            // 
            // btnFileExportTrichXuat
            // 
            this.btnFileExportTrichXuat.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnFileExportTrichXuat.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFileExportTrichXuat.Image = global::DetentionManageApp.Properties.Resources.icon_export_yellow;
            this.btnFileExportTrichXuat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFileExportTrichXuat.Location = new System.Drawing.Point(225, 10);
            this.btnFileExportTrichXuat.Name = "btnFileExportTrichXuat";
            this.btnFileExportTrichXuat.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnFileExportTrichXuat.Size = new System.Drawing.Size(208, 50);
            this.btnFileExportTrichXuat.TabIndex = 1;
            this.btnFileExportTrichXuat.Text = "Xuất lệnh trích xuất";
            this.btnFileExportTrichXuat.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnFileExportTrichXuat.UseVisualStyleBackColor = false;
            this.btnFileExportTrichXuat.Click += new System.EventHandler(this.btnFileExportTrichXuat_Click);
            // 
            // btnFileExportTamGiam
            // 
            this.btnFileExportTamGiam.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btnFileExportTamGiam.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFileExportTamGiam.Image = ((System.Drawing.Image)(resources.GetObject("btnFileExportTamGiam.Image")));
            this.btnFileExportTamGiam.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnFileExportTamGiam.Location = new System.Drawing.Point(11, 10);
            this.btnFileExportTamGiam.Name = "btnFileExportTamGiam";
            this.btnFileExportTamGiam.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnFileExportTamGiam.Size = new System.Drawing.Size(208, 50);
            this.btnFileExportTamGiam.TabIndex = 0;
            this.btnFileExportTamGiam.Text = "Xuất lệnh tạm giam";
            this.btnFileExportTamGiam.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnFileExportTamGiam.UseVisualStyleBackColor = false;
            this.btnFileExportTamGiam.Click += new System.EventHandler(this.btnFileExport_Click);
            // 
            // btnLocations
            // 
            this.btnLocations.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLocations.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLocations.Image = global::DetentionManageApp.Properties.Resources.icon_user;
            this.btnLocations.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnLocations.Location = new System.Drawing.Point(986, 10);
            this.btnLocations.Name = "btnLocations";
            this.btnLocations.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnLocations.Size = new System.Drawing.Size(161, 50);
            this.btnLocations.TabIndex = 2;
            this.btnLocations.Text = "QL Địa điểm";
            this.btnLocations.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnLocations.UseVisualStyleBackColor = true;
            this.btnLocations.Click += new System.EventHandler(this.btnLocations_Click);
            // 
            // btnChooseWordFileTrichXuat
            // 
            this.btnChooseWordFileTrichXuat.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChooseWordFileTrichXuat.Image = ((System.Drawing.Image)(resources.GetObject("btnChooseWordFileTrichXuat.Image")));
            this.btnChooseWordFileTrichXuat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnChooseWordFileTrichXuat.Location = new System.Drawing.Point(12, 68);
            this.btnChooseWordFileTrichXuat.Name = "btnChooseWordFileTrichXuat";
            this.btnChooseWordFileTrichXuat.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnChooseWordFileTrichXuat.Size = new System.Drawing.Size(223, 50);
            this.btnChooseWordFileTrichXuat.TabIndex = 2;
            this.btnChooseWordFileTrichXuat.Text = "Mẫu Lệnh Trích Xuất";
            this.btnChooseWordFileTrichXuat.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnChooseWordFileTrichXuat.UseVisualStyleBackColor = true;
            this.btnChooseWordFileTrichXuat.Click += new System.EventHandler(this.btnChooseWordFileTrichXuat_Click);
            // 
            // btnChooseWordFileTamGiam
            // 
            this.btnChooseWordFileTamGiam.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChooseWordFileTamGiam.Image = ((System.Drawing.Image)(resources.GetObject("btnChooseWordFileTamGiam.Image")));
            this.btnChooseWordFileTamGiam.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnChooseWordFileTamGiam.Location = new System.Drawing.Point(12, 12);
            this.btnChooseWordFileTamGiam.Name = "btnChooseWordFileTamGiam";
            this.btnChooseWordFileTamGiam.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnChooseWordFileTamGiam.Size = new System.Drawing.Size(223, 50);
            this.btnChooseWordFileTamGiam.TabIndex = 1;
            this.btnChooseWordFileTamGiam.Text = "Mẫu Lệnh Tạm Giam";
            this.btnChooseWordFileTamGiam.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnChooseWordFileTamGiam.UseVisualStyleBackColor = true;
            this.btnChooseWordFileTamGiam.Click += new System.EventHandler(this.btnChooseWordFile_Click);
            // 
            // btnChooseFile
            // 
            this.btnChooseFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnChooseFile.Font = new System.Drawing.Font("Times New Roman", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChooseFile.Image = ((System.Drawing.Image)(resources.GetObject("btnChooseFile.Image")));
            this.btnChooseFile.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnChooseFile.Location = new System.Drawing.Point(938, 33);
            this.btnChooseFile.Name = "btnChooseFile";
            this.btnChooseFile.Padding = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.btnChooseFile.Size = new System.Drawing.Size(234, 50);
            this.btnChooseFile.TabIndex = 3;
            this.btnChooseFile.Text = "Chọn file Excel dữ liệu ";
            this.btnChooseFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnChooseFile.UseVisualStyleBackColor = true;
            this.btnChooseFile.Click += new System.EventHandler(this.btnChooseFile_Click);
            // 
            // FormList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1184, 881);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.lblTitleDanhSach);
            this.Controls.Add(this.btnChooseWordFileTrichXuat);
            this.Controls.Add(this.btnChooseWordFileTamGiam);
            this.Controls.Add(this.btnChooseFile);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormList";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Danh sách tạm giam";
            this.Load += new System.EventHandler(this.FormList_Load);
            this.Resize += new System.EventHandler(this.FormList_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btnChooseFile;
        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnEdit;
        private System.Windows.Forms.Label lblTitleDanhSach;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.TextBox txtSearch;
        private System.Windows.Forms.Button btnClearSearch;
        private System.Windows.Forms.Button btnChooseWordFileTamGiam;
        private System.Windows.Forms.Button btnFileExportTamGiam;
        private System.Windows.Forms.Button btnLocations;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button btnFileExportTrichXuat;
        private System.Windows.Forms.Button btnChooseWordFileTrichXuat;
    }
}