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

namespace DetentionManageApp
{
    public partial class FormCreateEdit : Form
    {
        public DataTable VehicleData { get; private set; }
        private bool isEditMode;
        private string existingMaTamGiam;

        public FormCreateEdit()
        {
            InitializeComponent();
            isEditMode = false;
            txtMaTamGiam.ReadOnly = true;  // MaTamGiam is read-only
            GenerateMaTamGiam();
        }

        public FormCreateEdit(DataGridViewRow selectedRow)
        {
            InitializeComponent();
            LoadData(selectedRow);
            isEditMode = true;
            existingMaTamGiam = txtMaTamGiam.Text;
            txtMaTamGiam.ReadOnly = true;
        }

        private void LoadData(DataGridViewRow selectedRow)
        {
            txtMaTamGiam.Text = selectedRow.Cells["MaTamGiam"].Value.ToString();
            txtHo.Text = selectedRow.Cells["Ho"].Value.ToString();
            txtTen.Text = selectedRow.Cells["Ten"].Value.ToString();
            txtCCCD.Text = selectedRow.Cells["CCCD"].Value.ToString();
            txtSDT.Text = selectedRow.Cells["SDT"].Value.ToString();
            txtDiaChi.Text = selectedRow.Cells["DiaChi"].Value.ToString();
            txtBienSoXe.Text = selectedRow.Cells["BienSo"].Value.ToString();
            dtpNgayBatDauTamGiam.Value = DateTime.Parse(selectedRow.Cells["NgayBatDauTamGiam"].Value.ToString());
            dtpNgayKetThucTamGiam.Value = DateTime.Parse(selectedRow.Cells["NgayKetThucTamGiam"].Value.ToString());
        }

        private void GenerateMaTamGiam()
        {
            var filePath = "path_to_your_excel_file.xlsx";  // Replace with the path to your Excel file
            using (var package = new OfficeOpenXml.ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    txtMaTamGiam.Text = "MTG00001";
                    return;
                }

                int lastRow = worksheet.Dimension.End.Row;
                if (lastRow < 2)  // No data rows
                {
                    txtMaTamGiam.Text = "MTG00001";
                    return;
                }

                var lastMaTamGiam = worksheet.Cells[lastRow, 2].Value.ToString();  // Assuming MaTamGiam is in the 2nd column
                int newMaTamGiam = int.Parse(lastMaTamGiam.Substring(3)) + 1; // Extract the numeric part and increment
                txtMaTamGiam.Text = "MTG" + newMaTamGiam.ToString("D5");
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime ngayBatDau = dtpNgayBatDauTamGiam.Value;
                DateTime ngayKetThuc = dtpNgayKetThucTamGiam.Value;

                if (ngayBatDau >= ngayKetThuc)
                {
                    MessageBox.Show("[Ngày bắt đầu tạm giam] phải nhỏ hơn [Ngày kết thúc tạm giam]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Collect data and close form
                VehicleData = new DataTable();
                VehicleData.Columns.Add("NgayNhapThongTin");
                VehicleData.Columns.Add("MaTamGiam");
                VehicleData.Columns.Add("Ho");
                VehicleData.Columns.Add("Ten");
                VehicleData.Columns.Add("CCCD");
                VehicleData.Columns.Add("SDT");
                VehicleData.Columns.Add("DiaChi");
                VehicleData.Columns.Add("BienSo");
                VehicleData.Columns.Add("NgayBatDauTamGiam");
                VehicleData.Columns.Add("NgayKetThucTamGiam");

                DataRow row = VehicleData.NewRow();
                row["NgayNhapThongTin"] = DateTime.Now.ToString("yyyy-MM-dd");
                row["MaTamGiam"] = txtMaTamGiam.Text;
                row["Ho"] = txtHo.Text;
                row["Ten"] = txtTen.Text;
                row["CCCD"] = txtCCCD.Text;
                row["SDT"] = txtSDT.Text;
                row["DiaChi"] = txtDiaChi.Text;
                row["BienSo"] = txtBienSoXe.Text;
                row["NgayBatDauTamGiam"] = dtpNgayBatDauTamGiam.Value.ToString("yyyy-MM-dd");
                row["NgayKetThucTamGiam"] = dtpNgayKetThucTamGiam.Value.ToString("yyyy-MM-dd");

                VehicleData.Rows.Add(row);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu thông tin: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

}
