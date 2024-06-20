using Newtonsoft.Json;
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
using static DetentionManageApp.FormCreateEdit;

namespace DetentionManageApp
{
    public partial class FormCreateEdit : Form
    {
        public DataTable detentionData { get; private set; }
        private readonly FormMode formMode;

        public enum FormMode
        {
            Create,
            Edit
        }

        public FormCreateEdit(FormMode mode)
        {
            InitializeComponent();
            formMode = mode;
            if (formMode == FormMode.Create)
            {
                GenerateMaTamGiam();
            }
            btnSave.Text = formMode == FormMode.Create ? "Tạo mới" : "Chỉnh sửa";
            btnCancel.Text = "Hủy bỏ";
            txtMaTamGiam.Enabled = false;
        }

        public FormCreateEdit(FormMode mode, DataGridViewRow selectedRow) : this(mode)
        {
            LoadData(selectedRow);
        }

        private void LoadData(DataGridViewRow selectedRow)
        {
            txtMaTamGiam.Text = selectedRow.Cells["Mã tạm giam"].Value.ToString();
            txtHo.Text = selectedRow.Cells["Họ"].Value.ToString();
            txtTen.Text = selectedRow.Cells["Tên"].Value.ToString();
            txtCCCD.Text = selectedRow.Cells["CCCD"].Value.ToString();
            txtSDT.Text = selectedRow.Cells["SĐT"].Value.ToString();
            txtDiaChi.Text = selectedRow.Cells["Địa chỉ"].Value.ToString();
            dtpNgayBatDauTamGiam.Value = DateTime.Parse(selectedRow.Cells["Ngày bắt đầu"].Value.ToString());
            dtpNgayKetThucTamGiam.Value = DateTime.Parse(selectedRow.Cells["Ngày kết thúc"].Value.ToString());
        }

        private void GenerateMaTamGiam()
        {
            string jsonFilePath = Path.Combine(Application.StartupPath, "excelFilePath.json");
            string filePath = "";
            if (File.Exists(jsonFilePath))
            {
                string jsonContent = File.ReadAllText(jsonFilePath);
                dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
                filePath = jsonData.ExcelFilePath;
            }

            var existingNumbers = new HashSet<int>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null || worksheet.Dimension == null)
                {
                    txtMaTamGiam.Text = "MTG00001";
                    return;
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cellValue = worksheet.Cells[row, 2].Value?.ToString();
                    if (cellValue != null && cellValue.StartsWith("MTG"))
                    {
                        if (int.TryParse(cellValue.Substring(3), out int number))
                        {
                            existingNumbers.Add(number);
                        }
                    }
                }
            }

            // Tìm số thiếu đầu tiên
            int newMaTamGiam = 1;
            while (existingNumbers.Contains(newMaTamGiam))
            {
                newMaTamGiam++;
            }

            txtMaTamGiam.Text = "MTG" + newMaTamGiam.ToString("D5");
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime ngayBatDau = dtpNgayBatDauTamGiam.Value;
                DateTime ngayKetThuc = dtpNgayKetThucTamGiam.Value;

                if (ngayBatDau > ngayKetThuc)
                {
                    MessageBox.Show("[Ngày bắt đầu tạm giam] phải nhỏ hơn hoặc bằng [Ngày kết thúc tạm giam]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Collect data and close form
                detentionData = new DataTable();
                detentionData.Columns.Add("Ngày nhập");
                detentionData.Columns.Add("Mã tạm giam");
                detentionData.Columns.Add("Họ");
                detentionData.Columns.Add("Tên");
                detentionData.Columns.Add("CCCD");
                detentionData.Columns.Add("SĐT");
                detentionData.Columns.Add("Địa chỉ");
                detentionData.Columns.Add("Ngày bắt đầu");
                detentionData.Columns.Add("Ngày kết thúc");

                DataRow row = detentionData.NewRow();
                row["Ngày nhập"] = DateTime.Now.ToString("yyyy-MM-dd");
                row["Mã tạm giam"] = txtMaTamGiam.Text;
                row["Họ"] = txtHo.Text;
                row["Tên"] = txtTen.Text;
                row["CCCD"] = txtCCCD.Text;
                row["SĐT"] = txtSDT.Text;
                row["Địa chỉ"] = txtDiaChi.Text;
                row["Ngày bắt đầu"] = dtpNgayBatDauTamGiam.Value.ToString("yyyy-MM-dd");
                row["Ngày kết thúc"] = dtpNgayKetThucTamGiam.Value.ToString("yyyy-MM-dd");

                detentionData.Rows.Add(row);

                DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu thông tin: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

    }

}
