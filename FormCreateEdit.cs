using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
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

        /// <summary>
        /// Load data after select on edit mode
        /// </summary>
        /// <param name="selectedRow"></param>
        private void LoadData(DataGridViewRow selectedRow)
        {
            try
            {
                txtMaTamGiam.Text = selectedRow.Cells["Mã tạm giam"].Value.ToString();
                txtHo.Text = selectedRow.Cells["Họ"].Value.ToString();
                txtTen.Text = selectedRow.Cells["Tên"].Value.ToString();
                txtCCCD.Text = selectedRow.Cells["CCCD"].Value.ToString();
                dtpNgaySinh.Value = DateTime.Parse(selectedRow.Cells["Ngày sinh"].Value.ToString());
                cbGioiTinh.SelectedItem = selectedRow.Cells["Giới tính"].Value.ToString();
                txtSDT.Text = selectedRow.Cells["SĐT"].Value.ToString();
                txtDiaChi.Text = selectedRow.Cells["Địa chỉ"].Value.ToString();
                dtpNgayBatDauTamGiam.Value = DateTime.Parse(selectedRow.Cells["Ngày bắt đầu"].Value.ToString());
                dtpNgayKetThucTamGiam.Value = DateTime.Parse(selectedRow.Cells["Ngày kết thúc"].Value.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lấy dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Generate new MaTamGiam
        /// </summary>
        private void GenerateMaTamGiam(string jsonFilePath)
        {
            try
            {
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
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tạo Mã tạm giam: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Show Create or Edit form base on mode
        /// </summary>
        /// <param name="mode"></param>
        public FormCreateEdit(FormMode mode, string jsonFilePath)
        {
            InitializeComponent();
            formMode = mode;
            if (formMode == FormMode.Create)
            {
                GenerateMaTamGiam(jsonFilePath);
            }
            btnSave.Text = formMode == FormMode.Create ? "Tạo mới" : "Cập nhật";
            lblTitleThongTin.Text = formMode == FormMode.Create ? "Tạo mới thông tin" : "Chỉnh sửa thông tin";
            btnCancel.Text = "Hủy bỏ";
            txtMaTamGiam.Enabled = false;

            // Thêm các giá trị cho ComboBox Giới Tính
            cbGioiTinh.Items.AddRange(new string[] { "Nam", "Nữ", "Khác" });
            cbGioiTinh.SelectedIndex = 0; // Thiết lập giá trị mặc định

            // Đặt giá trị mặc định cho dtpNgaySinh
            dtpNgaySinh.Value = new DateTime(1990, 1, 1);
        }

        public FormCreateEdit(FormMode mode, string jsonFilePath, DataGridViewRow selectedRow) : this(mode, jsonFilePath)
        {
            LoadData(selectedRow);
        }

        /// <summary>
        /// Save button click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {
                    DateTime ngayBatDau = dtpNgayBatDauTamGiam.Value;
                    DateTime ngayKetThuc = dtpNgayKetThucTamGiam.Value;

                    if (ngayBatDau > ngayKetThuc)
                    {
                        MessageBox.Show("[ Ngày bắt đầu tạm giam ] phải nhỏ hơn [ Ngày kết thúc tạm giam ]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Collect data and close form
                    detentionData = new DataTable();
                    detentionData.Columns.Add("Ngày nhập");
                    detentionData.Columns.Add("Mã tạm giam");
                    detentionData.Columns.Add("Họ");
                    detentionData.Columns.Add("Tên");
                    detentionData.Columns.Add("CCCD");
                    detentionData.Columns.Add("Ngày sinh");
                    detentionData.Columns.Add("Giới tính");
                    detentionData.Columns.Add("SĐT");
                    detentionData.Columns.Add("Địa chỉ");
                    detentionData.Columns.Add("Ngày bắt đầu");
                    detentionData.Columns.Add("Ngày kết thúc");

                    DataRow row = detentionData.NewRow();
                    row["Ngày nhập"] = DateTime.Now.ToString("dd/MM/yyyy");
                    row["Mã tạm giam"] = txtMaTamGiam.Text.Trim();
                    row["Họ"] = txtHo.Text.Trim();
                    row["Tên"] = txtTen.Text.Trim();
                    row["CCCD"] = txtCCCD.Text.Trim();
                    row["Ngày sinh"] = dtpNgaySinh.Value.ToString("dd/MM/yyyy");
                    row["Giới tính"] = cbGioiTinh.SelectedItem.ToString();
                    row["SĐT"] = txtSDT.Text.Trim();
                    row["Địa chỉ"] = txtDiaChi.Text.Trim();
                    row["Ngày bắt đầu"] = dtpNgayBatDauTamGiam.Value.ToString("dd/MM/yyyy");
                    row["Ngày kết thúc"] = dtpNgayKetThucTamGiam.Value.ToString("dd/MM/yyyy");

                    detentionData.Rows.Add(row);

                    DialogResult = DialogResult.OK;
                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi lưu thông tin: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Cancel button click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        /// <summary>
        /// Capitalize Each Word
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        private string CapitalizeEachWord(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text.ToLower());
        }

        private void txtHo_Leave(object sender, EventArgs e)
        {
            txtHo.Text = CapitalizeEachWord(txtHo.Text);
        }

        private void txtTen_Leave(object sender, EventArgs e)
        {
            txtTen.Text = CapitalizeEachWord(txtTen.Text);
        }

    }

}
