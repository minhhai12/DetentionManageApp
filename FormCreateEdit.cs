using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DetentionManageApp
{
    public partial class FormCreateEdit : Form
    {
        public DataTable detentionData { get; private set; }
        private readonly FormMode formMode;
        string jsonFilePath;

        bool isSoThuLyExist = false;

        public Func<DataTable, bool> TrySaveCallback { get; set; }

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
                txtStt.Text = selectedRow.Cells["STT"].Value.ToString();
                txtSoThuLy.Text = selectedRow.Cells["Số thụ lý"].Value.ToString();
                if (DateTime.TryParse(selectedRow.Cells["Ngày thụ lý"].Value?.ToString(), out DateTime ngayThuLy))
                    dtpNgayThuLy.Value = ngayThuLy;
                txtHoVaTen.Text = selectedRow.Cells["Họ và tên"].Value.ToString();
                txtNamSinh.Text = selectedRow.Cells["Năm sinh"].Value.ToString();
                cbGioiTinh.SelectedItem = selectedRow.Cells["Giới tính"].Value.ToString();
                txtToiDanh.Text = selectedRow.Cells["Tội danh"].Value.ToString();
                txtSoGiam.Text = selectedRow.Cells["Số giam"].Value.ToString();
                if (DateTime.TryParse(selectedRow.Cells["Ngày quyết định"].Value?.ToString(), out DateTime ngayQd))
                    dtpNgayQuyetDinh.Value = ngayQd;
                // Load Số ngày tạm giam
                if (selectedRow.DataGridView.Columns.Contains("Số ngày tạm giam") && selectedRow.Cells["Số ngày tạm giam"].Value != null)
                {
                    cbSoNgayTamGiam.SelectedItem = selectedRow.Cells["Số ngày tạm giam"].Value.ToString();
                }
                txtThoiHanTamGiam.Text = selectedRow.Cells["Thời hạn tạm giam"].Value.ToString();
                txtDiaChi.Text = selectedRow.Cells["Địa chỉ"].Value.ToString();
                if (DateTime.TryParse(selectedRow.Cells["Ngày bắt đầu"].Value?.ToString(), out DateTime ngayBatDau))
                {
                    // Nếu data cũ bị sai (Ngày bắt đầu < Ngày thụ lý), tự động sửa lại bằng mức tối thiểu
                    if (ngayBatDau < dtpNgayBatDau.MinDate)
                    {
                        dtpNgayBatDau.Value = dtpNgayBatDau.MinDate;
                    }
                    else
                    {
                        dtpNgayBatDau.Value = ngayBatDau;
                    }
                }
                if (DateTime.TryParse(selectedRow.Cells["Ngày hết hạn"].Value?.ToString(), out DateTime ngayHetHan))
                    dtpNgayHetHan.Value = ngayHetHan;
                // 20260617 Update: Load location from selected row if exists
                if (selectedRow.DataGridView.Columns.Contains("Địa điểm") && selectedRow.Cells["Địa điểm"].Value != null)
                {
                    string savedLocation = selectedRow.Cells["Địa điểm"].Value.ToString();
                    if (!string.IsNullOrEmpty(savedLocation))
                    {
                        cbLocation.SelectedValue = savedLocation;
                    }
                    else
                    {
                        cbLocation.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lấy dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadLocationComboBox()
        {
            List<LocationItem> locations = new List<LocationItem>();
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DetentionManage");
            string locFilePath = Path.Combine(folderPath, "data-location.json");

            if (File.Exists(locFilePath))
            {
                try
                {
                    string jsonContent = File.ReadAllText(locFilePath);
                    var items = JsonConvert.DeserializeObject<List<LocationItem>>(jsonContent);
                    if (items != null)
                    {
                        locations.AddRange(items);
                    }
                }
                catch (Exception) { /* Bỏ qua nếu lỗi đọc file */ }
            }

            // Chèn option rỗng lên đầu
            locations.Insert(0, new LocationItem { Id = "", Name = "-- Không chọn địa điểm --" });

            cbLocation.DataSource = locations;
            cbLocation.DisplayMember = "Name";
            cbLocation.ValueMember = "Name"; // Dùng Name làm Value để lưu chữ vào file Excel
        }

        /// <summary>
        /// Generate new STT
        /// </summary>
        private void GenerateSTT()
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
                        txtStt.Text = "1";
                        return;
                    }

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        var cellValue = worksheet.Cells[row, 1].Value?.ToString();
                        if (cellValue != null)
                        {
                            int.TryParse(cellValue, out int number);
                            existingNumbers.Add(number);
                        }
                    }
                }

                // Tìm số thiếu đầu tiên
                int newSTT = 1;
                while (existingNumbers.Contains(newSTT))
                {
                    newSTT++;
                }

                txtStt.Text = newSTT.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tạo mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Check SoThuLy exist or not in excel file
        /// </summary>
        /// <param name="soThuLy"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public bool CheckSoThuLyExist(string soThuLy)
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

                if(string.IsNullOrEmpty(filePath))
                {
                    return false;
                }

                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        return false;
                    }

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        if (worksheet.Cells[row, 2].Value?.ToString() == soThuLy)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra Số thụ lý: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return false;
        }

        /// <summary>
        /// Show Create or Edit form base on mode
        /// </summary>
        /// <param name="mode"></param>
        public FormCreateEdit(FormMode mode, string jsonFilePathFromFormList)
        {
            InitializeComponent();
            LoadLocationComboBox();
            formMode = mode;
            jsonFilePath = jsonFilePathFromFormList;

            if (formMode == FormMode.Create)
            {
                GenerateSTT();
            }

            btnSave.Text = formMode == FormMode.Create ? "Tạo mới" : "Cập nhật";
            lblTitleThongTin.Text = formMode == FormMode.Create ? "Tạo mới thông tin" : "Chỉnh sửa thông tin";
            btnCancel.Text = "Hủy bỏ";
            txtStt.Enabled = false;
            txtSoThuLy.Enabled = formMode == FormMode.Create ? true : false;

            // Thêm các giá trị cho ComboBox Giới Tính
            cbGioiTinh.Items.AddRange(new string[] { "Nam", "Nữ" });
            cbGioiTinh.SelectedIndex = 0; // Thiết lập giá trị mặc định

            // Thiết lập giá trị ComboBox Số ngày tạm giam
            cbSoNgayTamGiam.Items.AddRange(new string[] { "45", "60", "75", "105" });
            cbSoNgayTamGiam.SelectedIndex = 0;

            // Khóa 2 ô không cho nhập thủ công
            txtThoiHanTamGiam.Enabled = false;
            dtpNgayHetHan.Enabled = false;

            // Đặt ngày mặc định
            dtpNgayBatDau.Value = DateTime.Now;
            dtpNgayThuLy.Value = DateTime.Now;
            dtpNgayQuyetDinh.Value = DateTime.Now;

            // Đăng ký sự kiện
            dtpNgayThuLy.ValueChanged += (s, e) => {
                UpdateNgayBatDauConstraints(); // Cập nhật lại lịch Ngày bắt đầu trước
                CalculateDetentionDates();     // Sau đó mới tính ngày hết hạn
            };

            cbSoNgayTamGiam.SelectedIndexChanged += (s, e) => CalculateDetentionDates();
            dtpNgayBatDau.ValueChanged += (s, e) => CalculateDetentionDates();

            // Chạy lần đầu lúc khởi tạo form
            UpdateNgayBatDauConstraints();
            CalculateDetentionDates();

            // Add color for button
            btnSave.BackColor = formMode == FormMode.Create ? Color.DarkGreen : Color.DarkBlue;
            btnSave.ForeColor = Color.White;
            btnCancel.BackColor = Color.LightGray;
        }

        public FormCreateEdit(FormMode mode, string jsonFilePathFromFormList, DataGridViewRow selectedRow) : this(mode, jsonFilePathFromFormList)
        {
            LoadData(selectedRow);
        }

        /// <summary>
        /// Hàm tự động tính Ngày hết hạn và Thời hạn tạm giam
        /// </summary>
        private void CalculateDetentionDates()
        {
            if (cbSoNgayTamGiam.SelectedItem != null && int.TryParse(cbSoNgayTamGiam.SelectedItem.ToString(), out int soNgay))
            {
                // Ngày hết hạn = Ngày thụ lý + Số ngày tạm giam - 1
                DateTime ngayHetHan = dtpNgayThuLy.Value.AddDays(soNgay - 1);
                dtpNgayHetHan.Value = ngayHetHan;

                // Thời hạn tạm giam = Ngày hết hạn - Ngày bắt đầu
                int thoiHan = (ngayHetHan - dtpNgayBatDau.Value).Days + 1;

                // Nếu người dùng chọn Ngày bắt đầu trễ hơn cả ngày hết hạn, số ngày sẽ âm (logic chặn lưu ở btnSave)
                txtThoiHanTamGiam.Text = thoiHan.ToString();
            }
        }

        /// <summary>
        /// Cập nhật giới hạn chọn lịch của Ngày bắt đầu dựa trên Ngày thụ lý
        /// </summary>
        private void UpdateNgayBatDauConstraints()
        {
            try
            {
                DateTime ngayThuLyDate = dtpNgayThuLy.Value.Date;

                // BẮT BUỘC: Nếu Value hiện tại đang nhỏ hơn MinDate sắp thiết lập, 
                // DateTimePicker sẽ văng lỗi. Nên ta phải dời Value lên trước.
                if (dtpNgayBatDau.Value.Date < ngayThuLyDate)
                {
                    dtpNgayBatDau.Value = ngayThuLyDate;
                }

                // Khóa không cho chọn ngày trước Ngày thụ lý
                dtpNgayBatDau.MinDate = ngayThuLyDate;
            }
            catch (Exception) { /* Bỏ qua lỗi lặt vặt nếu có lúc load form */ }
        }

        /// <summary>
        /// Save button click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result != DialogResult.Yes) return;
            
            try
            {
                DateTime ngayBatDau = dtpNgayBatDau.Value;
                DateTime ngayHetHan = dtpNgayHetHan.Value;

                if (ngayBatDau > ngayHetHan)
                {
                    MessageBox.Show("[ Ngày bắt đầu tạm giam ] phải nhỏ hơn [ Ngày hết hạn tạm giam ]", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 20240722 Update
                // Số thụ lý có thể trùng nhau nên không cần kiểm tra
                //if(formMode == FormMode.Create)
                //{
                //    if (CheckSoThuLyExist(txtSoThuLy.Text))
                //    {
                //        MessageBox.Show("Số thụ lý đã tồn tại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //        txtSoThuLy.Focus();
                //        return;
                //    }
                //}

                // Collect data and close form
                detentionData = new DataTable();
                detentionData.Columns.Add("STT");
                detentionData.Columns.Add("Số thụ lý");
                detentionData.Columns.Add("Ngày thụ lý");
                detentionData.Columns.Add("Họ và tên");
                detentionData.Columns.Add("Năm sinh");
                detentionData.Columns.Add("Giới tính");
                detentionData.Columns.Add("Tội danh");
                detentionData.Columns.Add("Số giam");
                detentionData.Columns.Add("Ngày quyết định");
                detentionData.Columns.Add("Thời hạn tạm giam");
                detentionData.Columns.Add("Địa chỉ");
                detentionData.Columns.Add("Ngày bắt đầu");
                detentionData.Columns.Add("Ngày hết hạn");
                detentionData.Columns.Add("Địa điểm");
                detentionData.Columns.Add("Số ngày tạm giam");

                DataRow row = detentionData.NewRow();
                row["STT"] = txtStt.Text;
                row["Số thụ lý"] = txtSoThuLy.Text.Trim();
                row["Ngày thụ lý"] = dtpNgayThuLy.Value.ToString("dd/MM/yyyy");
                row["Họ và tên"] = txtHoVaTen.Text.Trim();
                row["Năm sinh"] = txtNamSinh.Text.Trim();
                row["Giới tính"] = cbGioiTinh.SelectedItem.ToString();
                row["Tội danh"] = txtToiDanh.Text.Trim();
                row["Số giam"] = txtSoGiam.Text.Trim();
                row["Ngày quyết định"] = dtpNgayQuyetDinh.Value.ToString("dd/MM/yyyy");
                row["Thời hạn tạm giam"] = txtThoiHanTamGiam.Text.Trim();
                row["Địa chỉ"] = txtDiaChi.Text.Trim();
                row["Ngày bắt đầu"] = dtpNgayBatDau.Value.ToString("dd/MM/yyyy");
                row["Ngày hết hạn"] = dtpNgayHetHan.Value.ToString("dd/MM/yyyy");
                string selectedLocation = cbLocation.SelectedValue?.ToString();
                if (selectedLocation == "-- Không chọn địa điểm --" || string.IsNullOrEmpty(selectedLocation))
                {
                    selectedLocation = "";
                }
                row["Địa điểm"] = selectedLocation;
                row["Số ngày tạm giam"] = cbSoNgayTamGiam.SelectedItem.ToString();

                detentionData.Rows.Add(row);

                // 20250826 Update
                if (TrySaveCallback != null)
                {
                    bool saved = false;
                    try
                    {
                        saved = TrySaveCallback.Invoke(detentionData);
                    }
                    catch (Exception ex)
                    {
                        // Nếu callback ném lỗi chưa bắt, báo lỗi và giữ form
                        MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (!saved)
                    {
                        // Lưu thất bại -> Giữ nguyên form để người dùng sửa
                        return;
                    }
                }

                // Lưu thành công -> đóng form
                DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu thông tin: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void txtHoVaTen_Leave(object sender, EventArgs e)
        {
            txtHoVaTen.Text = CapitalizeEachWord(txtHoVaTen.Text);
        }

        private void txtSoThuLy_Leave(object sender, EventArgs e)
        {
            // 20250912 Update
            // Số thụ lý có thể trùng nhau nên cần xác nhận trước khi tiếp tục
            if (formMode == FormMode.Create)
            {
                if (CheckSoThuLyExist(txtSoThuLy.Text))
                {
                    DialogResult rs1 = MessageBox.Show("Số thụ lý này đã có.\n Bạn vẫn muốn tiếp tục?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (rs1 != DialogResult.Yes)
                    {
                        txtSoThuLy.Focus();
                    }
                }
            }
        }

    }

}
