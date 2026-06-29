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

        // 20260619 Update: Tạo biến cho Form Lần 2
        private bool isInitializing = true; // Cờ chặn sự kiện lúc form đang load

        private readonly string noiDungLan1 = "Xét thấy cần thiết tiếp tục tạm giam bị can để bảo đảm cho việc giải quyết vụ án,";
        private readonly string noiDungLan2 = "Xét thấy cần thiết tiếp tục tạm giam bị can để bảo đảm hoàn thành việc xét xử sơ thẩm,";

        private string thoiHanTamGiamLan1 = "";
        private string thoiHanTamGiamLan2 = "";

        /// <summary>
        /// Load data after select on edit mode
        /// </summary>
        /// <param name="selectedRow"></param>
        private void LoadData(DataGridViewRow selectedRow)
        {
            try
            {
                // 1. BẬT CỜ CHẶN: Khóa toàn bộ các sự kiện và logic tính toán trung gian khi đang nạp dữ liệu thô
                isInitializing = true;

                // Tạm thời hạ mức MinDate xuống thấp nhất để tránh lỗi văng ứng dụng khi gõ .Value trước
                dtpNgayBatDau.MinDate = new DateTime(1900, 1, 1);

                txtStt.Text = selectedRow.Cells["STT"].Value.ToString();
                txtSoThuLy.Text = selectedRow.Cells["Số thụ lý"].Value.ToString();

                if (DateTime.TryParseExact(selectedRow.Cells["Ngày thụ lý"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ngayThuLy))
                    dtpNgayThuLy.Value = ngayThuLy;

                txtHoVaTen.Text = selectedRow.Cells["Họ và tên"].Value.ToString();
                txtNamSinh.Text = selectedRow.Cells["Năm sinh"].Value.ToString();
                cbGioiTinh.SelectedItem = selectedRow.Cells["Giới tính"].Value.ToString();

                if (selectedRow.DataGridView.Columns.Contains("Nghề nghiệp") && selectedRow.Cells["Nghề nghiệp"].Value != null)
                {
                    txtNgheNghiep.Text = selectedRow.Cells["Nghề nghiệp"].Value.ToString();
                }
                if (selectedRow.DataGridView.Columns.Contains("Điều khoản") && selectedRow.Cells["Điều khoản"].Value != null)
                {
                    txtDieuKhoan.Text = selectedRow.Cells["Điều khoản"].Value.ToString();
                }
                txtToiDanh.Text = selectedRow.Cells["Tội danh"].Value.ToString();
                txtSoGiam.Text = selectedRow.Cells["Số giam"].Value.ToString();

                if (DateTime.TryParseExact(selectedRow.Cells["Ngày quyết định"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ngayQd))
                    dtpNgayQuyetDinh.Value = ngayQd;

                if (selectedRow.DataGridView.Columns.Contains("Số ngày tạm giam") && selectedRow.Cells["Số ngày tạm giam"].Value != null)
                {
                    cbSoNgayTamGiam.SelectedItem = selectedRow.Cells["Số ngày tạm giam"].Value.ToString();
                }

                txtDiaChi.Text = selectedRow.Cells["Địa chỉ"].Value.ToString();

                // Nạp Ngày bắt đầu và Ngày hết hạn Lần 1 từ dữ liệu gốc
                if (DateTime.TryParseExact(selectedRow.Cells["Ngày bắt đầu"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ngayBatDau))
                    dtpNgayBatDau.Value = ngayBatDau.Date;

                if (DateTime.TryParseExact(selectedRow.Cells["Ngày hết hạn"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ngayHetHan))
                    dtpNgayHetHan.Value = ngayHetHan;

                if (selectedRow.DataGridView.Columns.Contains("Địa điểm") && selectedRow.Cells["Địa điểm"].Value != null)
                {
                    string savedLocation = selectedRow.Cells["Địa điểm"].Value.ToString();
                    if (!string.IsNullOrEmpty(savedLocation))
                        cbLocation.SelectedValue = savedLocation;
                    else
                        cbLocation.SelectedIndex = 0;
                }

                // Load Số lần và dữ liệu Lần 2
                if (selectedRow.DataGridView.Columns.Contains("Số lần") && selectedRow.Cells["Số lần"].Value != null)
                {
                    string soLanSaved = selectedRow.Cells["Số lần"].Value.ToString();
                    if (!string.IsNullOrEmpty(soLanSaved))
                    {
                        cbSoLan.SelectedItem = soLanSaved;
                        bool isLan2 = (soLanSaved == "Lần 2");

                        if (isLan2)
                        {
                            cbSoLan.Enabled = false; // Khóa nếu dữ liệu gốc đã lưu là Lần 2
                        }

                        lblNgayBatDauLan2.Visible = isLan2;
                        dtpNgayBatDauLan2.Visible = isLan2;
                        lblNgayHetHanLan2.Visible = isLan2;
                        dtpNgayHetHanLan2.Visible = isLan2;
                        lblGiaHan.Visible = isLan2;
                        txtGiaHan.Visible = isLan2;
                        lblTiTleNgay3.Visible = isLan2;

                        if (isLan2)
                        {
                            txtGiaHan.Text = selectedRow.Cells["Gia hạn"].Value?.ToString() ?? "";

                            if (DateTime.TryParseExact(selectedRow.Cells["Ngày bắt đầu lần 2"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime nb2))
                                dtpNgayBatDauLan2.Value = nb2.Date;

                            if (DateTime.TryParseExact(selectedRow.Cells["Ngày hết hạn lần 2"].Value?.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime nh2))
                                dtpNgayHetHanLan2.Value = nh2.Date;
                        }
                    }
                }

                if (selectedRow.DataGridView.Columns.Contains("Nội dung") && selectedRow.Cells["Nội dung"].Value != null)
                {
                    txtNoiDung.Text = selectedRow.Cells["Nội dung"].Value.ToString();
                }

                if (selectedRow.DataGridView.Columns.Contains("Số lệnh trích xuất") && selectedRow.Cells["Số lệnh trích xuất"].Value != null)
                {
                    txtSoLenhTrichXuat.Text = selectedRow.Cells["Số lệnh trích xuất"].Value.ToString();
                }

                if (selectedRow.DataGridView.Columns.Contains("Thời gian") && selectedRow.Cells["Thời gian"].Value != null)
                {
                    string thoiGianStr = selectedRow.Cells["Thời gian"].Value.ToString();
                    if (DateTime.TryParseExact(thoiGianStr, "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime tgTrichXuat))
                    {
                        dtpThoiGianTrichXuat.Value = tgTrichXuat;
                    }
                }

                // 2. TẮT CỜ CHẶN VÀ THIẾT LẬP RÀO CHẮN CHUẨN: Lúc này toàn bộ dữ liệu thô sạch đã lên Form đầy đủ
                isInitializing = false;

                // Thiết lập lại MinDate chặn lịch chuẩn theo Ngày thụ lý mới nạp từ Excel để bảo vệ logic nhập liệu về sau
                dtpNgayBatDau.MinDate = dtpNgayThuLy.Value.Date;

                // 3. ÉP TÍNH TOÁN LẠI: Đồng bộ các biến tạm thời hạn chính xác theo hệ quy chiếu dữ liệu vừa load
                CalculateDetentionDates();
            }
            catch (Exception ex)
            {
                isInitializing = false;
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

                if (string.IsNullOrEmpty(filePath))
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

            pnLenhTamGiam.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            pnLenhTrichXuat.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            pnButton.Anchor = AnchorStyles.Top | AnchorStyles.Left;

            // Giới hạn chiều cao Form không vượt quá màn hình hiện tại
            int screenHeight = Screen.PrimaryScreen.WorkingArea.Height; // Lấy chiều cao vùng làm việc (đã trừ Taskbar)
            if (this.Height > screenHeight - 20) // Trừ hao thêm 20px cho viền
            {
                this.Height = screenHeight - 20;
            }

            // Đảm bảo Form luôn căn giữa màn hình
            this.StartPosition = FormStartPosition.CenterScreen;

            LoadLocationComboBox();
            formMode = mode;
            jsonFilePath = jsonFilePathFromFormList;

            if (formMode == FormMode.Create)
            {
                GenerateSTT();
            }

            // Khởi tạo các giá trị cho ComboBox Số lần tạm giam
            cbSoLan.Items.AddRange(new string[] { "Lần 1", "Lần 2" });
            cbSoLan.SelectedIndex = 0;
            cbSoLan.Enabled = (formMode == FormMode.Edit); // Chỉ cho phép sửa nếu ở chế độ Edit
            cbSoLan.SelectedIndexChanged += CbSoLan_SelectedIndexChanged;

            btnSave.Text = formMode == FormMode.Create ? "Tạo mới" : "Cập nhật";
            lblTitleThongTin.Text = formMode == FormMode.Create ? "Tạo mới thông tin" : "Chỉnh sửa thông tin";
            btnCancel.Text = "Hủy bỏ";
            txtStt.Enabled = false;
            txtSoThuLy.Enabled = formMode == FormMode.Create;

            // Thêm các giá trị cho ComboBox Giới Tính
            cbGioiTinh.Items.AddRange(new string[] { "Nam", "Nữ" });
            cbGioiTinh.SelectedIndex = 0;

            // Thiết lập giá trị ComboBox Số ngày tạm giam
            cbSoNgayTamGiam.Items.AddRange(new string[] { "45", "60", "75", "105" });
            cbSoNgayTamGiam.SelectedIndex = 0;

            // Khóa các ô 
            txtThoiHanTamGiam.Enabled = false;
            txtGiaHan.Enabled = false;

            dtpNgayHetHan.Enabled = false;
            dtpNgayBatDauLan2.Enabled = false;
            dtpNgayHetHanLan2.Enabled = false;

            // Đặt ngày mặc định
            dtpNgayBatDau.Value = DateTime.Now.Date;
            dtpNgayThuLy.Value = DateTime.Now.Date;
            dtpNgayQuyetDinh.Value = DateTime.Now.Date;

            // Đăng ký sự kiện (Có kèm cờ kiểm tra isInitializing để ko chạy loạn xạ khi load)
            dtpNgayThuLy.ValueChanged += (s, e) => {
                if (isInitializing) return;
                UpdateNgayBatDauConstraints();
                CalculateDetentionDates();
            };

            cbSoNgayTamGiam.SelectedIndexChanged += (s, e) => {
                if (isInitializing) return;
                CalculateDetentionDates();
            };

            dtpNgayBatDau.ValueChanged += (s, e) => {
                if (isInitializing) return;
                CalculateDetentionDates();
            };

            // Chạy lần đầu lúc khởi tạo form mới
            UpdateNgayBatDauConstraints();
            CalculateDetentionDates();

            if (formMode == FormMode.Create)
            {
                bool isLan2 = cbSoLan.SelectedItem.ToString() == "Lần 2";

                // Ẩn các control Lần 2 khi tạo mới
                lblNgayBatDauLan2.Visible = isLan2;
                dtpNgayBatDauLan2.Visible = isLan2;
                lblNgayHetHanLan2.Visible = isLan2;
                dtpNgayHetHanLan2.Visible = isLan2;
                lblGiaHan.Visible = isLan2;
                txtGiaHan.Visible = isLan2;
                lblTiTleNgay3.Visible = isLan2;

                txtNoiDung.Text = noiDungLan1;
            }

            txtSoLenhTrichXuat.TextChanged += txtSoLenhTrichXuat_TextChanged;
            txtSoLenhTrichXuat_TextChanged(null, EventArgs.Empty);

            btnSave.BackColor = formMode == FormMode.Create ? Color.DarkGreen : Color.DarkBlue;
            btnSave.ForeColor = Color.White;
            btnCancel.BackColor = Color.LightGray;

            isInitializing = false; // Mở khóa cho phép tính toán tự động
        }

        public FormCreateEdit(FormMode mode, string jsonFilePathFromFormList, DataGridViewRow selectedRow) : this(mode, jsonFilePathFromFormList)
        {
            LoadData(selectedRow);
        }

        /// <summary>
        /// Hàm tự động tính Ngày hết hạn và Thời hạn tạm giam chuẩn theo từng lần
        /// </summary>
        private void CalculateDetentionDates()
        {
            // Triệt tiêu hoàn toàn giờ/phút/giây ngầm bằng .Date
            DateTime ngayThuLyDate = dtpNgayThuLy.Value.Date;
            DateTime ngayBatDauDate = dtpNgayBatDau.Value.Date;

            if (cbSoNgayTamGiam.SelectedItem != null && int.TryParse(cbSoNgayTamGiam.SelectedItem.ToString(), out int soNgay))
            {
                // ====================================================
                // GIAI ĐOẠN LẦN 1
                // ====================================================
                // Ngày hết hạn Lần 1 = Ngày thụ lý + Số ngày tạm giam - 1
                DateTime ngayHetHanL1 = ngayThuLyDate.AddDays(soNgay - 1);
                dtpNgayHetHan.Value = ngayHetHanL1;

                // Thời hạn Lần 1 = Ngày hết hạn L1 - Ngày bắt đầu L1 + 1
                int thoiHan1 = (ngayHetHanL1 - ngayBatDauDate).Days + 1;
                thoiHanTamGiamLan1 = thoiHan1.ToString();


                // ====================================================
                // GIAI ĐOẠN LẦN 2
                // ====================================================
                // Ngày bắt đầu Lần 2 = Ngày hết hạn Lần 1 + 1 ngày
                DateTime ngayBatDauL2 = ngayHetHanL1.AddDays(1);
                dtpNgayBatDauLan2.Value = ngayBatDauL2;

                // Tính số ngày cộng thêm của riêng Lần 2 (15 hoặc 30 ngày)
                int ngayCongThem = (soNgay == 45 || soNgay == 60) ? 15 : 30;

                // Ngày hết hạn Lần 2 = Ngày bắt đầu Lần 2 + Số ngày cộng thêm - 1
                DateTime ngayHetHanL2 = ngayBatDauL2.AddDays(ngayCongThem - 1);
                dtpNgayHetHanLan2.Value = ngayHetHanL2;

                // CHỈNH SỬA: Tính thời hạn Lần 2 trực tiếp dựa trên 2 ngày của Lần 2
                int thoiHan2 = (ngayHetHanL2 - ngayBatDauL2).Days + 1;
                thoiHanTamGiamLan2 = thoiHan2.ToString();

                // Ô Gia hạn hiển thị tổng số ngày gia hạn (Số ngày gốc Lần 1 + số ngày cộng thêm Lần 2)
                int tongGiaHan = soNgay + ngayCongThem;
                txtGiaHan.Text = tongGiaHan.ToString();


                // ====================================================
                // CẬP NHẬT HIỂN THỊ LÊN GIAO DIỆN TEXTBOX
                // ====================================================
                if (cbSoLan.SelectedItem != null && cbSoLan.SelectedItem.ToString() == "Lần 2")
                {
                    // Nếu chọn Lần 2: Ô thời hạn hiển thị số ngày của Lần 2 (15 hoặc 30)
                    txtThoiHanTamGiam.Text = thoiHanTamGiamLan2;
                }
                else
                {
                    // Nếu chọn Lần 1: Ô thời hạn hiển thị số ngày của Lần 1
                    txtThoiHanTamGiam.Text = thoiHanTamGiamLan1;
                }
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
            catch (Exception) { }
        }

        private void CbSoLan_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isInitializing) return;

            bool isLan2 = cbSoLan.SelectedItem.ToString() == "Lần 2";

            lblNgayBatDauLan2.Visible = isLan2;
            dtpNgayBatDauLan2.Visible = isLan2;
            lblNgayHetHanLan2.Visible = isLan2;
            dtpNgayHetHanLan2.Visible = isLan2;
            lblGiaHan.Visible = isLan2;
            txtGiaHan.Visible = isLan2;
            lblTiTleNgay3.Visible = isLan2;

            if (isLan2)
            {
                txtNoiDung.Text = noiDungLan2;
            }
            else
            {
                txtNoiDung.Text = noiDungLan1;
            }

            CalculateDetentionDates(); // Nạp lại biến tạm thời hạn chính xác khi đảo chế độ
        }

        private void txtSoLenhTrichXuat_TextChanged(object sender, EventArgs e)
        {
            bool coDuLieu = !string.IsNullOrWhiteSpace(txtSoLenhTrichXuat.Text);
            lblThoiGianTrichXuat.Visible = coDuLieu;
            dtpThoiGianTrichXuat.Visible = coDuLieu;
        }

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
                detentionData.Columns.Add("Nghề nghiệp");
                detentionData.Columns.Add("Tội danh");
                detentionData.Columns.Add("Điều khoản");
                detentionData.Columns.Add("Số giam");
                detentionData.Columns.Add("Ngày quyết định");
                detentionData.Columns.Add("Thời hạn tạm giam");
                detentionData.Columns.Add("Địa chỉ");
                detentionData.Columns.Add("Ngày bắt đầu");
                detentionData.Columns.Add("Ngày hết hạn");
                detentionData.Columns.Add("Địa điểm");
                detentionData.Columns.Add("Số ngày tạm giam");
                detentionData.Columns.Add("Số lần");
                detentionData.Columns.Add("Gia hạn");
                detentionData.Columns.Add("Ngày bắt đầu lần 2");
                detentionData.Columns.Add("Ngày hết hạn lần 2");
                detentionData.Columns.Add("Nội dung");
                detentionData.Columns.Add("Số lệnh trích xuất");
                detentionData.Columns.Add("Thời gian");

                DataRow row = detentionData.NewRow();
                row["STT"] = txtStt.Text;
                row["Số thụ lý"] = txtSoThuLy.Text.Trim();
                row["Ngày thụ lý"] = dtpNgayThuLy.Value.ToString("dd/MM/yyyy");
                row["Họ và tên"] = txtHoVaTen.Text.Trim();
                row["Năm sinh"] = txtNamSinh.Text.Trim();
                row["Giới tính"] = cbGioiTinh.SelectedItem.ToString();
                row["Nghề nghiệp"] = txtNgheNghiep.Text.Trim();
                row["Tội danh"] = txtToiDanh.Text.Trim();
                row["Điều khoản"] = txtDieuKhoan.Text.Trim();
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

                bool isLan2 = cbSoLan.SelectedItem.ToString() == "Lần 2";

                row["Số lần"] = cbSoLan.SelectedItem.ToString();
                row["Gia hạn"] = isLan2 ? txtGiaHan.Text : "";
                row["Ngày bắt đầu lần 2"] = isLan2 ? dtpNgayBatDauLan2.Value.ToString("dd/MM/yyyy") : "";
                row["Ngày hết hạn lần 2"] = isLan2 ? dtpNgayHetHanLan2.Value.ToString("dd/MM/yyyy") : "";

                row["Nội dung"] = txtNoiDung.Text.Trim();
                row["Số lệnh trích xuất"] = txtSoLenhTrichXuat.Text.Trim();
                if (string.IsNullOrWhiteSpace(txtSoLenhTrichXuat.Text))
                {
                    row["Thời gian"] = "";
                }
                else
                {
                    row["Thời gian"] = dtpThoiGianTrichXuat.Value.ToString("dd/MM/yyyy HH:mm");
                }

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
                        MessageBox.Show("Lỗi khi lưu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (!saved) return;
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