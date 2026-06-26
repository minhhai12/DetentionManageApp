using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DetentionManageApp
{
    public partial class FormLocations : Form
    {
        private string jsonFilePath;
        private BindingList<LocationItem> locationList; // Danh sách gốc lưu toàn bộ data

        private string placeholderText = "Nhập từ khóa tìm kiếm...";

        public FormLocations()
        {
            InitializeComponent();
            SetupForm();
            LoadData();
        }

        private void SetupForm()
        {
            this.Text = "Danh sách địa điểm";
            this.StartPosition = FormStartPosition.CenterParent;

            // --- CẤU HÌNH DATAGRIDVIEW ---
            dgvLocations.AutoGenerateColumns = false;
            dgvLocations.AllowUserToAddRows = true;
            dgvLocations.AllowUserToDeleteRows = false;
            dgvLocations.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dgvLocations.MultiSelect = false;
            dgvLocations.RowHeadersVisible = false;

            dgvLocations.EditMode = DataGridViewEditMode.EditProgrammatically;

            CreateGridColumns();

            // Gắn các sự kiện của lưới
            dgvLocations.CellDoubleClick += DgvLocations_CellDoubleClick;
            dgvLocations.CellValidating += DgvLocations_CellValidating;
            dgvLocations.CellEndEdit += DgvLocations_CellEndEdit;
            dgvLocations.CellContentClick += DgvLocations_CellContentClick;
            dgvLocations.DataError += DgvLocations_DataError;

            // Gắn sự kiện cho ô tìm kiếm
            if (txtSearch != null)
                txtSearch.TextChanged += TxtSearch_TextChanged;
            if (btnClearSearch != null)
                btnClearSearch.Click += BtnClearSearch_Click;
        }

        private void CreateGridColumns()
        {
            dgvLocations.Columns.Clear();

            dgvLocations.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Id",
                Name = "Id",
                Visible = false
            });

            dgvLocations.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = "Name",
                Name = "Name",
                HeaderText = "Tên địa điểm (Double-click để Thêm/Sửa)",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            });

            dgvLocations.Columns.Add(new DataGridViewButtonColumn
            {
                Name = "btnDelete",
                HeaderText = "",
                Text = "",
                UseColumnTextForButtonValue = true,
                Width = 40,
                FlatStyle = FlatStyle.Flat
            });
        }

        private string GetJsonFilePath()
        {
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DetentionManage");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            return Path.Combine(folderPath, "data-location.json");
        }

        private void LoadData()
        {
            jsonFilePath = GetJsonFilePath();
            if (File.Exists(jsonFilePath))
            {
                try
                {
                    string jsonContent = File.ReadAllText(jsonFilePath);
                    var items = JsonConvert.DeserializeObject<List<LocationItem>>(jsonContent);
                    locationList = items != null ? new BindingList<LocationItem>(items) : new BindingList<LocationItem>();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi đọc file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    locationList = new BindingList<LocationItem>();
                }
            }
            else
            {
                locationList = new BindingList<LocationItem>();
            }

            dgvLocations.DataSource = locationList;
        }

        private void SaveDataToJson()
        {
            try
            {
                // LUÔN LƯU TỪ DANH SÁCH GỐC (locationList)
                var validItems = locationList.Where(l => !string.IsNullOrWhiteSpace(l.Name)).ToList();
                string jsonContent = JsonConvert.SerializeObject(validItems, Formatting.Indented);
                File.WriteAllText(jsonFilePath, jsonContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lưu file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- XỬ LÝ TÌM KIẾM ---

        private void FilterData(string searchText)
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Nếu ô tìm kiếm rỗng, hiển thị lại toàn bộ danh sách gốc
                dgvLocations.DataSource = locationList;
            }
            else
            {
                // Nếu có chữ, tạo một danh sách tạm thời chứa các kết quả khớp
                var filtered = locationList.Where(l => l.Name != null && l.Name.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                dgvLocations.DataSource = new BindingList<LocationItem>(filtered);
            }
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text == placeholderText)
            {
                return; // Dừng lại không thực hiện tìm kiếm
            }
            FilterData(txtSearch.Text);
        }

        private void BtnClearSearch_Click(object sender, EventArgs e)
        {
            if (txtSearch != null)
            {
                txtSearch.Text = string.Empty;
            }
            txtSearch.Focus();
        }

        // --- XỬ LÝ CÁC SỰ KIỆN CỦA LƯỚI ---

        private void DgvLocations_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && dgvLocations.Columns[e.ColumnIndex].Name == "Name")
            {
                dgvLocations.BeginEdit(true);
            }
        }

        private void DgvLocations_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dgvLocations.Columns[e.ColumnIndex].Name == "Name")
            {
                string newValue = e.FormattedValue.ToString().Trim();
                if (string.IsNullOrEmpty(newValue)) return;

                var row = dgvLocations.Rows[e.RowIndex];
                var item = row.DataBoundItem as LocationItem;
                string currentId = item?.Id;

                // Cần kiểm tra trùng lặp trên danh sách gốc (locationList)
                bool isDuplicate = locationList.Any(l =>
                    l.Name != null &&
                    l.Name.Equals(newValue, StringComparison.OrdinalIgnoreCase) &&
                    (currentId == null || l.Id != currentId));

                if (isDuplicate)
                {
                    MessageBox.Show("Địa điểm này đã tồn tại! Vui lòng nhập tên khác.", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
            }
        }

        private void DgvLocations_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var row = dgvLocations.Rows[e.RowIndex];
            var item = row.DataBoundItem as LocationItem;

            if (item != null && !string.IsNullOrWhiteSpace(item.Name))
            {
                item.Name = item.Name.Trim();

                if (string.IsNullOrEmpty(item.Id))
                {
                    item.Id = Guid.NewGuid().ToString();

                    // NẾU NGƯỜI DÙNG TẠO MỚI TRONG LÚC ĐANG SEARCH
                    // Thì item này chỉ mới được thêm vào danh sách lọc (tạm), ta phải đẩy nó vào danh sách gốc
                    if (!locationList.Contains(item))
                    {
                        locationList.Add(item);
                    }
                }

                SaveDataToJson();
            }
        }

        private void DgvLocations_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            if (dgvLocations.Columns[e.ColumnIndex].Name == "btnDelete")
            {
                DataGridViewRow row = dgvLocations.Rows[e.RowIndex];
                if (row.IsNewRow) return;

                var item = row.DataBoundItem as LocationItem;
                if (item != null)
                {
                    DialogResult rs = MessageBox.Show($"Bạn có chắc chắn muốn xóa địa điểm: {item.Name}?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (rs == DialogResult.Yes)
                    {
                        // Xóa ở danh sách hiển thị hiện tại (trường hợp đang dùng Filter)
                        var currentDataSource = dgvLocations.DataSource as BindingList<LocationItem>;
                        if (currentDataSource != null && currentDataSource != locationList)
                        {
                            currentDataSource.Remove(item);
                        }

                        // Xóa ở danh sách gốc
                        locationList.Remove(item);

                        SaveDataToJson();
                    }
                }
            }
        }

        private void DgvLocations_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
        }

        private void FormLocations_Resize(object sender, EventArgs e)
        {
            lblTitleDiaDiem.Location = new Point((this.ClientSize.Width - lblTitleDiaDiem.Width) / 2, 30);
        }

        private void txtSearch_Enter(object sender, EventArgs e)
        {
            // Nếu chữ trong ô đang là chữ gợi ý, thì xóa đi để người dùng nhập
            if (txtSearch.Text == placeholderText)
            {
                txtSearch.Text = "";
                txtSearch.ForeColor = Color.Black; // Đổi màu chữ về đen bình thường
            }
        }

        private void txtSearch_Leave(object sender, EventArgs e)
        {
            // Nếu người dùng không nhập gì cả mà click ra ngoài, thì hiện lại chữ gợi ý
            if (string.IsNullOrWhiteSpace(txtSearch.Text))
            {
                txtSearch.Text = placeholderText;
                txtSearch.ForeColor = Color.Gray; // Đổi màu chữ thành xám mờ
            }
        }

        private void dgvLocations_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            // Đảm bảo đang vẽ ở cột btnDelete và không phải dòng tiêu đề
            if (e.RowIndex >= 0 && dgvLocations.Columns[e.ColumnIndex].Name == "btnDelete")
            {
                // KHÔNG vẽ icon ở dòng trống (dùng để thêm mới) dưới cùng
                if (dgvLocations.Rows[e.RowIndex].IsNewRow) return;

                // Vẽ nền của nút (giữ lại hiệu ứng khi di chuột / click) nhưng bỏ qua phần chữ mặc định
                e.Paint(e.CellBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);

                Image icon = Properties.Resources.icon_trash_bin;

                if (icon != null)
                {
                    // Thiết lập kích thước icon hiển thị (ví dụ 20x20 pixel)
                    int iconWidth = 24;
                    int iconHeight = 24;

                    // Tính toán tọa độ X, Y để vẽ icon nằm ngay chính giữa cái nút
                    int x = e.CellBounds.Left + (e.CellBounds.Width - iconWidth) / 2;
                    int y = e.CellBounds.Top + (e.CellBounds.Height - iconHeight) / 2;

                    // Tiến hành vẽ icon
                    e.Graphics.DrawImage(icon, new Rectangle(x, y, iconWidth, iconHeight));
                }

                // Báo cho WinForms biết bạn đã tự vẽ xong, hệ thống không cần can thiệp nữa
                e.Handled = true;
            }
        }
    }

    public class LocationItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; }
    }
}