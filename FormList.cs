using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DetentionManageApp
{
    public partial class FormList : Form
    {
        private string excelFilePath;
        private DataTable dataTable;
        private BindingSource bindingSource = new BindingSource();

        SortOrderEnum sortOrderCustomDate_NgayBatDau = SortOrderEnum.None;
        SortOrderEnum sortOrderCustomDate_NgayHetHan = SortOrderEnum.None;

        public enum SortOrderEnum
        {
            None,
            Ascending,
            Descending
        }

        /// <summary>
        /// Initialization function
        /// </summary>
        public FormList()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            btnEdit.Visible = false;
            btnDelete.Visible = false;
            LoadExcelFilePath();
            LoadDataFromExcel();
            btnChooseFile.BackColor = Color.LightGray;
            btnCreate.BackColor = Color.DarkGreen;
            btnEdit.BackColor = Color.DarkBlue;
            btnDelete.BackColor = Color.DarkRed;
            btnCreate.ForeColor = Color.White;
            btnEdit.ForeColor = Color.White;
            btnDelete.ForeColor = Color.White;
        }

        string GetJsonFilePath()
        {
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DetentionManage");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            return Path.Combine(folderPath, "excelFilePath.json");
        }

        /// <summary>
        /// Load Excel file path from json file
        /// </summary>
        private void LoadExcelFilePath()
        {
            try
            {
                string jsonFilePath = GetJsonFilePath();
                if (File.Exists(jsonFilePath))
                {
                    string jsonContent = File.ReadAllText(jsonFilePath);
                    dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
                    excelFilePath = jsonData.ExcelFilePath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi lấy Excel file path: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Save excel file path to json file
        /// </summary>
        /// <exception cref="Exception"></exception>
        private void SaveExcelFilePath()
        {
            try
            {
                string jsonFilePath = GetJsonFilePath();
                dynamic jsonData = new { ExcelFilePath = excelFilePath };
                string jsonContent = JsonConvert.SerializeObject(jsonData);
                File.WriteAllText(jsonFilePath, jsonContent);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private bool IsColumnExists(DataTable table, string columnName)
        {
            foreach (DataColumn column in table.Columns)
            {
                if (column.ColumnName == columnName)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Load data from excel file to show on gridview
        /// </summary>
        private void LoadDataFromExcel()
        {
            try
            {
                if (string.IsNullOrEmpty(excelFilePath))
                    return;

                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        MessageBox.Show("File Excel không có dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        // Xóa dữ liệu từ BindingSource
                        bindingSource.DataSource = null;
                        dataGridView1.DataSource = bindingSource;

                        return;
                    }

                    dataTable = new DataTable();

                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

                    // Thêm cột tạm thời để lưu trữ DateTime cho việc sắp xếp
                    if (!IsColumnExists(dataTable, "Ngày hết hạn (For calculate and sort)"))
                    {
                        dataTable.Columns.Add("Ngày hết hạn (For calculate and sort)", typeof(DateTime));
                    }
                    if (!IsColumnExists(dataTable, "Ngày bắt đầu (For calculate and sort)"))
                    {
                        dataTable.Columns.Add("Ngày bắt đầu (For calculate and sort)", typeof(DateTime));
                    }

                    for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                    {
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                        var newRow = dataTable.NewRow();
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
                        }

                        // Chuyển đổi "Ngày hết hạn" thành DateTime
                        if (DateTime.TryParseExact(newRow["Ngày hết hạn"].ToString(), "dd/MM/yyyy", null, DateTimeStyles.None, out DateTime endDate))
                        {
                            newRow["Ngày hết hạn (For calculate and sort)"] = endDate;
                        }

                        // Chuyển đổi "Ngày bắt đầu" thành DateTime
                        if (DateTime.TryParseExact(newRow["Ngày bắt đầu"].ToString(), "dd/MM/yyyy", null, DateTimeStyles.None, out DateTime startDate))
                        {
                            newRow["Ngày bắt đầu (For calculate and sort)"] = startDate;
                        }

                        dataTable.Rows.Add(newRow);
                    }

                    if (dataTable.Rows.Count == 0)
                    {
                        MessageBox.Show("File Excel không có dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        // Xóa dữ liệu từ BindingSource
                        bindingSource.DataSource = null;
                        dataGridView1.DataSource = bindingSource;

                        return;
                    }

                    dataGridView1.DataSource = dataTable;

                    // Sort and highlight rows
                    dataGridView1.Sort(dataGridView1.Columns["Ngày hết hạn (For calculate and sort)"], ListSortDirection.Ascending);
                    HighlightRows();
                    dataGridView1.ClearSelection();

                    // Ẩn cột tạm thời
                    dataGridView1.Columns["Ngày bắt đầu (For calculate and sort)"].Visible = false;
                    dataGridView1.Columns["Ngày hết hạn (For calculate and sort)"].Visible = false;
                }
            }
            catch (Exception ex)
            {
                // Xóa dữ liệu từ BindingSource
                bindingSource.DataSource = null;
                dataGridView1.DataSource = bindingSource;

                MessageBox.Show("Lỗi dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Set highlight for row that has NgayKetThuc < 7
        /// </summary>
        private void HighlightRows()
        {
            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    var cellValue = row.Cells["Ngày hết hạn"].Value;
                    if (cellValue != null && DateTime.TryParseExact(cellValue.ToString(), "dd/MM/yyyy", null, DateTimeStyles.None, out DateTime endDate))
                    {
                        if (endDate < DateTime.Today)
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGray;
                        }
                        else if (endDate < DateTime.Today.AddDays(7))
                        {
                            row.DefaultCellStyle.BackColor = Color.DarkRed;
                            row.DefaultCellStyle.ForeColor = Color.White;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra Ngày hết hạn: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveSortedDataToExcel()
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Sheet1");

                    // Xóa dữ liệu cũ trong worksheet
                    worksheet.Cells.Clear();

                    // Ghi tiêu đề
                    for (int col = 0; col < dataTable.Columns.Count - 2; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                    }

                    // Sắp xếp dữ liệu theo "Ngày hết hạn" tăng dần
                    DataView dv = dataTable.DefaultView;
                    dv.Sort = "Ngày hết hạn (For calculate and sort) ASC";
                    DataTable sortedDataTable = dv.ToTable();

                    // Xóa cột "Ngày hết hạn" trước khi ghi dữ liệu vào file Excel
                    sortedDataTable.Columns.Remove("Ngày bắt đầu (For calculate and sort)");
                    sortedDataTable.Columns.Remove("Ngày hết hạn (For calculate and sort)");

                    // Ghi dữ liệu đã sắp xếp vào worksheet
                    for (int row = 0; row < sortedDataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < sortedDataTable.Columns.Count; col++)
                        {
                            worksheet.Cells[row + 2, col + 1].Value = sortedDataTable.Rows[row][col];
                        }
                    }

                    // Thực hiện việc tô màu
                    ApplyConditionalFormatting(worksheet, sortedDataTable);

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Lỗi khi lưu file Excel: " + ex.Message);
            }
        }

        private void ApplyConditionalFormatting(ExcelWorksheet worksheet, DataTable sortedDataTable)
        {
            // Lấy số dòng và cột
            int rows = sortedDataTable.Rows.Count;
            int columns = sortedDataTable.Columns.Count;

            var cellHeader = worksheet.Cells[1, 1, 1, columns]; // Dòng tiêu đề
            //cellHeader.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cellHeader.Style.Font.Bold = true;

            for (int row = 0; row < rows; row++)
            {
                var expiryDate = Convert.ToDateTime(sortedDataTable.Rows[row]["Ngày hết hạn"]);
                var cell = worksheet.Cells[row + 2, 1, row + 2, columns]; // Dòng hiện tại


                if (expiryDate < DateTime.Today)
                {
                    // Hết hạn
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                else if (expiryDate < DateTime.Today.AddDays(7))
                {
                    // Hết hạn trong vòng 7 ngày
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.DarkRed);
                    cell.Style.Font.Color.SetColor(Color.White);
                }
                //else
                //{
                //    // Còn hạn
                //    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                //    cell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                //}
            }
        }

        /// <summary>
        /// Create new data for Create function
        /// </summary>
        /// <param name="newData"></param>
        /// <exception cref="Exception"></exception>
        private void CreateNewDataToExcel(DataTable newData)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault() ?? package.Workbook.Worksheets.Add("Sheet1");

                    // Nếu không có dữ liệu trong file, thêm tiêu đề
                    if (worksheet.Dimension == null)
                    {
                        for (int col = 1; col <= newData.Columns.Count; col++)
                        {
                            worksheet.Cells[1, col].Value = newData.Columns[col - 1].ColumnName;
                        }
                    }

                    // Tìm dòng cuối cùng trong file
                    int lastUsedRow = worksheet.Dimension?.End.Row ?? 0;

                    // Thêm dữ liệu mới vào từ dòng tiếp theo
                    for (int row = 0; row < newData.Rows.Count; row++)
                    {
                        for (int col = 0; col < newData.Columns.Count; col++)
                        {
                            if (newData.Columns[col].DataType == typeof(DateTime))
                            {
                                DateTime dateValue = DateTime.Parse(newData.Rows[row][col].ToString());
                                worksheet.Cells[lastUsedRow + row + 1, col + 1].Value = dateValue.ToString("dd/MM/yyyy");
                            }
                            else
                            {
                                worksheet.Cells[lastUsedRow + row + 1, col + 1].Value = newData.Rows[row][col];
                            }
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Update data for Edit function
        /// </summary>
        /// <param name="updatedData"></param>
        /// <exception cref="Exception"></exception>
        private void UpdateDataInExcel(DataTable updatedData)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        throw new Exception("Không tìm thấy worksheet.");
                    }

                    string sTT = updatedData.Rows[0]["STT"].ToString();
                    int rowIndex = -1;

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        if (worksheet.Cells[row, 1].Value.ToString() == sTT)
                        {
                            rowIndex = row;
                            break;
                        }
                    }

                    if (rowIndex == -1)
                    {
                        throw new Exception("Không tìm thấy STT để cập nhật.");
                    }

                    for (int col = 1; col <= updatedData.Columns.Count; col++)
                    {
                        if (updatedData.Columns[col - 1].DataType == typeof(DateTime))
                        {
                            DateTime dateValue = DateTime.Parse(updatedData.Rows[0][col - 1].ToString());
                            worksheet.Cells[rowIndex, col].Value = dateValue.ToString("dd/MM/yyyy");
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, col].Value = updatedData.Rows[0][col - 1];
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Delete data for Delete function
        /// </summary>
        /// <exception cref="Exception"></exception>
        private void DeleteDataFromExcel()
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        throw new Exception("Không tìm thấy worksheet.");
                    }

                    foreach (DataGridViewRow selectedRow in dataGridView1.SelectedRows)
                    {
                        string sTT = selectedRow.Cells["STT"].Value.ToString();
                        int rowIndex = -1;

                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, 1].Value.ToString() == sTT)
                            {
                                rowIndex = row;
                                break;
                            }
                        }

                        if (rowIndex != -1)
                        {
                            worksheet.DeleteRow(rowIndex);
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        /// <summary>
        /// Filter data for Search func
        /// </summary>
        /// <param name="searchText"></param>
        private void FilterData(string searchText)
        {
            if (dataTable == null)
                return;

            var filteredTable = dataTable.Clone();
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (DataColumn col in dataTable.Columns)
                {
                    if (row[col].ToString().IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        filteredTable.ImportRow(row);
                        break;
                    }
                }
            }
            dataGridView1.DataSource = filteredTable;

            // Sort and highlight rows
            dataGridView1.Sort(dataGridView1.Columns["Ngày hết hạn (For calculate and sort)"], ListSortDirection.Ascending);
            HighlightRows();
            dataGridView1.ClearSelection();

            // Ẩn cột tạm thời
            dataGridView1.Columns["Ngày hết hạn (For calculate and sort)"].Visible = false;
        }

        /// <summary>
        /// Choose File button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        excelFilePath = openFileDialog.FileName;
                        SaveExcelFilePath();
                        LoadDataFromExcel();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Clear text in Search box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClearSearch_Click(object sender, EventArgs e)
        {
            txtSearch.Text = string.Empty;
            LoadDataFromExcel();
        }

        /// <summary>
        /// Create button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreate_Click(object sender, EventArgs e)
        {
            string jsonFilePath = GetJsonFilePath();
            FormCreateEdit createEditForm = new FormCreateEdit(FormCreateEdit.FormMode.Create, jsonFilePath);
            if (createEditForm.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    
                    CreateNewDataToExcel(createEditForm.detentionData);
                    MessageBox.Show("Tạo mới thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadDataFromExcel();
                    SaveSortedDataToExcel();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi tạo mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Edit button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnEdit_Click(object sender, EventArgs e)
        {
            string jsonFilePath = GetJsonFilePath();
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                FormCreateEdit createEditForm = new FormCreateEdit(FormCreateEdit.FormMode.Edit, jsonFilePath, selectedRow);
                if (createEditForm.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        UpdateDataInExcel(createEditForm.detentionData);
                        MessageBox.Show("Sửa thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDataFromExcel();
                        SaveSortedDataToExcel();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khi sửa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Delete button Click Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa các dòng đã chọn?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        DeleteDataFromExcel();
                        MessageBox.Show("Xóa thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDataFromExcel();
                        SaveSortedDataToExcel();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khi xóa dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn ít nhất một dòng để xóa.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// Set Edit and Delete button is disable when has no row is choose
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                btnEdit.Visible = true;
                btnDelete.Visible = true;
            }
            else
            {
                btnEdit.Visible = false;
                btnDelete.Visible = false;
            }
        }

        private void dataGridView1_Sorted(object sender, EventArgs e)
        {
            HighlightRows();
            dataGridView1.ClearSelection();
        }

        private void FormList_Resize(object sender, EventArgs e)
        {
            lblTitleDanhSach.Location = new Point((this.ClientSize.Width - lblTitleDanhSach.Width) / 2, 30);
        }

        private void FormList_Load(object sender, EventArgs e)
        {
            HighlightRows();
            dataGridView1.ClearSelection();
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            FilterData(txtSearch.Text);
        }

        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Determine which column header was clicked
            DataGridViewColumn clickedColumn = dataGridView1.Columns[e.ColumnIndex];

            // Check if the clicked column is the "Ngày bắt đầu" column
            if (clickedColumn.HeaderText == "Ngày bắt đầu")
            {
                // Perform custom sorting logic
                SortByNgayBatDau();
            }

            // Check if the clicked column is the "Ngày hết hạn" column
            if (clickedColumn.HeaderText == "Ngày hết hạn")
            {
                // Perform custom sorting logic
                SortByNgayHetHan();
            }
        }

        private void SortByNgayBatDau()
        {
            // Đảo chiều sắp xếp nếu cùng cột được chọn
            if (dataGridView1.SortedColumn != null && dataGridView1.SortedColumn.Name == "Ngày bắt đầu")
            {
                sortOrderCustomDate_NgayBatDau = sortOrderCustomDate_NgayBatDau == SortOrderEnum.Ascending ? SortOrderEnum.Descending : SortOrderEnum.Ascending;
            }
            else
            {
                sortOrderCustomDate_NgayBatDau = SortOrderEnum.Ascending;
            }

            // Cập nhật icon sắp xếp trên header cột
            UpdateSortGlyph(dataGridView1.Columns["Ngày bắt đầu"], sortOrderCustomDate_NgayBatDau);

            // Sắp xếp dữ liệu trong DataGridView
            switch (sortOrderCustomDate_NgayBatDau)
            {
                case SortOrderEnum.Ascending:
                    dataGridView1.Sort(dataGridView1.Columns["Ngày bắt đầu (For calculate and sort)"], ListSortDirection.Descending);
                    break;
                case SortOrderEnum.Descending:
                    dataGridView1.Sort(dataGridView1.Columns["Ngày bắt đầu (For calculate and sort)"], ListSortDirection.Ascending);
                    break;
                case SortOrderEnum.None:
                default:
                    break;
            }
        }

        private void SortByNgayHetHan()
        {
            // Đảo chiều sắp xếp nếu cùng cột được chọn
            if (dataGridView1.SortedColumn != null && dataGridView1.SortedColumn.Name == "Ngày hết hạn")
            {
                sortOrderCustomDate_NgayHetHan = sortOrderCustomDate_NgayHetHan == SortOrderEnum.Ascending ? SortOrderEnum.Descending : SortOrderEnum.Ascending;
            }
            else
            {
                sortOrderCustomDate_NgayHetHan = SortOrderEnum.Ascending;
            }

            // Cập nhật icon sắp xếp trên header cột
            UpdateSortGlyph(dataGridView1.Columns["Ngày hết hạn"], sortOrderCustomDate_NgayHetHan);

            // Sắp xếp dữ liệu trong DataGridView
            switch (sortOrderCustomDate_NgayHetHan)
            {
                case SortOrderEnum.Ascending:
                    dataGridView1.Sort(dataGridView1.Columns["Ngày hết hạn (For calculate and sort)"], ListSortDirection.Descending);
                    break;
                case SortOrderEnum.Descending:
                    dataGridView1.Sort(dataGridView1.Columns["Ngày hết hạn (For calculate and sort)"], ListSortDirection.Ascending);
                    break;
                case SortOrderEnum.None:
                default:
                    break;
            }
        }

        private void UpdateSortGlyph(DataGridViewColumn column, SortOrderEnum sortOrder)
        {
            // Xóa icon sắp xếp của các cột khác
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                col.HeaderCell.SortGlyphDirection = SortOrder.None;
            }

            // Thêm icon sắp xếp cho cột hiện tại
            column.HeaderCell.SortGlyphDirection = sortOrder == SortOrderEnum.Ascending ? SortOrder.Ascending : SortOrder.Descending;
        }

    }
}
