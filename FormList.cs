﻿using Newtonsoft.Json;
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

        /// <summary>
        /// Initialization function
        /// </summary>
        public FormList()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            btnEdit.Enabled = false;
            btnDelete.Enabled = false;
            LoadExcelFilePath();
            LoadDataFromExcel();
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
                    dataTable.Columns.Add("Ngày kết thúc (For calculate and sort)", typeof(DateTime));

                    for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                    {
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                        var newRow = dataTable.NewRow();
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
                        }

                        // Chuyển đổi "Ngày kết thúc" thành DateTime
                        if (DateTime.TryParseExact(newRow["Ngày kết thúc"].ToString(), "dd/MM/yyyy", null, DateTimeStyles.None, out DateTime endDate))
                        {
                            newRow["Ngày kết thúc (For calculate and sort)"] = endDate;
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
                    dataGridView1.Sort(dataGridView1.Columns["Ngày kết thúc (For calculate and sort)"], ListSortDirection.Ascending);
                    HighlightRows();
                    dataGridView1.ClearSelection();

                    // Ẩn cột tạm thời
                    dataGridView1.Columns["Ngày kết thúc (For calculate and sort)"].Visible = false;
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
        /// Set highlight for row that has NgayKetThuc < 5
        /// </summary>
        private void HighlightRows()
        {
            try
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    var cellValue = row.Cells["Ngày kết thúc"].Value;
                    if (cellValue != null && DateTime.TryParseExact(cellValue.ToString(), "dd/MM/yyyy", null, DateTimeStyles.None, out DateTime endDate))
                    {
                        if ((endDate - DateTime.Now).TotalDays < 5)
                        {
                            row.DefaultCellStyle.BackColor = Color.DarkRed;
                            row.DefaultCellStyle.ForeColor = Color.White;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra Ngày kết thúc: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                    string maTamGiam = updatedData.Rows[0]["Mã tạm giam"].ToString();
                    int rowIndex = -1;

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        if (worksheet.Cells[row, 2].Value.ToString() == maTamGiam)
                        {
                            rowIndex = row;
                            break;
                        }
                    }

                    if (rowIndex == -1)
                    {
                        throw new Exception("Không tìm thấy Mã tạm giam để cập nhật.");
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
                        string maTamGiam = selectedRow.Cells["Mã tạm giam"].Value.ToString();
                        int rowIndex = -1;

                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            if (worksheet.Cells[row, 2].Value.ToString() == maTamGiam)
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

                    package.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// Check MaTamGiam exist or not in excel file
        /// </summary>
        /// <param name="newData"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private bool CheckMaTamGiamExist(DataTable newData)
        {
            try
            {
                string newMaTamGiam = newData.Rows[0]["Mã tạm giam"].ToString();
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    var cellValue = row.Cells["Mã tạm giam"].Value;
                    if (cellValue != null && cellValue.ToString() == newMaTamGiam)
                    {
                        return true;
                    }
                }
                return false;
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
            dataGridView1.Sort(dataGridView1.Columns["Ngày kết thúc (For calculate and sort)"], ListSortDirection.Ascending);
            HighlightRows();
            dataGridView1.ClearSelection();

            // Ẩn cột tạm thời
            dataGridView1.Columns["Ngày kết thúc (For calculate and sort)"].Visible = false;
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
                    if (!CheckMaTamGiamExist(createEditForm.detentionData))
                    {
                        CreateNewDataToExcel(createEditForm.detentionData);
                        MessageBox.Show("Tạo mới thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadDataFromExcel();
                    }
                    else
                    {
                        MessageBox.Show("Mã tạm giam đã tồn tại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
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
                btnEdit.Enabled = true;
                btnDelete.Enabled = true;
            }
            else
            {
                btnEdit.Enabled = false;
                btnDelete.Enabled = false;
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
    }
}
