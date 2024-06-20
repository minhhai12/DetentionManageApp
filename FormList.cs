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
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DetentionManageApp
{
    public partial class FormList : Form
    {
        private string excelFilePath;
        private BindingSource bindingSource = new BindingSource();

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

        private void btnChooseFile_Click(object sender, EventArgs e)
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

        private void LoadExcelFilePath()
        {
            string jsonFilePath = Path.Combine(Application.StartupPath, "excelFilePath.json");
            if (File.Exists(jsonFilePath))
            {
                string jsonContent = File.ReadAllText(jsonFilePath);
                dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
                excelFilePath = jsonData.ExcelFilePath;
            }
        }

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

                    var dataTable = new DataTable();

                    foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                    {
                        dataTable.Columns.Add(firstRowCell.Text);
                    }

                    for (var rowNumber = 2; rowNumber <= worksheet.Dimension.End.Row; rowNumber++)
                    {
                        var row = worksheet.Cells[rowNumber, 1, rowNumber, worksheet.Dimension.End.Column];
                        var newRow = dataTable.NewRow();
                        foreach (var cell in row)
                        {
                            newRow[cell.Start.Column - 1] = cell.Text;
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
                    dataGridView1.Sort(dataGridView1.Columns["Ngày kết thúc"], ListSortDirection.Ascending);
                    HighlightRows();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải dữ liệu từ Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void HighlightRows()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var cellValue = row.Cells["Ngày kết thúc"].Value;
                if (cellValue != null && DateTime.TryParse(cellValue.ToString(), out DateTime endDate))
                {
                    if ((endDate - DateTime.Now).TotalDays < 5)
                    {
                        row.DefaultCellStyle.BackColor = Color.Red;
                    }
                }
            }
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            FormCreateEdit createEditForm = new FormCreateEdit(FormCreateEdit.FormMode.Create);
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

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                FormCreateEdit createEditForm = new FormCreateEdit(FormCreateEdit.FormMode.Edit, selectedRow);
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

        private void CreateNewDataToExcel(DataTable newData)
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
                        worksheet.Cells[lastUsedRow + row + 1, col + 1].Value = newData.Rows[row][col];
                    }
                }

                package.Save();
            }
        }

        private void UpdateDataInExcel(DataTable updatedData)
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
                    worksheet.Cells[rowIndex, col].Value = updatedData.Rows[0][col - 1];
                }

                package.Save();
            }
        }

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
                throw new Exception("Lỗi khi xóa dữ liệu: " + ex.Message);
            }
        }

        private bool CheckMaTamGiamExist(DataTable newData)
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


        private void SaveExcelFilePath()
        {
            string jsonFilePath = Path.Combine(Application.StartupPath, "excelFilePath.json");
            dynamic jsonData = new { ExcelFilePath = excelFilePath };
            string jsonContent = JsonConvert.SerializeObject(jsonData);
            File.WriteAllText(jsonFilePath, jsonContent);
        }

        private void FormList_Resize(object sender, EventArgs e)
        {
            lblTitleDanhSach.Location = new Point((this.ClientSize.Width - lblTitleDanhSach.Width) / 2, 30);
        }
    }
}
