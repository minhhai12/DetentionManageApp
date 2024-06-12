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

        public FormList()
        {
            InitializeComponent();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            btnEdit.Visible = false;
            LoadDataFromExcel();
        }

        private void btnChooseFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                    LoadDataFromExcel();
                }
            }
        }

        private void LoadDataFromExcel()
        {
            try
            {
                string jsonFilePath = Path.Combine(Application.StartupPath, "excelFilePath.json");
                if (File.Exists(jsonFilePath))
                {
                    string jsonContent = File.ReadAllText(jsonFilePath);
                    dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
                    excelFilePath = jsonData.ExcelFilePath;
                }

                if (string.IsNullOrEmpty(excelFilePath))
                    return;

                var package = new ExcelPackage(new FileInfo(excelFilePath));
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null || worksheet.Dimension == null)
                {
                    MessageBox.Show("File Excel không có dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    return;
                }

                dataGridView1.DataSource = dataTable;

                // Sort and highlight rows
                dataGridView1.Sort(dataGridView1.Columns["NgayKetThucTamGiam"], ListSortDirection.Ascending);

                HighlightRows();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải dữ liệu từ file Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void HighlightRows()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var cellValue = row.Cells["NgayKetThucTamGiam"].Value;
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
            FormCreateEdit createEditForm = new FormCreateEdit();
            if (createEditForm.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (CheckMaTamGiamExist(createEditForm.VehicleData))
                    {
                        MessageBox.Show("MaTamGiam đã tồn tại trong file Excel.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    SaveDataToExcel(createEditForm.VehicleData);
                    LoadDataFromExcel();
                    MessageBox.Show("Tạo mới thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi tạo mới: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private bool CheckMaTamGiamExist(DataTable newData)
        {
            var existingMaTamGiams = dataGridView1.Rows.Cast<DataGridViewRow>()
                .Select(row => row.Cells["MaTamGiam"].Value.ToString())
                .ToList();

            foreach (DataRow row in newData.Rows)
            {
                if (existingMaTamGiams.Contains(row["MaTamGiam"].ToString()))
                    return true;
            }

            return false;
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                var selectedRow = dataGridView1.SelectedRows[0];
                string originalMaTamGiam = selectedRow.Cells["MaTamGiam"].Value.ToString();
                FormCreateEdit createEditForm = new FormCreateEdit(selectedRow);
                if (createEditForm.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (isMaTamGiamChanged(createEditForm.VehicleData, originalMaTamGiam))
                        {
                            if (CheckMaTamGiamExist(createEditForm.VehicleData))
                            {
                                MessageBox.Show("MaTamGiam đã tồn tại trong file Excel.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                        }

                        SaveDataToExcel(createEditForm.VehicleData);
                        LoadDataFromExcel();
                        MessageBox.Show("Sửa thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khi sửa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private bool isMaTamGiamChanged(DataTable newData, string originalMaTamGiam)
        {
            return originalMaTamGiam != newData.Rows[0]["MaTamGiam"].ToString();
        }

        private void SaveDataToExcel(DataTable newData)
        {
            var package = new ExcelPackage(new FileInfo(excelFilePath));
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("Sheet1");
                // Nếu không có dữ liệu trong file, thêm tiêu đề
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

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Hiển thị nút "Sửa" khi chọn bất kỳ dòng nào trong DataGridView
                btnEdit.Visible = true;
            }
        }
    }
}
