using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Globalization;

namespace DetentionManageApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            CheckDetentionEndDates();
            Application.Run(new FormList());
        }

        private static void CheckDetentionEndDates()
        {
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DetentionManage");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string excelFilePath = "";
            string jsonFilePath = Path.Combine(folderPath, "data.json");
            if (File.Exists(jsonFilePath))
            {
                string jsonContent = File.ReadAllText(jsonFilePath);
                dynamic jsonData = JsonConvert.DeserializeObject(jsonContent);
                excelFilePath = jsonData.ExcelFilePath;
            }

            if (string.IsNullOrEmpty(excelFilePath))
                return;

            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        return;
                    }

                    // 1. Tìm chính xác các cột cần thiết
                    int ngayHetHanColIndex = -1;
                    int ngayHetHanLan2ColIndex = -1;
                    int soLanColIndex = -1;

                    int maxCol = worksheet.Dimension.End.Column;

                    for (int col = 1; col <= maxCol; col++)
                    {
                        string headerText = worksheet.Cells[1, col].Text.Trim();
                        if (headerText.Equals("Ngày hết hạn", StringComparison.OrdinalIgnoreCase)) ngayHetHanColIndex = col;
                        if (headerText.Equals("Ngày hết hạn lần 2", StringComparison.OrdinalIgnoreCase)) ngayHetHanLan2ColIndex = col;
                        if (headerText.Equals("Số lần", StringComparison.OrdinalIgnoreCase)) soLanColIndex = col;
                    }

                    // Nếu không tìm thấy cột Ngày hết hạn gốc thì thoát
                    if (ngayHetHanColIndex == -1) return;

                    // 2. Duyệt qua từng dòng dữ liệu
                    int maxRow = worksheet.Dimension.End.Row;
                    var dateCells = new System.Collections.Generic.List<DateTime>();

                    for (int row = 2; row <= maxRow; row++)
                    {
                        string soLan = soLanColIndex != -1 ? worksheet.Cells[row, soLanColIndex].Text.Trim() : "";
                        string dateText = "";

                        // Ưu tiên lấy Ngày hết hạn Lần 2 nếu đang ở chế độ Lần 2
                        if (soLan == "Lần 2" && ngayHetHanLan2ColIndex != -1)
                        {
                            dateText = worksheet.Cells[row, ngayHetHanLan2ColIndex].Text.Trim();
                        }
                        else
                        {
                            dateText = worksheet.Cells[row, ngayHetHanColIndex].Text.Trim();
                        }

                        if (DateTime.TryParseExact(dateText, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                        {
                            dateCells.Add(parsedDate);
                        }
                    }

                    // 3. Tính toán số lượng
                    var nearEndCount = dateCells.Count(endDate => endDate >= DateTime.Today && endDate < DateTime.Today.AddDays(7));
                    var overEndCount = dateCells.Count(endDate => endDate < DateTime.Today);

                    // 4. Hiển thị thông báo
                    if (nearEndCount > 0 || overEndCount > 0)
                    {
                        string message = $"Có [ {nearEndCount} ] người có Ngày hết hạn tạm giam dưới 7 ngày.\n Và [ {overEndCount} ] người có Ngày hết hạn tạm giam quá hạn.";
                        MessageBox.Show(message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra ngày hết hạn tạm giam: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}