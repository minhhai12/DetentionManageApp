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

                    // 1. Tìm chính xác cột chứa "Ngày hết hạn"
                    int ngayHetHanColIndex = -1;
                    int maxCol = worksheet.Dimension.End.Column;

                    for (int col = 1; col <= maxCol; col++)
                    {
                        string headerText = worksheet.Cells[1, col].Text.Trim();
                        if (headerText.Equals("Ngày hết hạn", StringComparison.OrdinalIgnoreCase))
                        {
                            ngayHetHanColIndex = col;
                            break;
                        }
                    }

                    // Nếu không tìm thấy cột Ngày hết hạn thì thoát để tránh lỗi
                    if (ngayHetHanColIndex == -1)
                    {
                        return;
                    }

                    // 2. Chỉ truy vấn trên đúng cột "Ngày hết hạn" đã tìm thấy, bỏ qua dòng tiêu đề (bắt đầu từ dòng 2)
                    int maxRow = worksheet.Dimension.End.Row;

                    var dateCells = worksheet.Cells[2, ngayHetHanColIndex, maxRow, ngayHetHanColIndex]
                        .Select(cell => cell.Text)
                        .Where(text => !string.IsNullOrEmpty(text))
                        .Select(text =>
                        {
                            DateTime.TryParseExact(text, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate);
                            return parsedDate;
                        })
                        .Where(date => date != DateTime.MinValue); // Bỏ qua các ô không parse được ngày

                    // Tính toán số lượng
                    var nearEndCount = dateCells.Count(endDate => endDate >= DateTime.Today && endDate < DateTime.Today.AddDays(7));
                    var overEndCount = dateCells.Count(endDate => endDate < DateTime.Today);

                    // 3. Hiển thị thông báo
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