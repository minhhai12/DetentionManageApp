using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

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

            CheckVehicleEndDates();
            Application.Run(new FormDanhSach());
        }

        private static void CheckVehicleEndDates()
        {
            string excelFilePath = "";
            string jsonFilePath = Path.Combine(Application.StartupPath, "excelFilePath.json");
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
                var package = new ExcelPackage(new FileInfo(excelFilePath));
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null || worksheet.Dimension == null)
                {
                    MessageBox.Show("File Excel không có dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var nearEndVehicles = worksheet.Cells[2, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                    .Where(cell => cell.Start.Column == worksheet.Dimension.End.Column &&
                                   DateTime.TryParse(cell.Text, out DateTime endDate) &&
                                   (endDate - DateTime.Now).TotalDays < 5)
                    .Select(cell => worksheet.Cells[cell.Start.Row, 1, cell.Start.Row, worksheet.Dimension.End.Column].Select(c => c.Text).ToList())
                    .ToList();

                if (nearEndVehicles.Any())
                {
                    string message = "Các phương tiện sau đây có ngày kết thúc tạm giam dưới 5 ngày:\n" +
                                     string.Join("\n", nearEndVehicles.Select(v => string.Join(", ", v)));
                    MessageBox.Show(message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra ngày kết thúc tạm giam: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
