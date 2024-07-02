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
            Application.Run(new FormList());
        }

        private static void CheckVehicleEndDates()
        {
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DetentionManage");
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            string excelFilePath = "";
            string jsonFilePath = Path.Combine(folderPath, "excelFilePath.json");
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
                    return;
                }

                var nearEndCount = worksheet.Cells[2, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                    .Where(cell => cell.Start.Column == worksheet.Dimension.End.Column &&
                                   DateTime.TryParse(cell.Text, out DateTime endDate) &&
                                   (endDate - DateTime.Now).TotalDays < 7)
                    .Count();

                if (nearEndCount > 0)
                {
                    string message = $"Có [ {nearEndCount} ] Số thụ lý có Ngày hết hạn tạm giam dưới 7 ngày.";
                    MessageBox.Show(message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi kiểm tra ngày hết hạn tạm giam: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
