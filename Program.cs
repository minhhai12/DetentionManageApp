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
                var package = new ExcelPackage(new FileInfo(excelFilePath));
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null || worksheet.Dimension == null)
                {
                    return;
                }

                var nearEndCount = 0;
                var overEndCount = 0;

                nearEndCount = worksheet.Cells[2, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                    .Where(cell => cell.Start.Column == worksheet.Dimension.End.Column &&
                                   DateTime.TryParse(cell.Text, out DateTime endDate) &&
                                   (endDate >= DateTime.Today && endDate < DateTime.Today.AddDays(7)))
                    .Count();

                overEndCount = worksheet.Cells[2, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column]
                    .Where(cell => cell.Start.Column == worksheet.Dimension.End.Column &&
                                   DateTime.TryParse(cell.Text, out DateTime endDate) &&
                                   (endDate < DateTime.Today))
                    .Count();

                if (nearEndCount > 0 || overEndCount > 0)
                {
                    string message = $"Có [ {nearEndCount} ] người có Ngày hết hạn tạm giam dưới 7 ngày.\n Và [ {overEndCount} ] người có Ngày hết hạn tạm giam quá hạn.";
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
