using Microsoft.Reporting.WinForms;
using System.Data;
using System.IO;

namespace DemoProject
{
    public static class ExportHelper
    {
        public static void ExportToFile(DataTable dt, string filePath, string format)
        {
            using (var report = new LocalReport())
            {
                // Load dynamic RDLC
                using (var rdlcStream = ReportHelper.GenerateDynamicRDLC(dt))
                {
                    report.LoadReportDefinition(rdlcStream);
                }

                report.DataSources.Clear();
                report.DataSources.Add(new ReportDataSource("MyDataSet", dt));

                string mimeType, encoding, fileNameExtension;
                string[] streams;
                Warning[] warnings;

                byte[] renderedBytes = report.Render(
                    format,           // "PDF", "WORDOPENXML", "EXCELOPENXML"
                    null,
                    out mimeType,
                    out encoding,
                    out fileNameExtension,
                    out streams,
                    out warnings
                );

                File.WriteAllBytes(filePath, renderedBytes);
            }
        }
    }
}
