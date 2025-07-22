using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace DemoProject
{
    public partial class ProjectsControl : UserControl
    {
        public ProjectsControl()
        {
            InitializeComponent();
        }
        private static int alertOffsetY = 0;
        bool MenuExpand = false;
        private void ShowAlert(string msg, AlertForm.AlertType type)
        {
            var alert = new AlertForm(msg, type);
            var screen = Screen.PrimaryScreen.WorkingArea;

            // Position each alert higher than the previous one
            int baseY = screen.Height - alert.Height - 10;
            alert.Location = new Point(screen.Width - alert.Width - 10, baseY - alertOffsetY);

            alertOffsetY += alert.Height + 10;
            alert.Show();

            // Reset offset after a short delay to allow reuse
            var resetTimer = new Timer();
            resetTimer.Interval = 4000; // Slightly longer than fade (3 sec)
            resetTimer.Tick += (s, e) =>
            {
                alertOffsetY = 0;
                resetTimer.Stop();
                resetTimer.Dispose();
            };
            resetTimer.Start();
        }
        private void skyButton1_Click(object sender, EventArgs e)
        {
                if (advancedDataGridView1.Rows.Count == 0)
                {
                    ShowAlert("لا يوجد بيانات للتصدير إلى PDF", AlertForm.AlertType.Error);
                    return;
                }

                string fileName = fileNameTextBox.Text.Trim();
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    MessageBox.Show("يرجى إدخال اسم للملف قبل التصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fullPath = Path.Combine(desktopPath, $"{fileName}.pdf");

                try
                {
                    // PDF setup
                    Document pdfDoc = new Document(PageSize.A4, 20f, 20f, 20f, 20f);
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fullPath, FileMode.Create));
                    pdfDoc.Open();

                    // ✅ Support Arabic + RTL
                    string arialPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                    BaseFont bf = BaseFont.CreateFont(arialPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                    iTextSharp.text.Font headerFont = new iTextSharp.text.Font(bf, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK);

                    // Count visible columns
                    int visibleColCount = advancedDataGridView1.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible);
                    PdfPTable table = new PdfPTable(visibleColCount);
                    table.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                    table.WidthPercentage = 100;

                    // ✅ Add header cells
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (!col.Visible) continue;

                        PdfPCell headerCell = new PdfPCell(new Phrase(col.HeaderText, headerFont));
                        headerCell.BackgroundColor = BaseColor.WHITE;
                        headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        headerCell.BorderWidth = 1;
                        headerCell.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                        table.AddCell(headerCell);
                    }

                    // ✅ Add data rows
                    foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                    {
                        if (row.IsNewRow) continue;

                        foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                        {
                            if (!col.Visible) continue;

                            string cellText = row.Cells[col.Index].Value?.ToString() ?? "";
                            PdfPCell dataCell = new PdfPCell(new Phrase(cellText, font));

                            dataCell.BorderWidth = 1;
                            dataCell.BackgroundColor = BaseColor.WHITE;
                            dataCell.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                            table.AddCell(dataCell);
                        }
                    }

                    pdfDoc.Add(table);
                    pdfDoc.Close();
                    writer.Close();

                    MessageBox.Show("تم حفظ الملف بنجاح على سطح المكتب", "تم", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"حدث خطأ أثناء حفظ ملف PDF:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            
        }
        public static DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            foreach (PropertyDescriptor prop in props)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in props)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
        }
        EngineeringManagementControl em = new EngineeringManagementControl();
        private void skyButton2_Click(object sender, EventArgs e)
        {
            if (advancedDataGridView1.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد الرفع المساحي", AlertForm.AlertType.Error);
                return;
            }

            string fileName = fileNameTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                MessageBox.Show("يرجى إدخال اسم للملف قبل التصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fullPath = Path.Combine(desktopPath, $"{fileName}.xlsx");

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Add(Type.Missing);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            sheet.Name = fileName;

            // ✅ Set sheet direction to RTL
            sheet.DisplayRightToLeft = true;

            int excelCol = 1;

            // ✅ Write headers and align right
            for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
            {
                var gridCol = advancedDataGridView1.Columns[col];
                if (gridCol.Visible)
                {
                    var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[1, excelCol];
                    cell.Value = gridCol.HeaderText;
                    cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    cell.Font.Bold = true;
                    excelCol++;
                }
            }

            // ✅ Write data and align right
            int excelRow = 2;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                excelCol = 1;
                for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
                {
                    var gridCol = advancedDataGridView1.Columns[col];
                    if (!gridCol.Visible) continue;

                    var value = row.Cells[col].Value;
                    var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[excelRow, excelCol];
                    cell.Value = value != null ? value.ToString() : "";
                    cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    excelCol++;
                }

                excelRow++;
            }

            // ✅ Auto fit and formatting
            sheet.Columns.AutoFit();
            sheet.Rows.AutoFit();
            sheet.Cells.Font.Size = 12;

            // ✅ Apply plain borders and remove styling
            int totalRows = excelRow - 1;
            int totalCols = excelCol - 1;
            var fullRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[totalRows, totalCols]];

            // Set borders
            fullRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            fullRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            // Set background to white (remove alternating rows, etc.)
            fullRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            try
            {
                workbook.SaveAs(fullPath);
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"حدث خطأ أثناء حفظ الملف:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProjectsControl_Load(object sender, EventArgs e)
        {
            LoadSourceData();
            em.GenerativePanalFlow();
            em.ApplyGuna2StyleToGrid(advancedDataGridView1);
            ArabicColumnGrid(); 
            LoadColumnsIntoCheckedListBox();
        }
        public void ArabicColumnGrid()
        {
            // Set Arabic headers manually
            advancedDataGridView1.Columns["PlateNumber"].HeaderText = "رقم اللوحة";
            advancedDataGridView1.Columns["LandName"].HeaderText = "اسم قطعة الأرض";
            advancedDataGridView1.Columns["ProjectName"].HeaderText = "اسم المشروع";
            advancedDataGridView1.Columns["GovernorateName"].HeaderText = "اسم المحافظة";

            advancedDataGridView1.Columns["CivilDefenseStatus"].HeaderText = "حالة موافقة الحماية المدنية";
            advancedDataGridView1.Columns["CivilDefenseFile"].HeaderText = "ملف موافقة الحماية المدنية";

            advancedDataGridView1.Columns["EnvironmentalStatus"].HeaderText = "حالة موافقة البيئة";
            advancedDataGridView1.Columns["EnvironmentalFile"].HeaderText = "ملف موافقة البيئة";

            advancedDataGridView1.Columns["PetroleumStatus"].HeaderText = "حالة موافقة وزارة البترول";
            advancedDataGridView1.Columns["PetroleumFile"].HeaderText = "ملف موافقة وزارة البترول";

            advancedDataGridView1.Columns["AviationStatus"].HeaderText = "حالة موافقة الطيران المدني";
            advancedDataGridView1.Columns["AviationFile"].HeaderText = "ملف موافقة الطيران المدني";

            advancedDataGridView1.Columns["TrafficStudyStatus"].HeaderText = "حالة الدراسة المرورية";
            advancedDataGridView1.Columns["TrafficStudyFile"].HeaderText = "ملف الدراسة المرورية";

            advancedDataGridView1.Columns["Model8Status"].HeaderText = "حالة نموذج 8 أو 10";
            advancedDataGridView1.Columns["Model8File"].HeaderText = "ملف نموذج 8 أو 10";
            advancedDataGridView1.Columns["Transaction_number"].HeaderText = "رقم المعاملة";
        }
        public void LoadSourceData()
        {
            var query = from p in projectsTableAdapter.GetData()
                        join g in governorateTableAdapter.GetData() on p.governorate_fk equals g.governorate_id
                        join l in landsTableAdapter.GetData() on p.land_fk equals l.land_id
                        select new
                        {
                            PlateNumber = l.plate_number,
                            LandName = l.land_name,
                            ProjectName = p.project_name,
                            GovernorateName = g.governorate,

                            CivilDefenseStatus = p.Civil_Defense_Approval_status,
                            CivilDefenseFile = (
                            from d in documentsTableAdapter.GetData()
                            join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                            where d.projects_fk == p.project_id && a.approvals == "موافقة الحماية المدنية"
                            select d.paths
                            ).FirstOrDefault(),

                            EnvironmentalStatus = p.Environmental_Approval_status,
                            EnvironmentalFile = (
                            from d in documentsTableAdapter.GetData()
                            join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                            where d.projects_fk == p.project_id && a.approvals == "موافقة البيئة"
                            select d.paths
                            ).FirstOrDefault(),

                            PetroleumStatus = p.Petroleum_Ministry_Approval_status,
                            PetroleumFile = (
                            from d in documentsTableAdapter.GetData()
                            join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                            where d.projects_fk == p.project_id && a.approvals == "موافقة وزارة البترول"
                            select d.paths
                            ).FirstOrDefault(),

                            AviationStatus = p.Civil_Aviation_Approval_status,
                            AviationFile = (
                            from d in documentsTableAdapter.GetData()
                            join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                            where d.projects_fk == p.project_id && a.approvals == "موافقة الطيران المدني"
                            select d.paths
                            ).FirstOrDefault(),

                            TrafficStudyStatus = p.Traffic_Study_Status,
                            TrafficStudyFile = (
                            from d in documentsTableAdapter.GetData()
                            join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                            where d.projects_fk == p.project_id && a.approvals == "الدراسة المرورية"
                            select d.paths
                            ).FirstOrDefault(),

                            Model8Status = p.Model_8_Status,
                            Model8File = (
                            from d in documentsTableAdapter.GetData()
                            join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                            where d.projects_fk == p.project_id && a.approvals == "نموذج 8 أو 10"
                            select d.paths
                            ).FirstOrDefault(),

                            Transaction_number = p.Transaction_number
                        };
            var joinedList = query.ToList();
            DataTable table = ToDataTable(joinedList);
            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = table;
            advancedDataGridView1.DataSource = bindingSource;
        }
        public void LoadColumnsIntoCheckedListBox()
        {
            // Clear previous items
            checkedListBox1.Items.Clear();

            // Make sure the DataGridView has a DataSource
            if (advancedDataGridView1.DataSource == null) return;

            // Loop through the columns and add their HeaderText or Name
            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                checkedListBox1.Items.Add(column.HeaderText, column.Visible); // Show as checked if visible
            }
        }
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        // hide and show columns
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                string header = checkedListBox1.Items[e.Index].ToString();

                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.HeaderText == header)
                    {
                        column.Visible = checkedListBox1.GetItemChecked(e.Index);
                        break;
                    }
                }
            });
        }

        private void skyButton3_Click(object sender, EventArgs e)
        {
            if (advancedDataGridView1.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات للتصدير إلى Word", AlertForm.AlertType.Error);
                return;
            }

            string fileName = fileNameTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                MessageBox.Show("يرجى إدخال اسم للملف قبل التصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fullPath = Path.Combine(desktopPath, $"{fileName}.docx");

            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Add();

                // Set paragraph format to RTL globally
                foreach (Microsoft.Office.Interop.Word.Paragraph para in doc.Paragraphs)
                {
                    para.Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                    para.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                }

                var tableData = new List<List<string>>();

                // Add visible headers
                var headers = new List<string>();
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.Visible)
                        headers.Add(col.HeaderText);
                }
                tableData.Add(headers);

                // Add visible row data
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.IsNewRow) continue;

                    var rowData = new List<string>();
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.Visible)
                        {
                            var value = row.Cells[col.Index].Value;
                            rowData.Add(value?.ToString() ?? "");
                        }
                    }
                    tableData.Add(rowData);
                }

                int rowCount = tableData.Count;
                int colCount = headers.Count;

                var range = doc.Range(0, 0);
                var table = doc.Tables.Add(range, rowCount, colCount);
                table.Borders.Enable = 1;
                table.Range.Font.Size = 11;
                table.Range.Font.Name = "Segoe UI";

                // Set RTL for the whole table
                table.Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                // Fill table cells
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        table.Cell(i + 1, j + 1).Range.Text = tableData[i][j];
                        table.Cell(i + 1, j + 1).Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                        table.Cell(i + 1, j + 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }

                doc.SaveAs2(fullPath);
                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"حدث خطأ أثناء التصدير إلى Word:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

        }
    }
}
