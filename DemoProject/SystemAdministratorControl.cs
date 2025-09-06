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
    public partial class SystemAdministratorControl : UserControl
    {
        public long UserId;
        public void SetUserData(string user, long id)
        {
            this.UserId = id;    // Or store it in a field/property
        }
        public SystemAdministratorControl()
        {
            InitializeComponent();
        }
        private void UpdateRowCount()
        {
            int count = 0;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Visible && !row.IsNewRow) // ✅ exclude the new row
                {
                    count++;
                }
            }
            rowCountLabel.Text = $"عدد الصفوف: {count}";
        }
        Dictionary<string, bool> pageAccess = new Dictionary<string, bool>();
        Dictionary<string, bool> functionAccess = new Dictionary<string, bool>();
        void ApplyPermissions(Control parent)
        {
            foreach (Control ctrl in parent.Controls)
            {
                if (ctrl.Tag is string tag)
                {
                    if (tag.StartsWith("Page:"))
                    {
                        string pageName = tag.Substring(5).Trim();
                        ctrl.Visible = pageAccess.ContainsKey(pageName);
                    }
                    else if (tag.StartsWith("Function:"))
                    {
                        string funcName = tag.Substring(9).Trim();
                        ctrl.Enabled = functionAccess.ContainsKey(funcName);
                    }
                }
                // Recurse into nested controls
                if (ctrl.HasChildren)
                    ApplyPermissions(ctrl);
            }
        }
        public void ApplyGuna2StyleToGrid(DataGridView dgv)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToResizeRows = false;
            dgv.BackgroundColor = Color.White;
            dgv.GridColor = Color.LightGray;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 80);
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 10F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.ColumnHeadersHeight = 40;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9.5F, FontStyle.Regular);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(44, 96, 240);
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.BorderStyle = BorderStyle.Fixed3D;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250);
            dgv.RowTemplate.Height = 30;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;      
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
        private void TrueFunction()
        {
            if (this.ParentForm is Form2 form)
            {
                form.ShowLoading(true);   // or false
            }
        }
        private void FalseFunction()
        {
            if (this.ParentForm is Form2 form)
            {
                form.ShowLoading(false);   // or false
            }
        }
        public void LoadSystemAdministratorData()
        {
            TrueFunction();

            // 1. Load permissions / access control
            this.functionsTableAdapter.Fill(this.dATABASE2DataSet.functions);
            this.pagesTableAdapter.Fill(this.dATABASE2DataSet.pages);
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
            this.usersTableAdapter.Fill(this.dATABASE2DataSet.users);
            this.rolesTableAdapter.Fill(this.dATABASE2DataSet.roles);

            tabPage5.Tag = "Function:Print";

            var userRow = dATABASE2DataSet.users.FirstOrDefault(u => u.id == UserId);
            var role = dATABASE2DataSet.roles.FirstOrDefault(r => r.user_id == UserId);

            if (userRow != null && role != null)
            {
                var access = this.accessTableAdapter.GetDataAccsesByRole(role.id);

                foreach (var accessRow in access)
                {
                    // Get page name
                    if (!accessRow.Ispages_idNull())
                    {
                        var pageRow = dATABASE2DataSet.pages.FirstOrDefault(p => p.id == accessRow.pages_id);
                        if (pageRow != null && !string.IsNullOrWhiteSpace(pageRow.Page_name))
                        {
                            string pageName = pageRow.Page_name.Trim();
                            if (!pageAccess.ContainsKey(pageName))
                                pageAccess.Add(pageName, true);
                        }
                    }

                    // Get function name
                    if (!accessRow.Isfunction_idNull())
                    {
                        var funcRow = dATABASE2DataSet.functions.FirstOrDefault(f => f.id == accessRow.function_id);
                        if (funcRow != null && !string.IsNullOrWhiteSpace(funcRow.Function_name))
                        {
                            string funcName = funcRow.Function_name.Trim();
                            if (!functionAccess.ContainsKey(funcName))
                                functionAccess.Add(funcName, true);
                        }
                    }
                }
            }

            ApplyPermissions(this);

            // 2. Load the joined data for grid
            var query =
                from p in projectsTableAdapter.GetData()
                join l in landsTableAdapter.GetData() on p.land_fk equals l.land_id into landGroup
                from l in landGroup.DefaultIfEmpty()
                join g in governorateTableAdapter.GetData() on p.governorate_fk equals g.governorate_id into govGroup
                from g in govGroup.DefaultIfEmpty()
                join inv in investmentsTableAdapter.GetData() on l.land_id equals inv.land_fk into invGroup
                from inv in invGroup.DefaultIfEmpty()
                select new
                {
            // بيانات الأرض
                    مسلسل_الارض = l.land_id,
                    رقم_القطعة = l.land_number,
                    رقم_اللوحة = l?.plate_number,
                    اسم_قطعة_الأرض = l?.land_name,
                    المساحة_الكلية = l?.total_area,
                    حالة_الرفع_المساحي = l?.Topographic_Survey_Status,
                    حالة_لوحة_الأرض = l?.Land_Plate_Status,
                    إحداثيات_شمال = l?.coordinates_N,
                    إحداثيات_شرق = l?.coordinates_E,
                    الرقم_المسلسل = l?.serial_number,
                    القرار_الجمهوري = l?.Republican_Decree,
                    حالة_القرار_الجمهوري = l?.Republican_Decree_Status,
                    المكتب_الاستشاري = l?.consulting_Office,
                    إجمالي_سعر_الأرض = l?.total_Land_Price,
                    جهة_الولاية = l?.Ownership_Authority,
                    العنوان = l?.Address,
            // بيانات المشروع
                    اسم_المشروع = p?.project_name,
                    موافقة_الحماية_المدنية = p?.Civil_Defense_Approval_status,
                    موافقة_البيئة = p?.Environmental_Approval_status,
                    الدراسة_المرورية = p?.Traffic_Study_Status,
                    نموذج_8_أو_10 = p?.Model_8_Status,
                    موافقة_البترول = p?.Petroleum_Ministry_Approval_status,
                    موافقة_الطيران_المدني = p?.Civil_Aviation_Approval_status,
                    رقم_المعاملة = p?.Transaction_number,
                    تاريخ_انتهاء_العقد_المشروع = p?.Contract_expiry_date,
                    اجمالي_المحلات = p?.Total_stores,
                    مؤجر = p?.Total_rented,
                    غير_مؤجر = p?.Total_Not_rented,
            // المحافظة
                    اسم_المحافظة = g?.governorate,
            // بيانات الاستثمار
                    كود_الاستثمار = inv?.investments_id,
                    اسم_الاستثمار = inv?.investment_name,
                    موقع_الاستثمار = inv?.Location,
                    الحي_التابع = inv?.Dependent_neighborhood,
                    نوع_النشاط = inv?.Activity_Type,
                    اسم_النشاط = inv?.Activity_Name,
                    رقم_المحل = inv?.Place_number,
                    رقم_مذكرة_العرض = inv?.Offer_memorandum_number,
                    رقم_العقد_الاستثمار = inv?.Contract_number,
                    تاريخ_بداية_العقد_الاستثمار = inv?.Contract_start_date,
                    تاريخ_انتهاء_العقد_الاستثمار = inv?.Contract_expiry_date,
                    قيمة_الإيجار = inv?.Rental_value,
                    ملف_مذكرة_العرض = inv?.Offer_memorandum_number_File,
                    ملف_العقد = inv?.Contract_number_File
                };

            var joinedList = query.OrderBy(r => int.TryParse(r.مسلسل_الارض, out var n) ? n : int.MaxValue)
                .ToList();
            DataTable original = ToDataTable(joinedList);

            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = original;
            advancedDataGridView1.DataSource = bindingSource;

            UpdateRowCount();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            LoadColumnsIntoCheckedListBox();

            FalseFunction();
        }

        private void SystemAdministratorControl_Load(object sender, EventArgs e)
        {
            LoadSystemAdministratorData();
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
        private static int alertOffsetY = 0; 
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
        private void skyButton2_Click(object sender, EventArgs e)
        {
            TrueFunction();
            if (advancedDataGridView1.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات", AlertForm.AlertType.Error);
                return;
            }

            string fileName = fileNameTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                ShowAlert("يرجى إدخال اسم للملف قبل التصدير", AlertForm.AlertType.Warning);
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

            int colCount = advancedDataGridView1.Columns.Cast<DataGridViewColumn>()
                          .Count(c => c.Visible && c.Name.ToLower() != "select");

            // --- ✅ Add Title Row ---
            var titleRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, colCount]];
            titleRange.Merge();

            // تنسيقات العنوان
            titleRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            titleRange.Value = fileNameTextBox.Text;
            int excelCol = 1;
            for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
            {
                var gridCol = advancedDataGridView1.Columns[col];
                if (gridCol.Visible && gridCol.Name.ToLower() != "select")
                {
                    var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[2, excelCol];
                    cell.Value = gridCol.HeaderText;
                    cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    cell.Font.Bold = true;
                    //Theme color gray

                    excelCol++;
                }
            }

            // ✅ Write data (Row 3 onwards)
            int excelRow = 3;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                excelCol = 1;
                for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
                {
                    var gridCol = advancedDataGridView1.Columns[col];
                    if (!gridCol.Visible || gridCol.Name.ToLower() == "select") continue;

                    var value = row.Cells[col].Value;
                    var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[excelRow, excelCol];

                    if (value is DateTime dtValue) // ✅ لو الخلية تاريخ
                    {
                        cell.Value = dtValue;
                        cell.NumberFormat = "dd/MM/yyyy"; // 🔹 التنسيق المطلوب
                    }
                    else
                    {
                        cell.Value = value != null ? value.ToString() : "";
                    }

                    cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    excelCol++;
                }

                excelRow++;
            }

            // ✅ Auto fit and formatting
            sheet.Cells.Font.Size = 14;
            sheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            sheet.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            sheet.Columns.AutoFit();
            sheet.Rows.AutoFit();
            titleRange.Font.Size = 28;
            titleRange.Font.Bold = true;
            titleRange.RowHeight = 80;



            // ✅ Borders and background cleanup
            int totalRows = excelRow - 1;
            int totalCols = colCount;
            var fullRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[totalRows, totalCols]];

            fullRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            fullRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            fullRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            titleRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            sheet.Cells.WrapText = false;
            // Make headers gray too
            var headerRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, colCount]];
            headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            try
            {
                workbook.SaveAs(fullPath);
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"حدث خطأ أثناء حفظ الملف:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                FalseFunction();
            }

            FalseFunction();
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
                ShowAlert("يرجى إدخال اسم للملف قبل التصدير", AlertForm.AlertType.Warning);             
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
                ShowAlert($"حدث خطأ أثناء التصدير إلى Word:\n{ex.Message}", AlertForm.AlertType.Error);
                MessageBox.Show($"حدث خطأ أثناء التصدير إلى Word:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
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
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            string name = SessionData.UserName;
            long userId = SessionData.UserId;
            ENGReportForm menu = new ENGReportForm(name, userId);
            menu.Show();
            menu.ShowMenuView();
            Form parentForm = this.FindForm();
            if (parentForm != null)
            {
                parentForm.Close(); // or parentForm.Hide(); if you just want to hide it
            }
        }
        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
            UpdateRowCount();
        }

        private void skyButton4_Click(object sender, EventArgs e)
        {
            
        }

        private void reportViewer1_Load(object sender, EventArgs e)
        {

        }
        private void LoadDynamicReport(DataTable dt)
        {
            var stream = ReportHelper.GenerateDynamicRDLC(dt, "MyDataSet");
            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.LoadReportDefinition(stream);
            reportViewer1.LocalReport.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("MyDataSet", dt));
            reportViewer1.RefreshReport();
        }
        private void skyButton4_Click_1(object sender, EventArgs e)
        {
            reportViewer1.Visible = true;

            // 1- فلترة الأعمدة
            DataTable original = ((DataView)((BindingSource)advancedDataGridView1.DataSource).List).ToTable();
            DataTable filtered = new DataTable();
            foreach (string col in checkedListBox1.CheckedItems)
                filtered.Columns.Add(col, original.Columns[col].DataType);

            foreach (DataRow row in original.Rows)
            {
                var newRow = filtered.NewRow();
                foreach (string col in checkedListBox1.CheckedItems)
                    newRow[col] = row[col];
                filtered.Rows.Add(newRow);
            }

            // 2- توليد RDLC بسيط ديناميكي
            string rdlc = GenerateDynamicRDLC(filtered);

            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(rdlc)))
            {
                reportViewer1.LocalReport.LoadReportDefinition(stream);
            }

            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", filtered));
            reportViewer1.RefreshReport();
        }
        private string GenerateDynamicRDLC(DataTable dt)
        {
            using (var ms = ReportHelper.GenerateDynamicRDLC(dt, "DataSet1"))
            {
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private void skyButton5_Click(object sender, EventArgs e)
        {
            
        }

        private void skyButton7_Click(object sender, EventArgs e)
        {
           
        }

        private void skyButton6_Click(object sender, EventArgs e)
        {
           
        }

        private void advancedDataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            advancedDataGridView1.Invalidate();
        }
    }
}
