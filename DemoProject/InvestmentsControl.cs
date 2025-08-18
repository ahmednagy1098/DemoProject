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
    public partial class InvestmentsControl : UserControl
    {
        public long UserId;
        public string UserName;
        public void SetUserData(string user, long id)
        {
            this.UserName = user;
            this.UserId = id;    // Or store it in a field/property
        }
        public InvestmentsControl()
        {
            InitializeComponent();

        }
        public void GenerativePanalFlow()
        {
            flowLayoutPanel1.AutoScroll = true;
            flowLayoutPanel1.WrapContents = true;  // Items will wrap to next row/column
            flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight; // Or TopDown
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
            //dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllHeaders;       
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
        public async Task LoadSourceDataAsync()
        {
            try
            {
                TrueFunction();
                await Task.Run(() =>
                {
                    var query = from I in investmentsTableAdapter.GetData()
                                join g in governorateTableAdapter.GetData() on I.governorate_fk equals g.governorate_id
                                join l in landsTableAdapter.GetData() on I.land_fk equals l.land_id
                                select new
                                {
                                    investments_id = I.investments_id,
                                    Land_Id = l.land_id,
                                    Location = I.Location,
                                    PlateNumber = l.plate_number,
                                    Land_Number = l.land_number,
                                    GovernorateName = g.governorate,
                                    Dependent_neighborhood = I.Dependent_neighborhood,
                                    Activity_Type = I.Activity_Type,
                                    Activity_Name = I.Activity_Name,
                                    Place_number = I.Place_number,
                                    Offer_memorandum_number = I.Offer_memorandum_number,
                                    Contract_number = I.Contract_number,
                                    Contract_start_date = I.Contract_start_date,
                                    Contract_expiry_date = I.Contract_expiry_date,
                                    Rental_value = I.Rental_value,
                                    OfferfilePaths = I.Offer_memorandum_number_File,
                                    ContractFilePaths = I.Contract_number_File
                                };

                    var joinedList = query.ToList();
                    DataTable original = ToDataTable(joinedList);

                    // Convert string date column to DateTime type
                    DataTable converted = original.Clone();
                    converted.Columns["Contract_expiry_date"].DataType = typeof(DateTime);
                    converted.Columns["Contract_start_date"].DataType = typeof(DateTime);

                    foreach (DataRow row in original.Rows)
                    {
                        var newRow = converted.NewRow();
                        foreach (DataColumn col in original.Columns)
                        {
                            if (col.ColumnName == "Contract_expiry_date" || col.ColumnName == "Contract_start_date")
                            {
                                if (DateTime.TryParse(row[col].ToString(), out DateTime dt))
                                    newRow[col.ColumnName] = dt;
                                else
                                    newRow[col.ColumnName] = DBNull.Value;
                            }
                            else
                            {
                                newRow[col.ColumnName] = row[col];
                            }
                        }
                        converted.Rows.Add(newRow);
                    }

                    // ✅ Switch back to UI thread to update controls
                    this.Invoke((Action)(() =>
                    {
                        BindingSource bindingSource = new BindingSource();
                        bindingSource.DataSource = converted;
                        advancedDataGridView1.DataSource = bindingSource;

                        if (!advancedDataGridView1.Columns.Contains("Select"))
                        {
                            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                            checkBoxColumn.HeaderText = "تحديد"; // "Select" in Arabic
                            checkBoxColumn.Name = "Select";
                            checkBoxColumn.Width = 60;
                            checkBoxColumn.ReadOnly = false;
                            checkBoxColumn.TrueValue = true;
                            checkBoxColumn.FalseValue = false;
                            advancedDataGridView1.Columns.Add(checkBoxColumn);
                        }

                        advancedDataGridView1.Columns["Select"].DisplayIndex = 0;
                    }));
                });
            }
            catch(Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع"+ex.Message,AlertForm.AlertType.Error);
            }
            FalseFunction();
        }
        public void ArabicColumnGrid()
        {
            // Set Arabic headers manually
            advancedDataGridView1.Columns["Land_Id"].HeaderText = "مسلسل القطعة";
            advancedDataGridView1.Columns["Location"].HeaderText = "الموقع";
            advancedDataGridView1.Columns["PlateNumber"].HeaderText = "رقم اللوحة";
            advancedDataGridView1.Columns["investments_id"].HeaderText = "مسلسل كل محل";
            advancedDataGridView1.Columns["Land_Number"].HeaderText = "رقم قطعة الأرض";
            //advancedDataGridView1.Columns["ProjectName"].HeaderText = "اسم المشروع";
            advancedDataGridView1.Columns["GovernorateName"].HeaderText = "اسم المحافظة";
            advancedDataGridView1.Columns["Dependent_neighborhood"].HeaderText = "الحي التابع";
            advancedDataGridView1.Columns["Activity_Type"].HeaderText = "نوع النشاط";
            advancedDataGridView1.Columns["Activity_Name"].HeaderText = "اسم النشاط";
            advancedDataGridView1.Columns["Place_number"].HeaderText = "رقم المكان";
            advancedDataGridView1.Columns["Offer_memorandum_number"].HeaderText = "رقم مذكرة العرض";
            advancedDataGridView1.Columns["Contract_number"].HeaderText = "رقم العقد";
            advancedDataGridView1.Columns["Rental_value"].HeaderText = "قيمة الايجار";
            advancedDataGridView1.Columns["Contract_start_date"].HeaderText = "تاريخ بداية العقد";
            advancedDataGridView1.Columns["Contract_expiry_date"].HeaderText = "تاريخ انتهاء العقد";
            advancedDataGridView1.Columns["OfferfilePaths"].Visible = false;
            advancedDataGridView1.Columns["ContractFilePaths"].Visible = false;
        }
        public void loadcomboxes()
        {
            LandData = landsTableAdapter.GetData();
            GoverData = governorateTableAdapter.GetData();
            var landNames = LandData.AsEnumerable().Select(r => r["land_name"].ToString()).Distinct().ToList();
            var plateNumbers = LandData.AsEnumerable().Select(r => r["plate_number"].ToString()).Distinct().ToList();
            var landIds = LandData.AsEnumerable().Select(r => r["land_id"].ToString()).Distinct().ToList();
            governorateLookup = GoverData.AsEnumerable()
            .ToDictionary(r => r["governorate_id"].ToString(), r => r["governorate"].ToString());
            var governorateNames = LandData.AsEnumerable()
             .Select(r =>
             {
                 string govId = r["governorate_fk"].ToString();
                 return governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";
             })
             .Distinct()
             .ToList();


            Land_Name_COB.DataSource = landNames;
            plate_number_COB.DataSource = plateNumbers;
            Land_ID_COB.DataSource = landIds;
            Governorate_COB.DataSource = governorateNames;
        }
        private void PositionHeaderButton()
        {
            int headerRight = guna2TabControl1.Left + guna2TabControl1.Width - guna2Button3.Width - 5;
            int headerTop = guna2TabControl1.Top;
            guna2Button3.Location = new Point(headerRight - 5, headerTop);
        }
        private async void InvestmentsControl_Load(object sender, EventArgs e)
        {
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            await LoadSourceDataAsync();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            // TODO: This line of code loads data into the 'database1DataSet.functions' table. You can move, or remove it, as needed.
            this.functionsTableAdapter.Fill(this.database1DataSet.functions);
            // TODO: This line of code loads data into the 'database1DataSet.pages' table. You can move, or remove it, as needed.
            this.pagesTableAdapter.Fill(this.database1DataSet.pages);
            // TODO: This line of code loads data into the 'database1DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.database1DataSet.access);
            // TODO: This line of code loads data into the 'database1DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.database1DataSet.users);
            // TODO: This line of code loads data into the 'database1DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.database1DataSet.roles);
            tabPage4.Tag = "Function:Add";
            tabPage5.Tag = "Function:Print";
            guna2Button3.Tag = "Function:Delete";
            var userRow = database1DataSet.users.FirstOrDefault(u => u.id == UserId);
            var role = database1DataSet.roles.FirstOrDefault(r => r.user_id == UserId);
            if (userRow == null || role == null) return;
            var access = this.accessTableAdapter.GetDataAccsesByRole(role.id);
            foreach (var accessRow in access)
            {
                // Get page name (if page_id exists)
                if (!accessRow.Ispages_idNull())
                {
                    var pageRow = database1DataSet.pages.FirstOrDefault(p => p.id == accessRow.pages_id);
                    if (pageRow != null && !string.IsNullOrWhiteSpace(pageRow.Page_name))
                    {
                        string pageName = pageRow.Page_name.Trim();
                        if (!pageAccess.ContainsKey(pageName))
                            pageAccess.Add(pageName, true);
                    }
                }

                // Get function name (if function_id exists)
                if (!accessRow.Isfunction_idNull())
                {
                    var funcRow = database1DataSet.functions.FirstOrDefault(f => f.id == accessRow.function_id);
                    if (funcRow != null && !string.IsNullOrWhiteSpace(funcRow.Function_name))
                    {
                        string funcName = funcRow.Function_name.Trim();
                        if (!functionAccess.ContainsKey(funcName))
                            functionAccess.Add(funcName, true);
                    }
                }
            }
            ApplyPermissions(this);
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

        private void skyButton2_Click(object sender, EventArgs e)
        {

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

            int excelCol = 1;

            // ✅ Write headers and align right
            for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
            {
                var gridCol = advancedDataGridView1.Columns[col];
                if (gridCol.Visible && gridCol.Name.ToLower() != "select")
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
                    if (!gridCol.Visible || gridCol.Name.ToLower() == "select") continue;

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
                    if (col.Visible && col.Name.ToLower() != "select")
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
                        if (col.Visible && col.Name.ToLower() != "select")
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
                        var cell = table.Cell(i + 1, j + 1);
                        cell.Range.Text = tableData[i][j];
                        cell.Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                        cell.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                        if (i == 0) // Header row
                        {
                            cell.Range.set_Style("Strong"); // Apply bold style explicitly
                            cell.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray25;
                        }
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
                int visibleColCount = advancedDataGridView1.Columns
                    .Cast<DataGridViewColumn>()
                    .Count(c => c.Visible && c.Name.ToLower() != "select");
                PdfPTable table = new PdfPTable(visibleColCount);
                table.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                table.WidthPercentage = 100;

                // ✅ Add header cells
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (!col.Visible || col.Name.ToLower() == "select") continue;

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
                        if (!col.Visible || col.Name.ToLower() == "select") continue;

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

        private async void guna2Button3_Click(object sender, EventArgs e)
        {
            // Show confirmation dialog
            DialogResult result = MessageBox.Show(
                "هل أنت متأكد أنك تريد حذف البيانات المحددة وجميع مستنداتها؟",
                "تأكيد الحذف",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            if (result != DialogResult.Yes)
            {
                return; // User clicked "No" — cancel deletion
            }

            int deleted = 0;

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells["Select"].Value ?? false))
                {
                    string InvestId = row.Cells["investments_id"].Value.ToString();
                    this.investmentsTableAdapter.DeleteQuery(InvestId);
                    deleted++;
                }
            }
            if (deleted > 0)
            {
                await LoadSourceDataAsync();
                ShowAlert($"{deleted} صف تم حذفه بنجاح", AlertForm.AlertType.Success);
            }
            else
            {
                ShowAlert("لم يتم تحديد أي صفوف للحذف", AlertForm.AlertType.Warning);
            }
            UpdateRowCount();
        }

        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            int lastId = int.Parse(this.investmentsTableAdapter.GetData()
              .OrderByDescending(r => Convert.ToInt32(r.investments_id))
              .First()
              .investments_id);

            int newId = lastId + 1;
            var igover = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text);
            bool hasError = false;

            if (string.IsNullOrWhiteSpace(Location_TB.Text))
            {
                Error_Location.Visible = true;
                Location_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Governorate_COB.Text))
            {
                Error_Governorate.Visible = true;
                Governorate_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Land_ID_COB.Text))
            {
                Land_ID_Error.Visible = true;
                Land_ID_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Dependent_neighborhood_TB.Text))
            {
                Error_Dependent_neighborhood.Visible = true;
                Dependent_neighborhood_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Activity_Type_TB.Text))
            {
                Error_Type.Visible = true;
                Activity_Type_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Activity_Name_TB.Text))
            {
                Error_Activiated.Visible = true;
                Activity_Name_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Place_number_Tb.Text))
            {
                Error_place_Number.Visible = true;
                Place_number_Tb.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Offer_memorandum_Number_TB.Text))
            {
                Error_Offer.Visible = true;
                Offer_memorandum_Number_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Contract_number_TB.Text))
            {
                Error_Contract.Visible = true;
                Contract_number_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Rental_value_TB.Text))
            {
                Error_Retal.Visible = true;
                Rental_value_TB.Focus();
                hasError = true;
            }
            if (hasError)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
                return;
            }
            this.investmentsTableAdapter.Insert(
                (newId).ToString(),
                Location_TB.Text,
                "No Name",
                igover.First().governorate_id,
                Land_ID_COB.Text,
                Dependent_neighborhood_TB.Text,
                Activity_Type_TB.Text,
                Activity_Name_TB.Text,
                Place_number_Tb.Text,
                Offer_memorandum_Number_TB.Text,
                Contract_number_TB.Text,
                Start.Value,
                Expiray.Value,
                decimal.Parse(Rental_value_TB.Text),
                OfferfilePath,
                ContractFilePath
                );
            await LoadSourceDataAsync();
            ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
            UpdateRowCount();
        }

        private async void guna2Button2_Click(object sender, EventArgs e)
        {
            var igover = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text);
            var x = this.investmentsTableAdapter.GetDataByInvestId(serial_number_TB.Text);
            if (x == null || x.Count == 0)
            {
                ShowAlert("No data to update it", AlertForm.AlertType.Error);
                return;
            }
            this.investmentsTableAdapter.UpdateQuery(Location_TB.Text,
                "No Name",
                igover.First().governorate_id,
                Land_ID_COB.Text,
                Dependent_neighborhood_TB.Text,
                Activity_Type_TB.Text,
                Activity_Name_TB.Text,
                Place_number_Tb.Text,
                Offer_memorandum_Number_TB.Text,
                Contract_number_TB.Text,
                Start.Value.ToString(),
                Expiray.Value.ToString(),
                decimal.Parse(Rental_value_TB.Text),
                OfferfilePath,
                ContractFilePath,
                serial_number_TB.Text
                );
            await LoadSourceDataAsync();
            ShowAlert("Data updated Successfully", AlertForm.AlertType.Success);

        }

        private async void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            loadcomboxes();
            await LoadSourceDataAsync();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
        }

        private void guna2ImageButton4_Click(object sender, EventArgs e)
        {
             OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All files (*.*)|*.*";
            dlg.Multiselect = false;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                OfferfilePath = dlg.FileName;
                AddFileIconToPanel1(OfferfilePath, Path.GetFileName(OfferfilePath));
            }

        }
        string ContractFilePath = null; 
        string OfferfilePath = null; 
        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
      

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All files (*.*)|*.*";
            dlg.Multiselect = false;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                ContractFilePath = dlg.FileName;
                AddFileIconToPanel(ContractFilePath, Path.GetFileName(ContractFilePath));
            }
            //رقم مذكرة العرض
        }
        private void AddFileIconToPanel(string filePath, string fileName)
        {
            PictureBox picBox = new PictureBox();
            picBox.Size = new Size(64, 64);
            picBox.SizeMode = PictureBoxSizeMode.Zoom;
            picBox.Cursor = Cursors.Hand;
            picBox.BackColor = Color.Transparent;
            // Try to load thumbnail if image
            try
            {
                picBox.Image = Icon.ExtractAssociatedIcon(filePath).ToBitmap();

                // Optional: Show real image if it's an image
                if (IsImageFile(filePath))
                {
                    picBox.Image = System.Drawing.Image.FromFile(filePath);
                }
            }
            catch
            {
                picBox.Image = SystemIcons.Question.ToBitmap(); // fallback icon
            }
            picBox.Click += (s, e) => OpenFile(filePath);
            Label lbl = new Label();
            lbl.Text = fileName;
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            lbl.AutoSize = false;
            lbl.Width = 100;
            lbl.Height = 30;
            lbl.Cursor = Cursors.Hand;
            lbl.Click += (s, e) => OpenFile(filePath); // Label also clickable
            // Container Panel for Icon + Label
            Panel container = new Panel();
            container.Size = new Size(100, 100);
            container.Controls.Add(picBox);
            container.Controls.Add(lbl);

            picBox.Location = new Point(18, 0);
            lbl.Location = new Point(0, 70);

            flowLayoutPanel1.Controls.Add(container);
        }
        private void AddFileIconToPanel1(string filePath, string fileName)
        {
            PictureBox picBox = new PictureBox();
            picBox.Size = new Size(64, 64);
            picBox.SizeMode = PictureBoxSizeMode.Zoom;
            picBox.Cursor = Cursors.Hand;
            picBox.BackColor = Color.Transparent;
            // Try to load thumbnail if image
            try
            {
                picBox.Image = Icon.ExtractAssociatedIcon(filePath).ToBitmap();

                // Optional: Show real image if it's an image
                if (IsImageFile(filePath))
                {
                    picBox.Image = System.Drawing.Image.FromFile(filePath);
                }
            }
            catch
            {
                picBox.Image = SystemIcons.Question.ToBitmap(); // fallback icon
            }
            picBox.Click += (s, e) => OpenFile(filePath);
            Label lbl = new Label();
            lbl.Text = fileName;
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            lbl.AutoSize = false;
            lbl.Width = 100;
            lbl.Height = 30;
            lbl.Cursor = Cursors.Hand;
            lbl.Click += (s, e) => OpenFile(filePath); // Label also clickable
            // Container Panel for Icon + Label
            Panel container = new Panel();
            container.Size = new Size(100, 100);
            container.Controls.Add(picBox);
            container.Controls.Add(lbl);

            picBox.Location = new Point(18, 0);
            lbl.Location = new Point(0, 70);

            flowLayoutPanel7.Controls.Add(container);
        }
        private bool IsImageFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif";
        }
        private void OpenFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = filePath,
                        UseShellExecute = true // ensures it opens with default app
                    });
                else
                    ShowAlert(filePath + "الملف غير موجود", AlertForm.AlertType.Error);
            }
            catch (Exception ex)
            {
                ShowAlert("فشل فتح الملف\n" + ex.Message, AlertForm.AlertType.Error);
                //MessageBox.Show("Failed to open file:\n" + ex.Message);
            }
        }
        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {

        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

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

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "Select" && e.RowIndex >= 0)
            {
                var cell = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                bool isChecked = Convert.ToBoolean(cell.Value ?? false);
                cell.Value = !isChecked;
            }
        }
        List<string> lastAddedFilePaths = new List<string>(); // at class level
        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
            UpdateRowCount();
        }
        private Dictionary<string, string> governorateLookup = new Dictionary<string, string>();
        private DataTable LandData;
        private DataTable GoverData;
        private void serial_number_TB_TextChanged(object sender, EventArgs e)
        {

            
            var x = this.investmentsTableAdapter.GetDataByInvestId(serial_number_TB.Text);
           
            if (x == null || x.Count == 0)
            {
                // No data found — clear all fields
                Location_TB.Text = "";
                Dependent_neighborhood_TB.Text = "";
                Activity_Type_TB.Text = "";
                Activity_Name_TB.Text = "";
                Governorate_COB.Text = "";
                Place_number_Tb.Text = "";
                Offer_memorandum_Number_TB.Text = "";
                Contract_number_TB.Text = "";
                Land_ID_COB.Text = "";
                Land_Name_COB.Text = "";
                plate_number_COB.Text = "";
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel7.Controls.Clear();
                return;
            }
            var l = this.landsTableAdapter.GetDataBySerial(x.First().land_fk);
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            Location_TB.Text = x.First().Location;
            Dependent_neighborhood_TB.Text = x.First().Dependent_neighborhood;
            Activity_Type_TB.Text = x.First().Activity_Type;
            Activity_Name_TB.Text = x.First().Activity_Name;
            Governorate_COB.SelectedItem = governorateLookup.TryGetValue
            (x.First().governorate_fk.ToString(), out var govName)
            ? govName
            : "غير معروف";
            Place_number_Tb.Text = x.First().Place_number;
            Offer_memorandum_Number_TB.Text = x.First().Offer_memorandum_number;
            Contract_number_TB.Text = x.First().Contract_number;
            if (x.First().land_fk == l.First().land_id)
            {
                Land_Name_COB.SelectedItem = l.First().land_name;
                plate_number_COB.SelectedItem = l.First().plate_number;
            }
            Land_ID_COB.SelectedItem = x.First().land_fk;
            Rental_value_TB.Text = x.First().Rental_value.ToString();
            ContractFilePath = x.First().Contract_number_File;
            
            OfferfilePath = x.First().Offer_memorandum_number_File;
            
            if (!string.IsNullOrWhiteSpace(ContractFilePath))
            {
                //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                AddFileIconToPanel(ContractFilePath, Path.GetFileName(ContractFilePath));
            }
            if (!string.IsNullOrWhiteSpace(OfferfilePath))
            {
                lastAddedFilePaths.Add(OfferfilePath);
                //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                AddFileIconToPanel1(OfferfilePath, Path.GetFileName(OfferfilePath));
            }

        }

        private void guna2CircleButton2_Click(object sender, EventArgs e)
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
    }
}
