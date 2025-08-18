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
    public partial class EngineeringManagementControl : UserControl
    {
        public long UserId;
        public string username;
        public void SetUserData(string user, long id)
        {
            this.username = user;
            this.UserId = id;    // Or store it in a field/property
        }
        public EngineeringManagementControl()
        {
            InitializeComponent();
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
        Dictionary<string, bool> pageAccess = new Dictionary<string, bool>();
        Dictionary<string, bool> functionAccess = new Dictionary<string, bool>();
        void ApplyPermissions(Control parent)
        {
            try
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
            catch (Exception ex)
            {
                ShowAlert("خطأ غير متوقع", AlertForm.AlertType.Error);
                MessageBox.Show(ex.Message,"خطأ غير متوقع",MessageBoxButtons.OK,MessageBoxIcon.Error);
            } 
        }
        private void comboBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            try
            {
            if (e.Index < 0) return;
            ComboBox combo = sender as ComboBox;
            e.DrawBackground();
            using (SolidBrush brush = new SolidBrush(e.ForeColor))
            {
                e.Graphics.DrawString(combo.Items[e.Index].ToString(), e.Font, brush, e.Bounds);
            }
            e.DrawFocusRectangle();
            }
            catch(Exception ex)
            {
                ShowAlert("خطأ غير متوقع", AlertForm.AlertType.Error);
                MessageBox.Show(ex.Message, "خطأ غير متوقع", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void comboBox1_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                ComboBox combo = sender as ComboBox;
                // Clear background
                e.Graphics.Clear(combo.BackColor);
                // Draw text
                TextRenderer.DrawText(e.Graphics, combo.Text, combo.Font,
                    new System.Drawing.Rectangle(2, 3, combo.Width - 20, combo.Height), combo.ForeColor);
                // Draw custom arrow ▼
                Point[] arrow = new Point[]
                {
                    new Point(combo.Width - 15, combo.Height / 2 - 2),
                    new Point(combo.Width - 10, combo.Height / 2 + 2),
                    new Point(combo.Width - 5, combo.Height / 2 - 2)
                };
                e.Graphics.FillPolygon(new SolidBrush(Color.FromArgb(68, 88, 112)), arrow);
            }
            catch (Exception ex)
            {
                ShowAlert("خطأ غير متوقع", AlertForm.AlertType.Error);
                MessageBox.Show(ex.Message, "خطأ غير متوقع", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                // --- Part 1: Fetch data in background ---
                var result = await Task.Run(() =>
                {
                    var LandData = landsTableAdapter.GetData();
                    var Governorates = governorateTableAdapter1.GetData();

                    var joinedData = from land in LandData.AsEnumerable()
                                     join gov in Governorates.AsEnumerable()
                                     on land.Field<string>("governorate_fk") equals gov.Field<string>("governorate_id") into gj
                                     from gov in gj.DefaultIfEmpty()
                                     select new
                                     {
                                         land_id = land.Field<string>("land_id"),
                                         land_name = land.Field<string>("land_name"),
                                         land_number = land.Field<string>("land_number"),
                                         total_area = land.Field<string>("total_area"),
                                         Topographic_Survey_Status = land.Field<string>("Topographic_Survey_Status"),
                                         Land_Plate_Status = land.Field<string>("Land_Plate_Status"),
                                         coordinates_N = land.Field<string>("coordinates_N"),
                                         coordinates_E = land.Field<string>("coordinates_E"),
                                         serial_number = land.Field<string>("serial_number"),
                                         plate_number = land.Field<string>("plate_number"),
                                         plate_numberFile = land.Field<string>("plate_numberFile"),
                                         governorate_fk = gov != null ? gov.Field<string>("governorate") : "",
                                         Republican_Decree = land.Field<string>("Republican_Decree"),
                                         Republican_Decree_Status = land.Field<string>("Republican_Decree_Status"),
                                         consulting_Office = land.Field<string>("consulting_Office"),
                                         total_Land_Price = land.Field<decimal>("total_Land_Price").ToString(),
                                         Ownership_Authority = land.Field<string>("Ownership_Authority"),
                                         Address = land.Field<string>("Address")
                                     };

                    var joinedList = joinedData.ToList();
                    DataTable table = ToDataTable(joinedList);

                    return new
                    {
                        Table = table,
                        Governorates
                    };
                });

                // --- Part 2: Bind data to UI (on UI thread) ---
                Governorate_COB.DataSource = result.Governorates;
                Governorate_COB.DisplayMember = "governorate";
                Governorate_COB.ValueMember = "governorate_id";

                BindingSource bindingSource = new BindingSource();
                bindingSource.DataSource = result.Table;
                advancedDataGridView1.DataSource = bindingSource;

                if (!advancedDataGridView1.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                    checkBoxColumn.HeaderText = "تحديد";
                    checkBoxColumn.Name = "Select";
                    checkBoxColumn.Width = 60;
                    checkBoxColumn.ReadOnly = false;
                    checkBoxColumn.TrueValue = true;
                    checkBoxColumn.FalseValue = false;
                    advancedDataGridView1.Columns.Add(checkBoxColumn);
                }
                advancedDataGridView1.Columns["Select"].DisplayIndex = 0;
            }
            catch (Exception ex)
            {
                ShowAlert("خطأ غير متوقع", AlertForm.AlertType.Error);
                MessageBox.Show(ex.Message, "خطأ غير متوقع", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            FalseFunction();
        }

        public void ApplyGuna2StyleToGrid(DataGridView dgv)
        {
            try
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
            catch (Exception ex)
            {
                ShowAlert("خطأ غير متوقع", AlertForm.AlertType.Error);
                MessageBox.Show(ex.Message, "خطأ غير متوقع", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void GenerativePanalFlow()
        {
            flowLayoutPanel1.AutoScroll = true;
            flowLayoutPanel1.WrapContents = true;  // Items will wrap to next row/column
            flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight; // Or TopDown
        }
        public void flagforHideColumnsEnG()
        {   
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
        private void PositionHeaderButton()
        {
            int headerRight = guna2TabControl1.Left + guna2TabControl1.Width - guna2Button3.Width - 5;
            int headerTop = guna2TabControl1.Top;
            guna2Button3.Location = new Point(headerRight-5, headerTop);
        }
        private async void EngineeringManagementControl_Load(object sender, EventArgs e)
        {
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height-1);
            PositionHeaderButton();
            await LoadSourceDataAsync();
            UpdateRowCount();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            GenerativePanalFlow();
            flagforHideColumnsEnG();
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
        private void dungeonLabel8_Click(object sender, EventArgs e)
        {
        }
        private void guna2ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Republican_Decree_Status_COB.Text);
            Error_Republican_Decree.Visible = !hasText;
        }
        private void tabPage4_Click(object sender, EventArgs e)
        {
        }
        private void tabPage5_Click(object sender, EventArgs e)
        {
        }
        List<string> lastAddedFilePaths = new List<string>(); // at class level
        public void Files()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All files (*.*)|*.*";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    byte[] fileData = File.ReadAllBytes(filePath);

                    // SaveFileToDatabase(filePath, fileName, fileData);
                    lastAddedFilePaths.Add(filePath); // Track it

                    AddFileIconToPanel(filePath, fileName);
                }
            }
        }
        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All files (*.*)|*.*";
            dlg.Multiselect = false;

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                LandfilePath = dlg.FileName;
                AddFileIconToPanel(LandfilePath, Path.GetFileName(LandfilePath));
            }
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
        private bool IsImageFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif";
        }
        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            lastAddedFilePaths.Clear(); // reset
            plateFilePath = "";
            LandfilePath = "";
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
        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            bool hasError = false;
            if (string.IsNullOrWhiteSpace(plate_number_TB.Text))
            {
                Error_plate.Visible = true;
                plate_number_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Land_Name_TB.Text))
            {
                Error_Land.Visible = true;
                Land_Name_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Address_TB.Text))
            {
                Error_Address.Visible = true;
                Address_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Governorate_COB.Text))
            {
                Error_Governorate.Visible = true;
                Governorate_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Ownership_Authority_TB.Text))
            {
                Error_Ownership.Visible = true;
                Ownership_Authority_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(total_land_price_TB.Text))
            {
                Error_total_land.Visible = true;
                total_land_price_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(total_area_TB.Text))
            {
                Error_total_area.Visible = true;
                total_area_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(coordinates_E_TB.Text))
            {
                Error_E.Visible = true;
                coordinates_E_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(coordinates_N_TB.Text))
            {
                Error_N.Visible = true;
                coordinates_N_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Topographic_Survey_Status_COB.Text))
            {
                Error_Topographic.Visible = true;
                Topographic_Survey_Status_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Land_Plate_Status_COB.Text))
            {
                Error_Land_Plate.Visible = true;
                Land_Plate_Status_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(consulting_Office_COB.Text))
            {
                Error_consulting.Visible = true;
                consulting_Office_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Republican_Decree_Status_COB.Text))
            {
                Error_Republican_Decree.Visible = true;
                Republican_Decree_Status_COB.Focus();
                hasError = true;
            }
            if (hasError)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
                return;
            }
            else
            {
                var IDGovernorate = this.governorateTableAdapter1.GetDataByGovernorate(Governorate_COB.Text).First().governorate_id;
                int lastId = int.Parse(this.landsTableAdapter.GetData()
                .OrderByDescending(r => Convert.ToInt32(r.land_id))
                .First()
                .land_id);
                int newId = lastId + 1;
                this.landsTableAdapter.Insert(
                    newId.ToString(),
                    Land_Name_TB.Text
                    ,total_area_TB.Text
                    ,Topographic_Survey_Status_COB.Text
                    , Land_Plate_Status_COB.Text,
                    coordinates_N_TB.Text,
                    coordinates_E_TB.Text,
                    serial_number_TB.Text,
                    plate_number_TB.Text,
                    IDGovernorate,
                    LandfilePath,
                    Republican_Decree_Status_COB.Text,
                    consulting_Office_COB.Text,
                    decimal.Parse(total_land_price_TB.Text),
                    Address_TB.Text,
                    Ownership_Authority_TB.Text,
                    Land_Number_TB.Text,
                    plateFilePath
                    );
                await LoadSourceDataAsync();
                ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
                UpdateRowCount();       
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
                if (column.HeaderText == "مستند القرار الجمهوري" || column.HeaderText== "مستند رقم اللوحة" )
                {
                    checkedListBox1.Items.Remove(column.HeaderText);
                }
            }
        }
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                string header = checkedListBox1.Items[e.Index].ToString();

                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.HeaderText == "مستند القرار الجمهوري" || column.HeaderText == "مستند رقم اللوحة")
                    {
                        break;                    
                    }
                    else if (column.HeaderText == header)
                    {
                        column.Visible = checkedListBox1.GetItemChecked(e.Index);
                        break;
                    }
                }
            });
        }
        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ensure click is not on header row
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var clickedColumn = advancedDataGridView1.Columns[e.ColumnIndex];
            // Check if the clicked column is the "pdf" column
            if (clickedColumn.Name == "Republican_Decree_Status")
            {
                // Get the file path from the cell value
                string filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["Republican_Decree"].Value?.ToString();
                if (!string.IsNullOrWhiteSpace(filePath))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(filePath); // Opens with default PDF reader
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("تعذر فتح الملف: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("الملف غير موجود أو المسار فارغ.");
                }
            }
            if (clickedColumn.Name == "plate_number")
            {
                // Get the file path from the cell value
                string filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["plate_numberFile"].Value?.ToString();
                if (!string.IsNullOrWhiteSpace(filePath))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(filePath); // Opens with default PDF reader
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("تعذر فتح الملف: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("الملف غير موجود أو المسار فارغ.");
                }
            }
        }
        private void comboBox1_Enter(object sender, EventArgs e)
        {
        }
        private void comboBox1_Leave(object sender, EventArgs e)
        {
        }
        private void total_area_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(total_area_TB.Text);
            Error_total_area.Visible = !hasText;
        }
        private void serial_number_TB_TextChanged(object sender, EventArgs e)     
        {
            bool hasText = !string.IsNullOrWhiteSpace(serial_number_TB.Text);
            Error_Serial.Visible = !hasText;
            var x = this.landsTableAdapter.GetDataBySerial(serial_number_TB.Text);
            if (x == null || x.Count == 0)
            {
                // No data found — clear all fields
                plate_number_TB.Text = "";
                Land_Number_TB.Text = "";
                Land_Name_TB.Text = "";
                Address_TB.Text = "";
                Governorate_COB.Text = "";
                Ownership_Authority_TB.Text = "";
                total_land_price_TB.Text = "";
                total_area_TB.Text = "";
                coordinates_E_TB.Text = "";
                coordinates_N_TB.Text = "";
                Topographic_Survey_Status_COB.SelectedIndex = -1;
                Land_Plate_Status_COB.SelectedIndex = -1;
                consulting_Office_COB.SelectedIndex = -1;
                Republican_Decree_Status_COB.SelectedIndex = -1;
                flowLayoutPanel1.Controls.Clear();
                lastAddedFilePaths.Clear();
                return;
            }
            plate_number_TB.Text = x.First().plate_number.ToString();
            Land_Number_TB.Text = x.First().land_number.ToString();
            Land_Name_TB.Text = x.First().land_name.ToString();
            Address_TB.Text = x.First().Address.ToString();
            Governorate_COB.SelectedValue = x.First().governorate_fk;// need to be update 
            Ownership_Authority_TB.Text = x.First().Ownership_Authority.ToString();
            total_land_price_TB.Text = x.First().total_Land_Price.ToString();
            total_area_TB.Text = x.First().total_area.ToString();
            coordinates_E_TB.Text = x.First().coordinates_E.ToString();
            coordinates_N_TB.Text = x.First().coordinates_N.ToString();
            Topographic_Survey_Status_COB.Text = x.First().Topographic_Survey_Status.ToString();
            Land_Plate_Status_COB.Text = x.First().Land_Plate_Status.ToString();
            consulting_Office_COB.Text = x.First().consulting_Office.ToString();
            Republican_Decree_Status_COB.Text = x.First().Republican_Decree_Status.ToString()=="X"? "X": "✔";
            string filePath = x.First().Republican_Decree.ToString();
            string plateFilePath = x.First().plate_numberFile.ToString();
            if (!string.IsNullOrWhiteSpace(filePath))
            {
                lastAddedFilePaths.Add(filePath);
                //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                AddFileIconToPanel(filePath, Path.GetFileName(filePath));
            }
            if (!string.IsNullOrWhiteSpace(plateFilePath))
            {
                lastAddedFilePaths.Add(plateFilePath);
                //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                AddFileIconToPanel(plateFilePath, Path.GetFileName(plateFilePath));
            }
        }
        private void plate_number_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(plate_number_TB.Text);
            Error_plate.Visible =!hasText;
        }
        private void Land_Name_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Land_Name_TB.Text);
            Error_Land.Visible = !hasText;
        }
        private void Address_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Address_TB.Text);
            Error_Address.Visible = !hasText;
        }
        private void Governorate_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Governorate_COB.Text);
            Error_Governorate.Visible = !hasText;
        }
        private void Ownership_Authority_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Ownership_Authority_TB.Text);
            Error_Ownership.Visible = !hasText;
        }
        private void total_land_price_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(total_land_price_TB.Text);
            Error_total_land.Visible = !hasText;
        }
        private void coordinates_E_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(coordinates_E_TB.Text);
            Error_E.Visible = !hasText;
        }
        private void coordinates_N_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(coordinates_N_TB.Text);
            Error_N.Visible = !hasText;
        }
        private void Topographic_Survey_Status_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Topographic_Survey_Status_COB.Text);
            Error_Topographic.Visible = !hasText;
        }
        private void Land_Plate_Status_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Land_Plate_Status_COB.Text);
            Error_Land_Plate.Visible = !hasText;
        }
        private void consulting_Office_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(consulting_Office_COB.Text);
            Error_consulting.Visible = !hasText;
        }
        private void serial_number_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }
        private void plate_number_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }
        private void total_area_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }
        private void total_land_price_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }
        private void nightLabel1_Click(object sender, EventArgs e)
        {
        }
        private void nightLabel2_Click(object sender, EventArgs e)
        {
        }
        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {
        }
        private void nightLabel3_Click(object sender, EventArgs e)
        {
        }
        private async void guna2Button2_Click(object sender, EventArgs e)
        {
            var x = this.landsTableAdapter.GetDataBySerial(serial_number_TB.Text);
            if (x == null || x.Count == 0)
            {
                ShowAlert("No data to update it", AlertForm.AlertType.Error);
                return;
            }
            var IDGovernorate = this.governorateTableAdapter1.GetDataByGovernorate(Governorate_COB.Text).First().governorate_id;
            string filePath = (lastAddedFilePaths != null && lastAddedFilePaths.Count > 0)
                           ? lastAddedFilePaths[0]
                           : null;
            string filePath2 = (lastAddedFilePaths != null && lastAddedFilePaths.Count > 1)
                           ? lastAddedFilePaths[1]
                           : null;
            this.landsTableAdapter.UpdateQuery(Land_Name_TB.Text,
                total_area_TB.Text,
                Topographic_Survey_Status_COB.Text,
                Land_Plate_Status_COB.Text,
                coordinates_N_TB.Text,
                coordinates_E_TB.Text,
                serial_number_TB.Text,
                plate_number_TB.Text,
                IDGovernorate,
                filePath,
                Republican_Decree_Status_COB.Text,
                consulting_Office_COB.Text,
                decimal.Parse(total_land_price_TB.Text),
                Address_TB.Text,
                Ownership_Authority_TB.Text,
                Land_Number_TB.Text,
                filePath2
                );
            await LoadSourceDataAsync();
            ShowAlert("Data updated Successfully", AlertForm.AlertType.Success);
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
        private async void guna2Button3_Click(object sender, EventArgs e)
        {
            // Confirm with user
            try
            {        
            DialogResult result = MessageBox.Show(
                "هل أنت متأكد من أنك تريد حذف الاراضي المحددة؟",
                "تأكيد الحذف",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );
            if (result != DialogResult.Yes)
            {
                return;
            }
            int deleted = 0;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (Convert.ToBoolean(row.Cells["Select"].Value ?? false))
                {
                    string LandId = row.Cells["land_id"].Value.ToString();
                    this.landsTableAdapter.DeleteQuery(LandId);
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
            catch (Exception ex)
            {
                ShowAlert("قم بمسح المشاريع و الاستثمارات المقامه عليها اولا",AlertForm.AlertType.Error);
                MessageBox.Show(ex.Message,"خطا غير متوقع",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
        }
        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
            UpdateRowCount();
        }
        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
        }
        string plateFilePath = null; // for c
        string LandfilePath = null; // for column 2
        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "All files (*.*)|*.*";
            dlg.Multiselect = false;
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                plateFilePath = dlg.FileName;
                AddFileIconToPanel(plateFilePath, Path.GetFileName(plateFilePath));
            }
        }
        private void EngineeringManagementControl_Resize(object sender, EventArgs e)
        {
            PositionHeaderButton();
        }
        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            string name = SessionData.UserName;
            long userId = SessionData.UserId;
            ENGReportForm menu = new ENGReportForm(name, userId);
            menu.Show();
            menu.ShowMenuView();
            // Close the current form that contains this UserControl
            Form parentForm = this.FindForm();
            if (parentForm != null)
            {
                parentForm.Close(); // or parentForm.Hide(); if you just want to hide it
            }
        }
    }
}
