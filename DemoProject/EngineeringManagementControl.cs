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
        public void LoadSourceData()
        {
            try
            {
                TrueFunction();

                // --- Part 1: Fetch data ---
                var LandData = landsTableAdapter.GetData();
                var Governorates = governorateTableAdapter.GetData();

                var joinedData = from land in LandData.AsEnumerable()
                                 join gov in Governorates.AsEnumerable()
                                 on land.Field<string>("governorate_fk") equals gov.Field<string>("governorate_id") into gj
                                 from gov in gj.DefaultIfEmpty()
                                 where !(land.Field<string>("land_id").StartsWith("@")
                                         || land.Field<string>("land_id").StartsWith("#")
                                         || land.Field<string>("land_id").StartsWith("*"))
                                 select new
                                 {
                                     serial_number = land.serial_number,
                                     land_name = land.land_name,
                                     land_number = land.land_number,
                                     plate_number = land.plate_number,
                                     total_area = decimal.TryParse(land.total_area, out var areaVal) ? areaVal : 0,
                                     Topographic_Survey_Status = land.Topographic_Survey_Status,
                                     Land_Plate_Status = land.Land_Plate_Status,
                                     coordinates_N = land.coordinates_N,
                                     coordinates_E = land.coordinates_E,
                                     plate_numberFile = land.plate_numberFile,
                                     governorate_fk = gov != null ? gov.governorate : "",
                                     Republican_Decree = land.Republican_Decree,
                                     Republican_Decree_Status = land.Republican_Decree_Status,
                                     consulting_Office = land.consulting_Office,
                                     total_Land_Price = land.total_Land_Price,
                                     Ownership_Authority = land.Ownership_Authority,
                                     Address = land.Address,
                                     price_per_meter= land.price_per_meter,
                                     land_id = land.land_id
                                 };

                var joinedList = joinedData.OrderBy(r => int.TryParse(r.land_id, out var n) ? n : int.MaxValue)
                .ToList();
                DataTable table = ToDataTable(joinedList);
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                    }
                }

                // --- Part 2: Bind data to UI ---
                Governorate_COB.DataSource = Governorates;
                Governorate_COB.DisplayMember = "governorate";
                Governorate_COB.ValueMember = "governorate_id";
                AdjustComboBox(Governorate_COB);
                BindingSource bindingSource = new BindingSource();
                bindingSource.DataSource = table;
                advancedDataGridView1.DataSource = bindingSource;
                if (advancedDataGridView1.Columns.Contains("total_area"))
                {
                    advancedDataGridView1.Columns["total_area"].DefaultCellStyle.Format = "N2";
                    
                }
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

        public void  TranslateToArabic()
        {
            advancedDataGridView1.Columns["land_id"].HeaderText = "م";
            advancedDataGridView1.Columns["land_name"].HeaderText = "اسم قطعة الأرض";
            advancedDataGridView1.Columns["land_number"].HeaderText = "رقم القطعة";
            advancedDataGridView1.Columns["total_area"].HeaderText = "المساحة الكلية";
            advancedDataGridView1.Columns["Topographic_Survey_Status"].HeaderText = "الرفع المساحي";
            advancedDataGridView1.Columns["Land_Plate_Status"].HeaderText = "استخراج اللوحة";
            advancedDataGridView1.Columns["coordinates_N"].HeaderText = "N الاحداثيات";
            advancedDataGridView1.Columns["coordinates_E"].HeaderText = "E الاحداثيات";
            advancedDataGridView1.Columns["serial_number"].HeaderText = "رقم مسلسل";
            advancedDataGridView1.Columns["plate_number"].HeaderText = "رقم اللوحة";
            advancedDataGridView1.Columns["plate_numberFile"].HeaderText = "مستند رقم اللوحة";
            advancedDataGridView1.Columns["governorate_fk"].HeaderText = "المحافظة";
            advancedDataGridView1.Columns["Republican_Decree"].HeaderText = "مستند القرار الجمهوري";
            advancedDataGridView1.Columns["Republican_Decree_Status"].HeaderText = "عقد الأرض";
            advancedDataGridView1.Columns["consulting_Office"].HeaderText = "المكتب الاستشاري";
            advancedDataGridView1.Columns["total_Land_Price"].HeaderText = "سعر قيمة الارض";
            advancedDataGridView1.Columns["Ownership_Authority"].HeaderText = "جهة الولاية";
            advancedDataGridView1.Columns["Address"].HeaderText = "العنوان";
            advancedDataGridView1.Columns["price_per_meter"].HeaderText = "سعر المتر";
            advancedDataGridView1.Columns["plate_numberFile"].Visible = false;
            advancedDataGridView1.Columns["Republican_Decree"].Visible = false;
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
        private void AdjustComboBox(ComboBox comboBox)
        {
            // 🔸 اضبط العرض حسب أطول عنصر
            int width = comboBox.DropDownWidth;
            using (Graphics g = comboBox.CreateGraphics())
            {
                System.Drawing.Font font = comboBox.Font;
                int vertScrollBarWidth =
                    (comboBox.Items.Count > comboBox.MaxDropDownItems)
                    ? SystemInformation.VerticalScrollBarWidth : 0;

                foreach (var item in comboBox.Items)
                {
                    int newWidth = TextRenderer.MeasureText(item.ToString(), font).Width + vertScrollBarWidth;
                    if (width < newWidth) width = newWidth;
                }
            }
            comboBox.DropDownWidth = width;

            // 🔸 اضبط عدد العناصر اللي تظهر
            comboBox.MaxDropDownItems = 20;
            comboBox.DropDownHeight = 150;
            // 🔸 اضبط ارتفاع كل عنصر (اختياري)
            comboBox.ItemHeight = 22;
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
            HashSet<string> uniqueSeries = new HashSet<string>();
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    var value = row.Cells["serial_number"].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        uniqueSeries.Add(value);
                    }
                }
            }

            rowCountLabel.Text = $"عدد الانشطة: {count}";
            RowCountSeries.Text = $"عدد الاراضي: {uniqueSeries.Count}";
        }
        private void PositionHeaderButton()
        {
            int headerRight = guna2TabControl1.Left + guna2TabControl1.Width - guna2Button3.Width - 5;
            int headerTop = guna2TabControl1.Top;
            guna2Button3.Location = new Point(headerRight-5, headerTop);
        }
        private void EngineeringManagementControl_Load(object sender, EventArgs e)
        {
            TrueFunction();
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height-1);
            PositionHeaderButton();
            LoadSourceData();
            TranslateToArabic();
            UpdateRowCount();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            GenerativePanalFlow();
            flagforHideColumnsEnG();
            LoadColumnsIntoCheckedListBox();
             // TODO: This line of code loads data into the 'database1DataSet.functions' table. You can move, or remove it, as needed.
            this.functionsTableAdapter.Fill(this.dATABASE2DataSet.functions);
            // TODO: This line of code loads data into the 'database1DataSet.pages' table. You can move, or remove it, as needed.
            this.pagesTableAdapter.Fill(this.dATABASE2DataSet.pages);
            // TODO: This line of code loads data into the 'database1DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
            // TODO: This line of code loads data into the 'database1DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.dATABASE2DataSet.users);
            // TODO: This line of code loads data into the 'database1DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.dATABASE2DataSet.roles);
            tabPage4.Tag = "Function:Add";
            tabPage5.Tag = "Function:Print";
            guna2Button3.Tag = "Function:Delete";
            var userRow = dATABASE2DataSet.users.FirstOrDefault(u => u.id == UserId);
            var role = dATABASE2DataSet.roles.FirstOrDefault(r => r.user_id == UserId);
            if (userRow == null || role == null) return;
            var access = this.accessTableAdapter.GetDataAccsesByRole(role.id);
            foreach (var accessRow in access)
            {
                // Get page name (if page_id exists)
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
                // Get function name (if function_id exists)
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
            ApplyPermissions(this);
            FalseFunction();
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
        private void guna2Button1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
              "هل أنت متأكد من أنك تريد اضافة هذه البيانات؟",
              "تأكيد الاضافة",
              MessageBoxButtons.YesNo,
              MessageBoxIcon.Question
          );
            if (result != DialogResult.Yes)
            {
                return;
            }
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
                
                var IDGovernorate = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text).First().governorate_id;
                    int lastId = int.Parse(this.landsTableAdapter.GetData()
                    .OrderByDescending(r => Convert.ToInt32(r.land_id))
                    .First()
                    .land_id);
                    int newId = lastId + 1;
                    this.landsTableAdapter.Insert(
                        newId.ToString(),
                        Land_Number_TB.Text
                        , Land_Name_TB.Text
                        , total_area_TB.Text
                        , Topographic_Survey_Status_COB.Text,
                        Land_Plate_Status_COB.Text,
                        coordinates_N_TB.Text,
                        coordinates_E_TB.Text,
                        serial_number_TB.Text,
                        plate_number_TB.Text,
                        plateFilePath,
                        IDGovernorate,
                        LandfilePath,
                        Republican_Decree_Status_COB.Text,
                        consulting_Office_COB.Text,
                        decimal.Parse(total_land_price_TB.Text),
                        Ownership_Authority_TB.Text,
                        Address_TB.Text,
                        int.Parse(guna2TextBox1.Text.ToString())
                        );
                    LoadSourceData();
                    ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
                    UpdateRowCount();
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
                guna2TextBox1.Text = "";
                Topographic_Survey_Status_COB.SelectedIndex = -1;
                Land_Plate_Status_COB.SelectedIndex = -1;
                consulting_Office_COB.SelectedIndex = -1;
                Republican_Decree_Status_COB.SelectedIndex = -1;
                flowLayoutPanel1.Controls.Clear();
                lastAddedFilePaths.Clear();

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
                    if (column.HeaderText == header)
                    {
                        column.Visible = checkedListBox1.GetItemChecked(e.Index);
                        break;
                    }
                    if (column.HeaderText == "مستند القرار الجمهوري" || column.HeaderText == "مستند رقم اللوحة")
                    {
                        column.Visible = false;
                        continue;                   
                    }
                }
            });
        }
        string LAND_ID = "1";
        string CleanString(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace("\r", "").Replace("\n", "").Trim();
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
            if (Update_Radio.Checked)
            {
                flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                lastAddedFilePaths.Clear();
                try
                {
                    if (e.RowIndex >= 0 && e.RowIndex < advancedDataGridView1.Rows.Count)
                    {
                        DataGridViewRow selectedRow = advancedDataGridView1.Rows[e.RowIndex];

                        // Safe value retrieval with error handling
                        LAND_ID= GetSafeCellValue(selectedRow, "land_id");
                        guna2TextBox1.Text = GetSafeCellValue(selectedRow, "price_per_meter");
                        serial_number_TB.Text = GetSafeCellValue(selectedRow, "serial_number");
                        plate_number_TB.Text = GetSafeCellValue(selectedRow, "plate_number");
                        Land_Number_TB.Text = GetSafeCellValue(selectedRow, "land_number");
                        Land_Name_TB.Text = GetSafeCellValue(selectedRow, "land_name");
                        Address_TB.Text = GetSafeCellValue(selectedRow, "Address");
                        Governorate_COB.Text = GetSafeCellValue(selectedRow, "governorate_fk");
                        Ownership_Authority_TB.Text = GetSafeCellValue(selectedRow, "Ownership_Authority");
                        total_land_price_TB.Text = GetSafeCellValue(selectedRow, "total_Land_Price");
                        total_area_TB.Text = GetSafeCellValue(selectedRow, "total_area");
                        coordinates_E_TB.Text = GetSafeCellValue(selectedRow, "coordinates_E");
                        coordinates_N_TB.Text = GetSafeCellValue(selectedRow, "coordinates_N");
                        string topoStatus = CleanString(GetSafeCellValue(selectedRow, "Topographic_Survey_Status"));
                        int idx1 = Topographic_Survey_Status_COB.FindStringExact(topoStatus);
                        if (idx1 >= 0) Topographic_Survey_Status_COB.SelectedIndex = idx1;

                        string plateStatus = CleanString(GetSafeCellValue(selectedRow, "Land_Plate_Status"));
                        int idx2 = Land_Plate_Status_COB.FindStringExact(plateStatus);
                        if (idx2 >= 0) Land_Plate_Status_COB.SelectedIndex = idx2;

                        string office = CleanString(GetSafeCellValue(selectedRow, "consulting_Office"));
                        int idx3 = consulting_Office_COB.FindStringExact(office);
                        if (idx3 >= 0) consulting_Office_COB.SelectedIndex = idx3;

                        string decreeStatus = CleanString(GetSafeCellValue(selectedRow, "Republican_Decree_Status"));
                        int idx4 = Republican_Decree_Status_COB.FindStringExact(decreeStatus);
                        if (idx4 >= 0) Republican_Decree_Status_COB.SelectedIndex = idx4;
                        plateFilePath = GetSafeCellValue(selectedRow, "plate_numberFile");
                        LandfilePath = GetSafeCellValue(selectedRow, "Republican_Decree");

                        if (!string.IsNullOrWhiteSpace(LandfilePath))
                        {
                            lastAddedFilePaths.Add(LandfilePath);
                            //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                            AddFileIconToPanel(LandfilePath, Path.GetFileName(LandfilePath));
                        }
                        if (!string.IsNullOrWhiteSpace(plateFilePath))
                        {
                            lastAddedFilePaths.Add(plateFilePath);
                            //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
                            AddFileIconToPanel(plateFilePath, Path.GetFileName(plateFilePath));
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"خطأ في تحميل البيانات: {ex.Message}");
                }
            }
        }
        private string GetSafeCellValue(DataGridViewRow row, string columnName)
        {
            try
            {
                if (row.Cells[columnName] != null && row.Cells[columnName].Value != null)
                    return row.Cells[columnName].Value.ToString();
                return string.Empty;
            }
            catch
            {
                return string.Empty;
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
            try
            {
                bool hasText = !string.IsNullOrWhiteSpace(total_area_TB.Text);
                Error_total_area.Visible = !hasText;
                if (string.IsNullOrWhiteSpace(guna2TextBox1.Text)
                    || string.IsNullOrWhiteSpace(total_area_TB.Text))
                {
                    ShowAlert("قم بادخال سعر المتر لحساب الاجمالي", AlertForm.AlertType.Error);
                    return;
                }
                Meter = int.Parse(guna2TextBox1.Text);
                Total_Area = int.Parse(total_area_TB.Text);
                total_land_price_TB.Text = (Meter * Total_Area).ToString();

            }
            catch (Exception ex)
            {
                ShowAlert(ex.Message, AlertForm.AlertType.Error); 
            }

        }
        private void serial_number_TB_TextChanged(object sender, EventArgs e)     
        {
           /* bool hasText = !string.IsNullOrWhiteSpace(serial_number_TB.Text);
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
                guna2TextBox1.Text = "";
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
            guna2TextBox1.Text = x.First().price_per_meter.ToString();
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
            }*/
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
        private void guna2Button2_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "هل أنت متأكد من أنك تريد تعديل هذه البيانات؟",
                "تأكيد التعديل",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (result != DialogResult.Yes)
            {
                return;
            }
            var x = this.landsTableAdapter.GetDataBySerial(serial_number_TB.Text);
            if (x == null || x.Count == 0)
            {
                ShowAlert("لا يوجد هذا البيان للتعديل", AlertForm.AlertType.Error);
                return;
            }
            var IDGovernorate = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text).First().governorate_id;
                this.landsTableAdapter.UpdateQuery(
                Land_Number_TB.Text,
                Land_Name_TB.Text,
                total_area_TB.Text,
                Topographic_Survey_Status_COB.Text,
                Land_Plate_Status_COB.Text,
                coordinates_N_TB.Text,
                coordinates_E_TB.Text,
                serial_number_TB.Text,
                plate_number_TB.Text,
                plateFilePath,
                IDGovernorate,
                LandfilePath,
                Republican_Decree_Status_COB.Text,
                consulting_Office_COB.Text,
                decimal.Parse(total_land_price_TB.Text),
                Ownership_Authority_TB.Text,
                Address_TB.Text,
                int.Parse(guna2TextBox1.Text),
                LAND_ID
                );
            LoadSourceData();
            ShowAlert("تم االتعديل بنجاح", AlertForm.AlertType.Success);
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
        private void guna2Button3_Click(object sender, EventArgs e)
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
                LoadSourceData();
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
        string plateFilePath = ""; // for c
        string LandfilePath = ""; // for column 2
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
            menu.ShowEngView();
            // Close the current form that contains this UserControl
            Form parentForm = this.FindForm();
            if (parentForm != null)
            {
                parentForm.Close(); // or parentForm.Hide(); if you just want to hide it
            }
        }

        private void Add_Radio_CheckedChanged(object sender, EventArgs e)
        {
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button2.Visible = !Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Enabled = !Add_Radio.Checked;
        }

        private void Update_Radio_CheckedChanged(object sender, EventArgs e)
        {
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button1.Visible = !Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button1.Enabled = !Update_Radio.Checked;
        }
        private Dictionary<string, string> columnMap = new Dictionary<string, string>
        {
            { "م", "land_id" },
            { "رقم القطعة", "land_number" },
            { "قطعة الأرض", "land_name" },
            { "المساحة", "total_area" },
            { "الرفع المساحي", "Topographic_Survey_Status" },
            { "استخراج اللوحة", "Land_Plate_Status" },
            { "N الاحداثيات", "coordinates_N" },
            { "E الاحداثيات", "coordinates_E" },
            { "رقم مسلسل", "serial_number" },
            { "المحافظة", "governorate_fk" },
            { "عقد الأرض", "Republican_Decree_Status" },
            { "جهة الولاية", "Ownership_Authority" },
            { "العنوان", "Address" },
            {"سعر قيمة الارض","total_Land_Price" },
            { "المكتب الاستشاري", "consulting_Office" },
            { "سعر المتر", "price_per_meter" }
        };
        private void skyButton4_Click(object sender, EventArgs e)
        {
            reportViewer1.Visible = true;
            reportViewer1.LocalReport.DataSources.Clear();
            DataTable original = ((DataView)((BindingSource)advancedDataGridView1.DataSource).List).ToTable();
            DataTable filtered = new DataTable();

            foreach (string headerText in checkedListBox1.CheckedItems)
            {
                if (columnMap.ContainsKey(headerText)) // map header → real column
                {
                    string colName = columnMap[headerText];
                    filtered.Columns.Add(colName, original.Columns[colName].DataType);
                }
            }

            foreach (DataRow row in original.Rows)
            {
                var newRow = filtered.NewRow();
                foreach (string headerText in checkedListBox1.CheckedItems)
                {
                    if (columnMap.ContainsKey(headerText))
                    {
                        string colName = columnMap[headerText];
                        newRow[colName] = row[colName];
                    }
                }
                filtered.Rows.Add(newRow);
            }

            string rdlc = GenerateDynamicRDLC(filtered);

            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(rdlc)))
            {
                reportViewer1.LocalReport.LoadReportDefinition(stream);
            }

            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(
                new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", filtered));//change this data set to new one
            reportViewer1.RefreshReport();
        }
        private string GenerateDynamicRDLC(DataTable dt)
        {
            using (var ms = ReportHelperEnhanced.GenerateDynamicRDLC(dt, columnMap, "DataSet1",Report_TB.Text))
            {
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }
        int Meter=0;
        int Total_Area = 0;
        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(guna2TextBox1.Text)
                || !string.IsNullOrWhiteSpace(total_area_TB.Text))
            {
                return;
            }
            Meter = int.Parse(guna2TextBox1.Text);
            Total_Area = int.Parse(total_area_TB.Text);
            total_land_price_TB.Text = (Meter * Total_Area).ToString();
        }

        private void guna2TextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }

        private void advancedDataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            advancedDataGridView1.Invalidate();
        }
    }
}
