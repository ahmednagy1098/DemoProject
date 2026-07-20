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
      
        public long UserId;
        public void SetUserData(string user, long id)
        {
            this.UserId = id;    // Or store it in a field/property
        }
       
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Save filter and sort
            Properties.Settings.Default.LastFilter = advancedDataGridView1.FilterString;
            Properties.Settings.Default.LastSort = advancedDataGridView1.SortString;
            Properties.Settings.Default.Save();
        }
        public ProjectsControl()
        {
            InitializeComponent();
            Petroleum_Ministry_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Civil_Defense_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Civil_Study_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Model_8_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Traffic_Study_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Civil_Aviation_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Environmental_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            //dreamButton1.Cursor = Cursors.Hand;
            //dreamButton1.Click += (s, e) => ShowDocumentList("CivilDefenseFile", "موافقات الحماية المدنية");

            //dreamButton3.Cursor = Cursors.Hand;
            //dreamButton3.Click += (s, e) => ShowDocumentList("EnvironmentalFile", "موافقات البيئة");

            //dreamButton4.Cursor = Cursors.Hand;
            //dreamButton4.Click += (s, e) => ShowDocumentList("PetroleumFile", "موافقات وزارة البترول");

            //dreamButton2.Cursor = Cursors.Hand;
            //dreamButton2.Click += (s, e) => ShowDocumentList("AviationFile", "موافقات الطيران المدني");

            //dreamButton5.Cursor = Cursors.Hand;
            //dreamButton5.Click += (s, e) => ShowDocumentList("TrafficStudyFile", "الدراسة المرورية");

            //dreamButton6.Cursor = Cursors.Hand;
            //dreamButton6.Click += (s, e) => ShowDocumentList("Model8File", "نماذج 8 أو 10");

            //dreamButton16.Cursor = Cursors.Hand;
            //dreamButton16.Click += (s, e) => ShowDocumentList("Transaction_numberFile", "رخصة التشغيل");
            
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
        //EngineeringManagementControl em = new EngineeringManagementControl();
        private void skyButton2_Click(object sender, EventArgs e)
        {

        }

        private Dictionary<string, string> governorateLookup = new Dictionary<string, string>();
        private DataTable LandData;
        private DataTable GoverData;

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
            AdjustComboBox(Land_Name_COB);
            AdjustComboBox(plate_number_COB);
            AdjustComboBox(Land_ID_COB);
            AdjustComboBox(Governorate_COB);
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
                        ctrl.Visible = functionAccess.ContainsKey(funcName);
                    }
                }

                // Recurse into nested controls
                if (ctrl.HasChildren)
                    ApplyPermissions(ctrl);
            }
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
        private void PositionHeaderButton()
        {
            int headerRight = guna2TabControl1.Left + guna2TabControl1.Width - guna2Button3.Width - 5;
            int headerTop = guna2TabControl1.Top;
            guna2Button3.Location = new Point(headerRight - 5, headerTop);
            //Add_Radio.Location = new Point(headerRight - 680, headerTop);
            //Update_Radio.Location = new Point(headerRight-730, headerTop);
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
        private DialogResult ShowArabicMessageBox(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 420,
                Height = 200,
                StartPosition = FormStartPosition.CenterParent,
                Text = caption,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true,
                ShowIcon = false
            };

            Label messageLabel = new Label()
            {
                Text = text,
                Dock = DockStyle.Top,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 10),
                Height = 80
            };

            FlowLayoutPanel buttonsPanel = new FlowLayoutPanel()
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 50
            };

            Button yesButton = new Button() { Text = "نعم", DialogResult = DialogResult.Yes, Width = 80, Height = 35 };
            Button noButton = new Button() { Text = "لا", DialogResult = DialogResult.No, Width = 80, Height = 35 };
            Button cancelButton = new Button() { Text = "إلغاء", DialogResult = DialogResult.Cancel, Width = 80, Height = 35 };

            buttonsPanel.Controls.AddRange(new Control[] { yesButton, noButton, cancelButton });

            prompt.Controls.Add(messageLabel);
            prompt.Controls.Add(buttonsPanel);

            prompt.AcceptButton = yesButton;
            prompt.CancelButton = cancelButton;

            return prompt.ShowDialog();
        }
        private void ShowDocumentList(string columnName, string title)
        {
            var dt = advancedDataGridView1.DataSource as DataTable;
            if (dt == null)
            {
               //MessageBox.Show("البيانات غير متاحة حالياً.", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ✅ Collect visible projects from grid
            var docs = dt.AsEnumerable()
                .Where(r => r.RowState != DataRowState.Deleted)
                .Select(r => new
                {
                    project_name = r["LandName"]?.ToString(),
                    FilePath = r[columnName]?.ToString()
                })
                .ToList();

            if (docs.Count == 0)
            {
                MessageBox.Show("لا توجد مشاريع متاحة حالياً.", "معلومات", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ✅ Ask whether to show existing or missing
            var choice = ShowArabicMessageBox(
    $"هل ترغب بعرض الاراضي التي تحتوي على {title}؟\nاختر (نعم) لعرض الاراضي التي تحتوي، (لا) لعرض الاراضي التي تفتقد المستند.",
    "اختيار نوع العرض");

            if (choice == DialogResult.Cancel)
                return;

            bool showWithDocs = (choice == DialogResult.Yes);

            var filtered = showWithDocs
                ? docs.Where(x => !string.IsNullOrEmpty(x.FilePath)).ToList()
                : docs.Where(x => string.IsNullOrEmpty(x.FilePath)).ToList();

            if (filtered.Count == 0)
            {
                MessageBox.Show("لا توجد نتائج مطابقة.", "معلومات", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // ✅ Create popup (like your bridges popup)
            Form popup = new Form
            {
                StartPosition = FormStartPosition.CenterParent,
                Text = title,
                Size = new Size(500, 600),
                MinimizeBox = false,
                MaximizeBox = false,
                ShowIcon = false,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                RightToLeft = RightToLeft.Yes,
                RightToLeftLayout = true
            };

            Label titleLabel = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 60,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 11, FontStyle.Bold),
                Text = $"{title} ({filtered.Count})"
            };

            ListBox listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 10),
                HorizontalScrollbar = true
            };

            int index = 1;
            foreach (var item in filtered)
            {
                listBox.Items.Add($"{index}. {item.project_name}");
                index++;
            }

            Button closeButton = new Button
            {
                Text = "إغلاق",
                Dock = DockStyle.Bottom,
                Height = 40,
                Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold)
            };
            closeButton.Click += (s, ev) => popup.Close();

            // ✅ Double-click to filter the main grid
            listBox.DoubleClick += (s, ev) =>
            {
                if (listBox.SelectedItem == null)
                    return;

                string selectedText = listBox.SelectedItem.ToString();
                int dotIndex = selectedText.IndexOf('.');
                string selectedName = (dotIndex >= 0)
                    ? selectedText.Substring(dotIndex + 1).Trim()
                    : selectedText.Trim();

                try
                {
                    if (dt == null)
                    {
                        MessageBox.Show("البيانات الأصلية غير متوفرة لتطبيق الفلتر.",
                                        "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    DataView view = new DataView(dt)
                    {
                        RowFilter = $"LandName = '{selectedName.Replace("'", "''")}'"
                    };

                    advancedDataGridView1.DataSource = view;
                    UpdateRowCount();
                    popup.Close();
                    ShowAlert($"تم عرض بيانات الارض: {selectedName}", AlertForm.AlertType.Info);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"حدث خطأ أثناء تطبيق الفلترة:\n{ex.Message}",
                                    "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            popup.Controls.Add(listBox);
            popup.Controls.Add(closeButton);
            popup.Controls.Add(titleLabel);
            popup.ShowDialog(this);
        }

        private void ProjectsControl_Load(object sender, EventArgs e)
        {
            
            guna2TabControl1.DrawMode = TabDrawMode.OwnerDrawFixed;
            guna2TabControl1.DrawItem += guna2TabControl1_DrawItem;

            //guna2TabControl1.SelectedTab = guna2TabControl1.TabPages[2];
            //guna2TabControl1.SelectedIndex = 2;
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            Add_Radio.Parent = guna2TabControl1.Parent;
            Add_Radio.BringToFront();
            Update_Radio.Parent = guna2TabControl1.Parent;
            Update_Radio.BringToFront();
            
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            LoadProjectData();
            LoadNotes();
            SaveState();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ArabicColumnGrid(); 
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
            tabPage4.Tag = "Function:Edit";
            tabPage2.Tag = "Function:Edit";
            tabPage5.Tag = "Function:Print";
            guna2Button3.Tag = "Function:Delete";
            crownLabel1.Tag= "Function:Delete";
            guna2Shapes1.Tag = "Function:Delete";
            guna2Shapes2.Tag = "Function:Delete";
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
        }
        public void ArabicColumnGrid()
        {
            // Set Arabic headers manually
            if (fillterProjects=="")
            {
                advancedDataGridView1.Columns["land_fk"].HeaderText = "مسلسل القطعة الكود";
                advancedDataGridView1.Columns["Project_Id"].HeaderText = "مسلسل المشروع";
                advancedDataGridView1.Columns["Project_Id"].Visible = false;
                advancedDataGridView1.Columns["PlateNumber"].HeaderText = "رقم اللوحة";
                advancedDataGridView1.Columns["LandName"].HeaderText = "اسم قطعة الأرض";
                advancedDataGridView1.Columns["Land_Number"].HeaderText = "رقم قطعة الأرض";
                advancedDataGridView1.Columns["project_name"].HeaderText = "اسم المشروع / نشاط";
                advancedDataGridView1.Columns["Address"].HeaderText = "العنوان";
                advancedDataGridView1.Columns["area"].HeaderText = "المساحة";
                advancedDataGridView1.Columns["project_name"].Visible = false;
                advancedDataGridView1.Columns["LandName"].Visible = false;
                advancedDataGridView1.Columns["GovernorateName"].HeaderText = "اسم المحافظة";
                advancedDataGridView1.Columns["Name_Projects"].HeaderText = "اسم المكان";
                advancedDataGridView1.Columns["Civil_Defense_Approval_status"].HeaderText = "حالة موافقة الحماية المدنية";
                advancedDataGridView1.Columns["CivilDefenseFile"].HeaderText = "ملف موافقة الحماية المدنية";

                advancedDataGridView1.Columns["Civil_Defense_Study_status"].HeaderText = "حالة دراسة حماية المدنية";
                advancedDataGridView1.Columns["CivilDefenseStudyFile"].HeaderText = "ملف دراسة حماية المدنية";

                advancedDataGridView1.Columns["Environmental_Approval_status"].HeaderText = "حالة موافقة البيئة";
                advancedDataGridView1.Columns["EnvironmentalFile"].HeaderText = "ملف موافقة البيئة";

                advancedDataGridView1.Columns["Petroleum_Ministry_Approval_status"].HeaderText = "حالة موافقة وزارة البترول";
                advancedDataGridView1.Columns["PetroleumFile"].HeaderText = "ملف موافقة وزارة البترول";

                advancedDataGridView1.Columns["Civil_Aviation_Approval_status"].HeaderText = "حالة موافقة الطيران المدني";
                advancedDataGridView1.Columns["AviationFile"].HeaderText = "ملف موافقة الطيران المدني";

                advancedDataGridView1.Columns["Traffic_Study_Status"].HeaderText = "حالة الدراسة المرورية";
                advancedDataGridView1.Columns["TrafficStudyFile"].HeaderText = "ملف الدراسة المرورية";

                advancedDataGridView1.Columns["Model_8_Status"].HeaderText = "حالة نموذج 8 أو 10";
                advancedDataGridView1.Columns["Model8File"].HeaderText = "ملف نموذج 8 أو 10";
                advancedDataGridView1.Columns["Transaction_number"].HeaderText = "رخصة التشغيل";
                advancedDataGridView1.Columns["Transaction_numberFile"].HeaderText = "ملف رقم المعاملة";
                advancedDataGridView1.Columns["Contract_expiry_date"].HeaderText = "تاريخ انتهاء العقد";
                //advancedDataGridView1.Columns["Total_stores"].HeaderText = "اجمالي محلات";
                //advancedDataGridView1.Columns["Total_rented"].HeaderText = "مؤجر";
                //advancedDataGridView1.Columns["Total_Not_rented"].HeaderText = "غير مؤجر";
                advancedDataGridView1.Columns["Secured_Certificate"].HeaderText = "ملاحظات";
                advancedDataGridView1.Columns["Architectural_and_Structural_Board"].HeaderText = "شركة ادارة المحطة";
                // advancedDataGridView1.Columns["Reconciliation_Form_Stamp"].HeaderText = "رخصة التشغيل";
                //advancedDataGridView1.Columns["Consultant_Surveying"].HeaderText = "الرفع المساحي الاستشاري";
                advancedDataGridView1.Columns["consulting_Office"].HeaderText = "المكتب الاستشاري";
                advancedDataGridView3.Columns["Id"].HeaderText = "مسلسل مميز";
                advancedDataGridView3.Columns["Note_Page"].HeaderText = "الملاحظات";
                advancedDataGridView3.Columns["Code"].HeaderText = "مسلسل";
                advancedDataGridView3.Columns["DateTime"].HeaderText = "التاريخ";
                advancedDataGridView3.Columns["status"].HeaderText = "موقف الملحوظه";
                advancedDataGridView3.Columns["Id"].Visible = false;
            }
            else
            {
                advancedDataGridView1.Columns["Name_Projects"].HeaderText = "اسم المكان";
                advancedDataGridView1.Columns["land_fk"].HeaderText = "مسلسل القطعة الكود";
               
                 
                       advancedDataGridView1.Columns["Project_Id"].HeaderText = "مسلسل المشروع";
                advancedDataGridView1.Columns["Project_Id"].Visible = false;
            
                          advancedDataGridView1.Columns["Address"].HeaderText = "العنوان";
                advancedDataGridView1.Columns["GovernorateName"].HeaderText = "اسم المحافظة";

                //advancedDataGridView1.Columns["consulting_Office"].HeaderText = "المكتب الاستشاري";

                //advancedDataGridView1.Columns["Architectural_and_Structural_Board"].HeaderText = "شركة ادارة المحطة";



                advancedDataGridView1.Columns["Contract_expiry_date"].HeaderText = "تاريخ انتهاء العقد";
                advancedDataGridView1.Columns["Secured_Certificate"].HeaderText = "ملاحظات";
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
        public void loadDocs()
        {
            try
            {
                TrueFunction();

                // --- Part 1: Fetch and join data ---
                var query =
                    from p in projectsTableAdapter.GetData()
                    join g in governorateTableAdapter.GetData() on p.governorate_fk equals g.governorate_id
                    join l in landsTableAdapter.GetData() on p.land_fk equals l.land_id
                    join d in documentsTableAdapter.GetData() on p.project_id equals d.projects_fk
                    join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                    where a.approvals == "عقود محلات"
                    select new
                    {
                        document_id = d.document_id,
                        Project_Id = p.project_id,
                        Name_Projects = p.Name_Projects,
                        project_name = p.project_name,
                        MarketsFile = d.paths   // ملف واحد في كل صف
                    };

                var joinedList = query
                    .OrderBy(r => int.TryParse(r.Project_Id, out var n) ? n : int.MaxValue)
                    .ToList();

                DataTable original = ToDataTable(joinedList);

                // --- Bind data to DataGridView ---
                advancedDataGridView2.DataSource = original;

                if (!advancedDataGridView2.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
                    checkBoxColumn.HeaderText = "تحديد"; // "Select" in Arabic
                    checkBoxColumn.Name = "Select";
                    checkBoxColumn.Width = 60;
                    checkBoxColumn.ReadOnly = false;
                    checkBoxColumn.TrueValue = true;
                    checkBoxColumn.FalseValue = false;
                    advancedDataGridView2.Columns.Add(checkBoxColumn);
                }

                advancedDataGridView2.Columns["Select"].DisplayIndex = 0;
                advancedDataGridView2.Columns["Project_Id"].HeaderText = "مسلسل مشروع";
                advancedDataGridView2.Columns["project_name"].HeaderText = "اسم المشروع / النشاط";
                advancedDataGridView2.Columns["Name_Projects"].HeaderText = "اسم المكان";
                advancedDataGridView2.Columns["MarketsFile"].HeaderText = "ملف المحل";
                advancedDataGridView2.Columns["document_id"].Visible = false;
                advancedDataGridView2.Columns["document_id"].HeaderText = "مسلسل المستند";
                advancedDataGridView2.Columns["project_name"].Visible = false;
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            finally
            {
                FalseFunction();
            }
        }
        public void loadInvestmentsWithoutDocs()
        {
            try
            {
                TrueFunction();

                var projects = projectsTableAdapter.GetData();
                var investments = investmentsTableAdapter.GetData();
                var documents = documentsTableAdapter.GetData();
                var approvals = approvalsTableAdapter.GetData();

                // هات كل الـ paths الخاصة بعقود المحلات
                var docPaths =
                    (from d in documents
                     join a in approvals on d.approvals_fk equals a.approval_id
                     where a.approvals == "عقود محلات"
                     select d.paths?.Trim().ToLower()).ToList();

                // هات الاستثمارات مع المشاريع
                var query =
                    from i in investments
                    join p in projects on i.land_fk equals p.land_fk
                    where !docPaths.Any(path => !string.IsNullOrEmpty(path) &&
                                                path.Contains(i.Activity_Name.Trim().ToLower()))
                    select new
                    {
                        Project_Id = p.project_id,
                        Name_Projects = p.Name_Projects,  // ✅ هنا جبنا العمود اللي طلبته
                        project_name = p.project_name,
                        InvestmentName = i.investment_name,
                        Activity_Name = i.Activity_Name
                    };

                var list = query
                    .OrderBy(r => int.TryParse(r.Project_Id, out var n) ? n : int.MaxValue)
                    .ToList();

                DataTable original = ToDataTable(list);

                advancedDataGridView2.DataSource = original;

                advancedDataGridView2.Columns["Activity_Name"].Visible = false;
                advancedDataGridView2.Columns["InvestmentName"].HeaderText = "اسم الاستثمار";
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            finally
            {
                FalseFunction();
            }
        }
        public void LoadProjectData()
        {
            DataTable converted = null;
            try
            {
                TrueFunction();

                // Step 1: Load all data once
                var projects = projectsTableAdapter.GetData().ToList();
                var governorates = governorateTableAdapter.GetData().ToList();
                var lands = landsTableAdapter.GetData().ToList();
                var documents = documentsTableAdapter.GetData().ToList();
                var approvals = approvalsTableAdapter.GetData().ToList();
                var investments = investmentsTableAdapter.GetData().ToList();

                // Step 2: Pre-join documents + approvals to avoid repeating lookups
                var docsWithApprovals = (
                    from d in documents
                    join a in approvals on d.approvals_fk equals a.approval_id
                    select new
                    {
                        ProjectId = d.projects_fk,
                        ApprovalName = a.approvals,
                        Path = d.paths
                    }
                ).ToList();

                // Step 3: Create main query (pure in-memory join)
                var query =
                    from p in projects
                    join g in governorates on p.governorate_fk equals g.governorate_id
                    join l in lands on p.land_fk equals l.land_id
                    let inv = investments.Where(x => x.land_fk == l.land_id)
                    select new
                    {

                        project_name = p.project_name,
                        LandName = l.Island_nameNull() ? "" : l.land_name,
                        Name_Projects = p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                        Address = l.IsAddressNull() ? "" : l.Address,
                        GovernorateName = g.IsgovernorateNull() ? "" : g.governorate,
                        PlateNumber = l.Isplate_numberNull() ? "" : l.plate_number,
                        Land_Number = l.Island_numberNull() ? "" : l.land_number,
                        land_fk = p.Island_fkNull()?"":p.land_fk,
                        Project_Id = p.project_id,
                        area = l.Istotal_areaNull()?"":l.total_area,
                        //Total_stores = inv.Where(i => i.Visable_Value == true).Sum(i => i.IsShops_CountNull() ? 0 : i.Shops_Count),

                        //                    Total_rented =
                        //inv.Where(i =>
                        //    (!i.IsActivity_NameNull() && i.Activity_Name != "لا يوجد") &&
                        //    i.IsRental_StatusNull() && i.Visable_Value == true
                        //)
                        //.Sum(i => i.IsShops_CountNull() ? 0 : i.Shops_Count),

                        //                    Total_Not_rented =
                        //inv.Where(i =>
                        //    (!i.IsActivity_NameNull() && i.Activity_Name == "لا يوجد") &&
                        //    (i.IsRental_StatusNull()||i.Rental_Status== "غير مؤجر (*)") && i.Visable_Value == true
                        //)
                        //.Sum(i => i.IsShops_CountNull() ? 0 : i.Shops_Count),

                        consulting_Office = l.Isconsulting_OfficeNull()?"":l.consulting_Office,
                        
                        Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull()?"":p.Architectural_and_Structural_Board,
                        //Reconciliation_Form_Stamp = p.IsReconciliation_Form_StampNull()?"":p.Reconciliation_Form_Stamp,
                        Consultant_Surveying = p.Consultant_Surveying,

                        Civil_Defense_Approval_status = p.Civil_Defense_Approval_status,
                        CivilDefenseFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة الحماية المدنية")
                .Select(d => d.Path)),

                        Civil_Defense_Study_status = p.Civil_Defense_Study_status,
                        CivilDefenseStudyFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "دراسة حماية مدنيه")
                .Select(d => d.Path)),

                        Environmental_Approval_status = p.Environmental_Approval_status,
                        EnvironmentalFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة البيئة")
                .Select(d => d.Path)),

                        Petroleum_Ministry_Approval_status = p.Petroleum_Ministry_Approval_status,
                        PetroleumFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة وزارة البترول")
                .Select(d => d.Path)),

                        Civil_Aviation_Approval_status = p.Civil_Aviation_Approval_status,
                        AviationFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة الطيران المدني")
                .Select(d => d.Path)),

                        Traffic_Study_Status = p.Traffic_Study_Status,
                        TrafficStudyFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "الدراسة المرورية")
                .Select(d => d.Path)),

                        Model_8_Status = p.Model_8_Status,
                        Model8File = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "نموذج 8 أو 10")
                .Select(d => d.Path)),

                        Transaction_number = p.Transaction_number,
                        Transaction_numberFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "رقم المعاملة")
                .Select(d => d.Path)),
                        Contract_expiry_date = p.Contract_expiry_date,
                        Secured_Certificate = p.IsSecured_certificateNull()?"": p.Secured_certificate
                    };
                
                var joinedList = query.OrderBy(r => int.TryParse(r.Project_Id, out var n) ? n : int.MaxValue)
                    .Where(r=>r.Project_Id!="0"&&r.land_fk != "215" && r.land_fk != "216"&& r.Secured_Certificate!= "تم الغاء قطعة الارض" &&
                    !r.land_fk.Contains("$")&& !r.land_fk.Contains("#")&& !r.land_fk.Contains("&")
                    )
                    .ToList();
                DataTable original = ToDataTable(joinedList);

                // Step 4: Convert Contract_expiry_date column
                DataTable tempConverted = original.Clone();
                tempConverted.Columns["Contract_expiry_date"].DataType = typeof(DateTime);

                foreach (DataRow row in original.Rows)
                {
                    var newRow = tempConverted.NewRow();
                    foreach (DataColumn col in original.Columns)
                    {
                        if (col.ColumnName == "Contract_expiry_date")
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
                    tempConverted.Rows.Add(newRow);
                }

                converted = tempConverted;

                // Step 5: Format DateTime columns
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                }

                // Step 6: Bind data to DataGridView
                advancedDataGridView1.DataSource = converted;

                if (!advancedDataGridView1.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn
                    {
                        HeaderText = "تحديد",
                        Name = "Select",
                        Width = 60,
                        ReadOnly = false,
                        TrueValue = true,
                        FalseValue = false
                    };
                    advancedDataGridView1.Columns.Add(checkBoxColumn);
                }
                var getAccesUserFunction = accessTableAdapter.GetDataByOFAccessFunctionUserId(UserId, 4, 5);

                if (getAccesUserFunction.Count == 1)
                {
                    if (!advancedDataGridView1.Columns.Contains("Update"))
                    {
                        DataGridViewButtonColumn btnUpdate = new DataGridViewButtonColumn
                        {
                            HeaderText = "تعديل",
                            Name = "Update",
                            Text = "تعديل",
                            Tag = "Function:Edit",
                            UseColumnTextForButtonValue = true,
                            DisplayIndex = 0,
                            Width = 80
                        };

                        advancedDataGridView1.Columns.Add(btnUpdate);
                    }

                }
                advancedDataGridView1.Columns["Select"].DisplayIndex = 0;
                originalData = converted;
                loadDocs();
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            finally
            {
                FalseFunction();
            }
           
        }


        public void LoadColumnsIntoCheckedListBox() // add from the columns to the list as items
        {
            // Clear previous items
            checkedListBox1.Items.Clear();
            checkedListBox1.Items.Add("اختيار الكل", true);
            // Make sure the DataGridView has a DataSource
            if (advancedDataGridView1.DataSource == null) return;

            // Loop through the columns and add their HeaderText or Name
            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {//test
                string columnName = column.Name?.ToLower() ?? "";
                if (columnName.Contains("file") ||
                    //columnName.Contains("transaction_number")||
                    //columnName.Contains("architectural_and_structural_board") ||
                    columnName.Contains("consultant_surveying") ||
                    //columnName.Contains("consulting_office") ||
                    columnName.Contains("contract_expiry_date")||
                    columnName.Contains("project_id"))
                {
                    checkedListBox1.Items.Remove(column.HeaderText);
                    column.Visible = false; // hide it in the grid
                    continue;
                }
                checkedListBox1.Items.Add(column.HeaderText, column.Visible); // Show as checked if visible
            }
        } 
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        } // not used yet
        // hide and show columns
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                string header = checkedListBox1.Items[index: e.Index].ToString();

                // --- Handle "اختيار الكل" (Select All) ---
                if (header == "اختيار الكل")
                {
                    bool checkAll = e.NewValue == CheckState.Checked;

                    // Apply check/uncheck to all items except the first one
                    for (int i = 1; i < checkedListBox1.Items.Count; i++)
                    {
                        checkedListBox1.SetItemChecked(i, checkAll);
                    }

                    // Update all DataGridView columns visibility
                    foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                    {
                        string columnName = column.Name?.ToLower() ?? "";
                        if (columnName.Contains("file") ||
                            //columnName.Contains("transaction_number") ||
                           //columnName.Contains("architectural_and_structural_board") ||
                            columnName.Contains("consultant_surveying") ||
                            //columnName.Contains("consulting_office") ||
                            columnName.Contains("contract_expiry_date") ||
                            columnName.Contains("project_id"))
                        {
                            continue;
                        }

                            column.Visible = checkAll;
                    }

                    return;
                }

                // --- Handle normal single-item check/uncheck ---
                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.HeaderText == header)
                    {
                        column.Visible = e.NewValue == CheckState.Checked;
                        break;
                    }
                }
            });
        
        } // items check lists if true then apper if false then hide

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
            var documents = this.documentsTableAdapter.GetData();

            int lastDocumentId = 0;

            if (documents != null && documents.Count > 0)
            {
                lastDocumentId = documents
                    .Select(d =>
                    {
                        int parsed;
                        return int.TryParse(d.document_id, out parsed) ? parsed : 0;
                    })
                    .OrderBy(id => id)   // ترتيب كأرقام
                    .Last();             // آخر ID بعد الترتيب
            }
            else
            {
                lastDocumentId = -1; // أو 0 حسب احتياجك
            }

            int idNext = lastDocumentId;
            var row = this.projectsTableAdapter.GetData()
            .OrderByDescending(r => Convert.ToInt32(r.project_id))
            .FirstOrDefault();

            int lastId = row != null ? Convert.ToInt32(row.project_id) : 0;
            int newId = lastId + 1;
            
            var igover = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text);
            var projects = this.projectsTableAdapter.GetData().FindByproject_id(serial_number_TB.Text);
            if (projects != null)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
            }
            bool hasError = false;

            /*if (string.IsNullOrWhiteSpace(serial_number_TB.Text))
            {
                Error_Serial.Visible = true;
                serial_number_TB.Focus();
                hasError = true;
            }*/

            if (string.IsNullOrWhiteSpace(plate_number_COB.Text))
            {
                Error_plate.Visible = true;
                plate_number_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Land_Name_COB.Text))
            {
                Error_Land.Visible = true;
                Land_Name_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Investment_Name_TB.Text))
            {
                Error_Investment_name.Visible = true;
                Investment_Name_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Governorate_COB.Text))
            {
                Error_Governorate.Visible = true;
                Governorate_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Civil_Aviation_COB.Text))
            {
                Error_Civil_Aviation.Visible = true;
                Civil_Aviation_COB.Focus();
                hasError = true;
            }
            if (Civil_Aviation_COB.Text == "✔"&&
                (!approvalFiles.ContainsKey("Civil_Aviation_COB") ||
                approvalFiles["Civil_Aviation_COB"] == null ||
                approvalFiles["Civil_Aviation_COB"].Count == 0))
            {
                Error_Civil_Aviation.Visible = true;
                Civil_Aviation_COB.Focus();
                hasError = true;
            }


            if (string.IsNullOrWhiteSpace(Civil_Defense_COB.Text))
            {
                Error_Civil_Defense.Visible = true;
                Civil_Defense_COB.Focus();
                hasError = true;
            }
            if (Civil_Defense_COB.Text== "✔"&&
                (!approvalFiles.ContainsKey("Civil_Defense_COB") ||
                approvalFiles["Civil_Defense_COB"] == null ||
                approvalFiles["Civil_Defense_COB"].Count == 0))
            {
                Error_Civil_Defense.Visible = true;
                Civil_Defense_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Civil_Study_COB.Text))
            {
                Error_Civil_Study.Visible = true;
                Civil_Study_COB.Focus();
                hasError = true;
            }
            if (Civil_Study_COB.Text== "✔"&&
                (!approvalFiles.ContainsKey("Civil_Study_COB") ||
                approvalFiles["Civil_Study_COB"] == null ||
                approvalFiles["Civil_Study_COB"].Count == 0))
            {
                Error_Civil_Study.Visible = true;
                Civil_Study_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Environmental_COB.Text))
            {
                Error_Environmental.Visible = true;
                Environmental_COB.Focus();
                hasError = true;
            }
            if (Environmental_COB.Text == "✔" &&
                (!approvalFiles.ContainsKey("Environmental_COB") ||
                approvalFiles["Environmental_COB"] == null ||
                approvalFiles["Environmental_COB"].Count == 0))
            {
                Error_Environmental.Visible = true;
                Environmental_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Model_8_COB.Text))
            {
                Error_Model_8.Visible = true;
                Model_8_COB.Focus();
                hasError = true;
            }
            if (Model_8_COB.Text== "✔" &&
                (!approvalFiles.ContainsKey("Model_8_COB") ||
                approvalFiles["Model_8_COB"] == null ||
                approvalFiles["Model_8_COB"].Count == 0))
            {
                Error_Model_8.Visible = true;
                Model_8_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Petroleum_Ministry_COB.Text))
            {
                Error_Petroleum_Ministry.Visible = true;
                Petroleum_Ministry_COB.Focus();
                hasError = true;
            }
            if (Petroleum_Ministry_COB.Text == "✔" &&
                (!approvalFiles.ContainsKey("Petroleum_Ministry_COB") ||
                approvalFiles["Petroleum_Ministry_COB"] == null ||
                approvalFiles["Petroleum_Ministry_COB"].Count == 0))
            {
                Error_Petroleum_Ministry.Visible = true;
                Petroleum_Ministry_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Traffic_Study_COB.Text))
            {
                Error_Traffic_Study.Visible = true;
                Traffic_Study_COB.Focus();
                hasError = true;
            }
            if (Traffic_Study_COB.Text=="✔" &&
                (!approvalFiles.ContainsKey("Traffic_Study_COB") ||
                approvalFiles["Traffic_Study_COB"] == null ||
                approvalFiles["Traffic_Study_COB"].Count == 0))
            {
                Error_Traffic_Study.Visible = true;
                Traffic_Study_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Transaction_number_TB.Text))
            {
                Error_Transaction_number.Visible = true;
                Transaction_number_TB.Focus();
                hasError = true;
            }
            if (!string.IsNullOrWhiteSpace(Transaction_number_TB.Text)&&
                (!approvalFiles.ContainsKey("Transaction_number_TB") ||
                approvalFiles["Transaction_number_TB"] == null ||
                approvalFiles["Transaction_number_TB"].Count == 0))
            {
                Error_Transaction_number.Visible = true;
                Transaction_number_TB.Focus();
                hasError = true;
            }
            if (hasError)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
                return;
            }
            /*if (projects.project_id== serial_number_TB.Text)
            {
                ShowAlert("موجود بل فعل", AlertForm.AlertType.Error);
                return;
            }*/
            this.projectsTableAdapter.Insert(
                (newId).ToString(),
                Investment_Name_TB.Text,
                igover.First().governorate_id,
                Civil_Defense_COB.Text,
                Environmental_COB.Text,
                Traffic_Study_COB.Text,
                Model_8_COB.Text,
                Petroleum_Ministry_COB.Text,
                Civil_Aviation_COB.Text,
                Land_ID_COB.Text,
                Transaction_number_TB.Text,
                guna2DateTimePicker1.Value,
                int.Parse(Total_stores_TB.Text),
                int.Parse(Total_rented_TB.Text),
                int.Parse(Total_Not_rented_TB.Text),
                "x", //Secured_certificate_COM.Text
                "x",// Architectural_and_Structural_Board_COM.Text
                "x",//Reconciliation_Form_Stamp_COM.Text
                "x",//Consultant_Surveying_COM.Text
                Name_Pro_TB.Text,
                Civil_Study_COB.Text
                
                );
            foreach (var entry in approvalFiles)
            {
                string approvalID = entry.Key;
                if (approvalID == "Petroleum_Ministry_COB") { approvalID = "APP003"; }
                if (approvalID == "Civil_Defense_COB") { approvalID = "APP001"; }
                if (approvalID == "Civil_Study_COB") { approvalID = "APP010"; }
                if (approvalID == "Model_8_COB") { approvalID = "APP006"; }
                if (approvalID == "Traffic_Study_COB") { approvalID = "APP005"; }
                if (approvalID == "Civil_Aviation_COB") { approvalID = "APP004"; }
                if (approvalID == "Environmental_COB") { approvalID = "APP002"; }
                if (approvalID == "Transaction_number_TB") { approvalID = "APP007"; }
                List<string> paths = entry.Value;
                foreach (var path in paths)
                {
                    byte[] fileData = File.ReadAllBytes(path);
                    byte[] defaultValue = new byte[] { 0x00 };
                    this.documentsTableAdapter.Insert(
                        (++idNext).ToString(), 
                        approvalID,
                        newId.ToString(),
                        path,
                        defaultValue
                    );
                }
            }
            LoadProjectData();
            ShowAlert("تمت الاضافه بنجاح",AlertForm.AlertType.Success);
            UpdateRowCount();
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();
        }

        List<string> AllfilePaths = new List<string>();
        private Dictionary<string, List<string>> approvalFiles = new Dictionary<string, List<string>>();// at class level
        public void Files(string approvalKey)
        {
            try
            {

          
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All files (*.*)|*.*";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (!approvalFiles.ContainsKey(approvalKey))
                    approvalFiles[approvalKey] = new List<string>();

                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    byte[] fileData = File.ReadAllBytes(filePath);

                    approvalFiles[approvalKey].Add(filePath); // ✅ Add file under approval key

                    AddFileIconToPanel(filePath, fileName, approvalKey); // existing
                }
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void OpenFile(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    ShowAlert("المسار فارغ", AlertForm.AlertType.Error);
                    return;
                }

                string finalPath = filePath;

                // لو الملف مش موجود في المسار الأساسي
                if (!File.Exists(finalPath))
                {
                    // تحويل D:\sho8l -> Z:\
                    if (filePath.StartsWith(@"D:\sho8l\", StringComparison.OrdinalIgnoreCase))
                    {
                        finalPath = @"Z:\" + filePath.Substring(@"D:\sho8l\".Length);
                    }
                    // تحويل Z:\ -> D:\sho8l\
                    else if (filePath.StartsWith(@"Z:\", StringComparison.OrdinalIgnoreCase))
                    {
                        finalPath = @"D:\sho8l\" + filePath.Substring(3);
                    }
                }

                // فتح الملف لو موجود
                if (File.Exists(finalPath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = finalPath,
                        UseShellExecute = true
                    });
                }
                else
                {
                    ShowAlert("الملف غير موجود في أي مسار متاح", AlertForm.AlertType.Error);
                }
            }
            catch (Exception ex)
            {
                ShowAlert("فشل فتح الملف\n" + ex.Message, AlertForm.AlertType.Error);
            }
        }

        private bool IsImageFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif";
        }
        
        private void AddFileIconToPanel(string filePath, string fileName, string namepdf)
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
            switch (namepdf)
            {
                case "Petroleum_Ministry_COB":
                    flowLayoutPanel3.Controls.Add(container);
                    break;
                case "Civil_Defense_COB":
                    flowLayoutPanel2.Controls.Add(container);
                    break;
                case "Civil_Study_COB":
                    flowLayoutPanel9.Controls.Add(container);
                    break;
                case "Model_8_COB":
                    flowLayoutPanel1.Controls.Add(container);
                    break;
                case "Traffic_Study_COB":
                    flowLayoutPanel4.Controls.Add(container);
                    break;
                case "Civil_Aviation_COB":
                    flowLayoutPanel6.Controls.Add(container);
                    break;
                case "Environmental_COB":
                    flowLayoutPanel5.Controls.Add(container);
                    break;
                case "Transaction_number_TB":
                    flowLayoutPanel7.Controls.Add(container);
                    break;
                case "عقود محلات":
                    flowLayoutPanel8.Controls.Add(container);
                    flowLayoutPanel8.FlowDirection = FlowDirection.TopDown;
                    flowLayoutPanel8.WrapContents = false;  // ensures it grows vertically
                    flowLayoutPanel8.AutoScroll = true;
                    AllfilePaths.Add(filePath);
                    break;
                case "كروكيات":
                    flowLayoutPanel8.Controls.Add(container);
                    flowLayoutPanel8.FlowDirection = FlowDirection.TopDown;
                    flowLayoutPanel8.WrapContents = false;  // ensures it grows vertically
                    flowLayoutPanel8.AutoScroll = true;
                    AllfilePaths.Add(filePath);
                    break;
            }

        }
        private void guna2ImageButton6_Click(object sender, EventArgs e)
        {

        }

        private void Transaction_number_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Transaction_number_TB.Text);
            flowLayoutPanel7.Visible = hasText;
            guna2ImageButton4.Visible = hasText;
            guna2ImageButton1.Visible = hasText;
            nightLabel8.Visible = hasText;
            nightLabel7.Visible = hasText;
            Error_Transaction_number.Visible = !hasText;
        }

        private void COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            var comboBox = sender as ComboBox;
            if (comboBox == null) return;

            bool isTrue = comboBox.SelectedItem?.ToString().ToLower() == "✔";

            switch (comboBox.Name)
            {
                case "Petroleum_Ministry_COB":
                    flowLayoutPanel3.Visible = isTrue;
                    guna2ImageButton5.Visible = isTrue;
                    nightLabel10.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Petroleum_Ministry_COB"))  Files("Petroleum_Ministry_COB");
                    break;
                case "Civil_Defense_COB":
                    flowLayoutPanel2.Visible = isTrue;
                    guna2ImageButton3.Visible = isTrue;
                    nightLabel9.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Civil_Defense_COB")) Files("Civil_Defense_COB");
                    break;
                case "Civil_Study_COB":
                    flowLayoutPanel9.Visible = isTrue;
                    guna2ImageButton12.Visible = isTrue;
                    nightLabel15.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Civil_Study_COB")) Files("Civil_Study_COB");
                    break;
                case "Model_8_COB":
                    flowLayoutPanel1.Visible = isTrue;
                    guna2ImageButton2.Visible = isTrue;
                    nightLabel1.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Model_8_COB")) Files("Model_8_COB");
                    break;
                case "Traffic_Study_COB":
                    flowLayoutPanel4.Visible = isTrue;
                    guna2ImageButton7.Visible = isTrue;
                    nightLabel2.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Traffic_Study_COB")) Files("Traffic_Study_COB");
                    break;
                case "Civil_Aviation_COB":
                    flowLayoutPanel6.Visible = isTrue;
                    guna2ImageButton11.Visible = isTrue;
                    nightLabel3.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Civil_Aviation_COB")) Files("Civil_Aviation_COB");
                    break;
                case "Environmental_COB":
                    flowLayoutPanel5.Visible = isTrue;
                    guna2ImageButton9.Visible = isTrue;
                    nightLabel6.Visible = isTrue;
                    if (isTrue && !approvalFiles.ContainsKey("Environmental_COB")) Files("Environmental_COB");
                    break;
                    // Add other cases...
            }
        }

        private void guna2ImageButton5_Click(object sender, EventArgs e)
        {
            flowLayoutPanel3.Controls.Clear();
            if (approvalFiles.ContainsKey("Petroleum_Ministry_COB"))
            {
                flowLayoutPanel3.Controls.Clear();

                approvalFiles["Petroleum_Ministry_COB"] = new List<string>();

                Petroleum_Ministry_COB.Text = "X";
            }
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {

            flowLayoutPanel2.Controls.Clear();
            if (approvalFiles.ContainsKey("Civil_Defense_COB"))
            {
                flowLayoutPanel2.Controls.Clear();

                approvalFiles["Civil_Defense_COB"] = new List<string>();

                Civil_Defense_COB.Text = "X";
            }
        }
        private void guna2ImageButton12_Click(object sender, EventArgs e)
        {
            flowLayoutPanel9.Controls.Clear();
            if (approvalFiles.ContainsKey("Civil_Study_COB"))
            {
                flowLayoutPanel9.Controls.Clear();

                approvalFiles["Civil_Study_COB"] = new List<string>();

                Civil_Study_COB.Text = "X";
            }
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {

            flowLayoutPanel1.Controls.Clear();
            if (approvalFiles.ContainsKey("Model_8_COB"))
            {
                flowLayoutPanel1.Controls.Clear();

                approvalFiles["Model_8_COB"] = new List<string>();

                Model_8_COB.Text = "X";
            }

        }

        private void guna2ImageButton7_Click(object sender, EventArgs e)
        {
            flowLayoutPanel4.Controls.Clear();
            if (approvalFiles.ContainsKey("Traffic_Study_COB"))
            {
                flowLayoutPanel4.Controls.Clear();

                approvalFiles["Traffic_Study_COB"] = new List<string>();

                Traffic_Study_COB.Text = "X";
            }
        }

        private void guna2ImageButton11_Click(object sender, EventArgs e)
        {

            flowLayoutPanel6.Controls.Clear();
            if (approvalFiles.ContainsKey("Civil_Aviation_COB"))
            {
                flowLayoutPanel6.Controls.Clear();

                approvalFiles["Civil_Aviation_COB"] = new List<string>();

                Civil_Aviation_COB.Text = "X";
            }
        }

        private void guna2ImageButton9_Click(object sender, EventArgs e)
        {

            flowLayoutPanel5.Controls.Clear();
            if (approvalFiles.ContainsKey("Environmental_COB"))
            {
                flowLayoutPanel5.Controls.Clear();

                approvalFiles["Environmental_COB"] = new List<string>();

                Environmental_COB.Text = "X";
            }
        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {

            flowLayoutPanel7.Controls.Clear();
            if (approvalFiles.ContainsKey("Transaction_number_TB"))
            {
                flowLayoutPanel7.Controls.Clear();

                approvalFiles["Transaction_number_TB"] = new List<string>();

                Transaction_number_TB.Text = "X";
            }
        }

        private void Land_Name_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            //bool hasText = !string.IsNullOrWhiteSpace(Land_Name_COB.Text);
            //Error_Land.Visible = !hasText;
            //string selectedLandName = Land_Name_COB.SelectedItem?.ToString();
            //if (string.IsNullOrEmpty(selectedLandName)) return;

            //var matches = LandData.AsEnumerable()
            //    .Where(row => row["land_name"].ToString() == selectedLandName)
            //    .ToList();

            //if (matches.Count == 1)
            //{
            //    // Only one match — auto-select
            //    var row = matches[0];
            //    plate_number_COB.SelectedItem = row["plate_number"].ToString();
            //    Land_ID_COB.SelectedItem = row["land_id"].ToString();
            //    string govId = row["governorate_fk"].ToString();
            //    Governorate_COB.SelectedItem = governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";
            //}
            //else if (matches.Count > 1)
            //{
            //    // Multiple matches — populate the related combo boxes
            //    var plateNumbers = matches.Select(r => r["plate_number"].ToString()).Distinct().ToList();
            //    var landIds = matches.Select(r => r["land_id"].ToString()).Distinct().ToList();
            //    var governorates = matches.Select(r =>
            //    {
            //        string govId = r["governorate_fk"].ToString();
            //        return governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";
            //    }).Distinct().ToList();

            //    plate_number_COB.DataSource = new BindingSource(plateNumbers, null);
            //    Land_ID_COB.DataSource = new BindingSource(landIds, null);
            //    Governorate_COB.DataSource = new BindingSource(governorates, null);

            //}
            if (Land_Name_COB.SelectedItem != null)
                SyncSelection(selectedName: Land_Name_COB.SelectedItem.ToString());
        }

            private void plate_number_COB_SelectedIndexChanged(object sender, EventArgs e)
            {
            if (plate_number_COB.SelectedItem != null)
                SyncSelection(selectedPlate: plate_number_COB.SelectedItem.ToString());
            //bool hasText = !string.IsNullOrWhiteSpace(plate_number_COB.Text);
            //string selectedPlate = plate_number_COB.SelectedItem?.ToString();
            //if (string.IsNullOrEmpty(selectedPlate)) return;

            //var match = LandData.AsEnumerable()
            //    .FirstOrDefault(row => row["plate_number"].ToString() == selectedPlate);

            //if (match != null)
            //{
            //    string landId = match["land_id"].ToString();
            //    string landName = match["land_name"].ToString();
            //    string govId = match["governorate_fk"].ToString();
            //    string governorateName = governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";

            //    Land_Name_COB.SelectedItem = landName;
            //    Land_ID_COB.SelectedItem = landId;
            //    Governorate_COB.SelectedItem = governorateName;
            //    AdjustComboBox(Land_ID_COB);
            //    AdjustComboBox(Land_Name_COB);
            //    AdjustComboBox(Governorate_COB);
            //}

            //Error_plate.Visible = !hasText;
        }

        private void ProjectsControl_Leave(object sender, EventArgs e)
        {
            Properties.Settings.Default.LastFilter = advancedDataGridView1.FilterString;
            Properties.Settings.Default.LastSort = advancedDataGridView1.SortString;
            Properties.Settings.Default.Save();
        }

        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            var grid = sender as Zuby.ADGV.AdvancedDataGridView;
            Properties.Settings.Default.LastFilter = (advancedDataGridView1.DataSource as BindingSource)?.Filter;
            Properties.Settings.Default.LastFilterUI = grid.FilterString; // The visible string
            Properties.Settings.Default.Save();
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;

            // Update the row count
            UpdateRowCount();
            if (isRestoring) return;
            SaveState();
        }
        private void UpdateRowCount()
        {
            int count = 0;

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Visible && !row.IsNewRow) // ✅ exclude the new row
                {
                    //var value = row.Cells["Land_Id"].Value?.ToString();
                    count++;
                }
            }
            HashSet<string> uniqueLands = new HashSet<string>();
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    var value = row.Cells["land_fk"].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        uniqueLands.Add(value);
                    }
                }
            }
            int totalStoresSum = 0; // ✅ to store the sum of Total_stores

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    var value = row.Cells["land_fk"].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        uniqueLands.Add(value);
                    }

                    // ✅ Sum Total_stores if numeric
                    //var totalStoresValue = row.Cells["Total_stores"].Value;
                    //if (totalStoresValue != null && int.TryParse(totalStoresValue.ToString(), out int stores))
                    //{
                    //    totalStoresSum += stores;
                    //}
                }
            }
            var projects = projectsTableAdapter.GetData().ToList();
            var investments = investmentsTableAdapter.GetData().ToList();

            // Filter projects with Total_stores > 1
            var storesVsInvestments =
             from p in projects
             where !p.IsTotal_storesNull() && p.Total_stores > 1
             join i in investments on p.land_fk equals i.land_fk into invGroup
             select new
             {
                 ProjectId = p.project_id,
                 project_name = p.project_name,
                 TotalStores = p.Total_stores,
                 ActualInvestments = invGroup.Count()
             };

            int documentsCount = 0;
            
            foreach (DataGridViewRow row in advancedDataGridView2.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    documentsCount++;
                }
            }
            int uniqueProjectsWithDocs = 0;
            HashSet<string> projectIdsWithDocs = new HashSet<string>();

            foreach (DataGridViewRow row in advancedDataGridView2.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    var projectId = row.Cells["Project_Id"].Value?.ToString();
                    if (!string.IsNullOrEmpty(projectId))
                    {
                        projectIdsWithDocs.Add(projectId);
                    }
                }
            }
            int sumTotalStores = storesVsInvestments.Sum(x => x.TotalStores);
            int sumInvestments = storesVsInvestments.Sum(x => x.ActualInvestments);
            uniqueProjectsWithDocs = projectIdsWithDocs.Count;
            rowCountLabel.Text = $"عدد الانشطة: {count}";
            RowCountSeries.Text = $"عدد المشاريع: {uniqueLands.Count}";
            TotalStoresLabel.Text = $"عدد المحلات: {totalStoresSum}";
            TotalStoresLabel2.Text = $"عدد المحلات: {totalStoresSum}";
            InvestmentsCountLabel.Text = $"اجمالي الاستثمارات التفصيلية: {sumInvestments}";
            DocumentsCountLabel.Text = $"عدد المستندات المتاحة: {documentsCount}";
            DocumentsProjectsCountLabel.Text = $"عدد المشاريع التي لديها مستندات محلات: {uniqueProjectsWithDocs}";

            // 🔹 Now add document file counts for approvals
            if (fillterProjects == "")
            {

            
            var projectsWithFiles =
                from DataGridViewRow row in advancedDataGridView1.Rows
                where row.Visible && !row.IsNewRow
                select new
                {
                    CivilDefenseFile = row.Cells["CivilDefenseFile"].Value?.ToString(),
                    CivilDefenseStudyFile = row.Cells["CivilDefenseStudyFile"].Value?.ToString(),
                    EnvironmentalFile = row.Cells["EnvironmentalFile"].Value?.ToString(),
                    PetroleumFile = row.Cells["PetroleumFile"].Value?.ToString(),
                    AviationFile = row.Cells["AviationFile"].Value?.ToString(),
                    TrafficStudyFile = row.Cells["TrafficStudyFile"].Value?.ToString(),
                    Model8File = row.Cells["Model8File"].Value?.ToString(),
                    Transaction_numberFile = row.Cells["Transaction_numberFile"].Value?.ToString()
                };



                int civilDefenseCount = projectsWithFiles.Sum(x => CountFiles(x.CivilDefenseFile));
                int civilDefenseStudyCount = projectsWithFiles.Sum(x => CountFiles(x.CivilDefenseStudyFile));
                int environmentalCount = projectsWithFiles.Sum(x => CountFiles(x.EnvironmentalFile));
                int petroleumCount = projectsWithFiles.Sum(x => CountFiles(x.PetroleumFile));
                int aviationCount = projectsWithFiles.Sum(x => CountFiles(x.AviationFile));
                int trafficStudyCount = projectsWithFiles.Sum(x => CountFiles(x.TrafficStudyFile));
                int model8Count = projectsWithFiles.Sum(x => CountFiles(x.Model8File));
                int Transaction_numberCount = projectsWithFiles.Sum(x => CountFiles(x.Transaction_numberFile));

                // 🔹 Update your 6 labels (replace with your actual label names)
                dreamButton1.Text = $"موافقة الحماية المدنية";
            dreamButton9.Text = $"{ civilDefenseCount}"; // 
            dreamButton3.Text = $"موافقة البيئة";
            dreamButton10.Text = $"{environmentalCount}";
            dreamButton4.Text = $"موافقة البترول";
            dreamButton11.Text = $"{petroleumCount}";
            dreamButton2.Text = $"موافقة الطيران المدني";
            dreamButton8.Text = $"{aviationCount}";
            dreamButton5.Text = $"الدراسات المرورية";
            dreamButton12.Text = $"{trafficStudyCount}";
            dreamButton6.Text = $"نموذج 8 أو 10";
            dreamButton7.Text = $"{model8Count}";
            dreamButton16.Text = $"رخصة التشغيل";
            dreamButton15.Text = $"{Transaction_numberCount}";
            dreamButton28.Text = $"دراسة الحماية المدنية";
            dreamButton27.Text = $"{civilDefenseStudyCount}";

            }
        }
        private int CountFiles(string files)
        {
            if (string.IsNullOrWhiteSpace(files))
                return 0;

            return files
                .Split(new[] { " , " }, StringSplitOptions.RemoveEmptyEntries)
                .Length;
        }
        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            var grid = sender as Zuby.ADGV.AdvancedDataGridView;
            Properties.Settings.Default.LastSort = grid.SortString;
            Properties.Settings.Default.Save();
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
            if (isRestoring) return;

            SaveState();
        }

        private void guna2ImageButton4_Click(object sender, EventArgs e)
        {
            Files("Transaction_number_TB");
        }
        private void ShowFileListWindow(DataTable docs)
        {
            Form f = new Form();
            f.Text = "اختر ملفًا لفتحه";
            f.StartPosition = FormStartPosition.CenterParent;
            f.Size = new Size(500, 300);
            f.FormBorderStyle = FormBorderStyle.FixedDialog;

            ListBox list = new ListBox();
            list.Dock = DockStyle.Fill;
            list.Font = new System.Drawing.Font("Segoe UI", 11);
            list.DrawMode = DrawMode.OwnerDrawFixed;
            list.ItemHeight = 30;

            // store all info in dictionary
            Dictionary<string, DocInfo> fileData = new Dictionary<string, DocInfo>();

            foreach (DataRow row in docs.Rows)
            {
                string fullPath = row["paths"].ToString();
                int docID = Convert.ToInt32(row["document_id"]);
                string name = Path.GetFileName(fullPath);

                list.Items.Add(name);

                fileData[name] = new DocInfo
                {
                    FileName = name,
                    FullPath = fullPath,
                    DocumentID = docID
                };
            }

            // Draw filename + [حذف]
            list.DrawItem += (s, e) =>
            {
                e.DrawBackground();

                if (e.Index >= 0)
                {
                    string name = list.Items[e.Index].ToString();

                    // draw filename
                    e.Graphics.DrawString(name, list.Font,
                        Brushes.Black, e.Bounds.Left + 5, e.Bounds.Top + 5);

                    // draw delete text
                    string deleteText = "[حذف]";
                    SizeF size = e.Graphics.MeasureString(deleteText, list.Font);

                    float x = e.Bounds.Right - size.Width - 10;
                    float y = e.Bounds.Top + 5;

                    e.Graphics.DrawString(deleteText, list.Font, Brushes.Red, x, y);
                }

                e.DrawFocusRectangle();
            };

            // Handle clicks (detect delete area)
            list.MouseClick += (s, e) =>
            {
                int index = list.IndexFromPoint(e.Location);
                if (index < 0) return;

                string name = list.Items[index].ToString();
                DocInfo info = fileData[name];

                string deleteText = "[حذف]";

                System.Drawing.Rectangle itemRect = list.GetItemRectangle(index);
                SizeF textSize = list.CreateGraphics().MeasureString(deleteText, list.Font);

                System.Drawing.Rectangle deleteRect = new System.Drawing.Rectangle(
                    itemRect.Right - (int)textSize.Width - 10,
                    itemRect.Top,
                    (int)textSize.Width + 10,
                    itemRect.Height
                );

                // ✔ Delete clicked
                if (deleteRect.Contains(e.Location))
                {
                    if (MessageBox.Show("هل تريد حذف المستند من القاعدة والملفات؟",
                                        "تأكيد", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        try
                        {
                            // 1) Delete from database using TableAdapter
                            this.documentsTableAdapter.DeleteQuery1(info.DocumentID.ToString());

                            // 2) Delete physical file
                            if (File.Exists(info.FullPath))
                                File.Delete(info.FullPath);

                            // 3) Remove from ListBox
                            list.Items.RemoveAt(index);
                            fileData.Remove(name);

                            MessageBox.Show("تم حذف المستند بنجاح.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("خطأ أثناء الحذف: " + ex.Message);
                        }
                    }

                    return;
                }

                // ✔ Otherwise: open file
                OpenFile(info.FullPath);
                f.Close();
            };

            f.Controls.Add(list);
            f.ShowDialog();
        }



        private class FileItem
        {
            public string FileName { get; set; }
            public string FullPath { get; set; }

            public override string ToString() => FileName;
        }

        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var clickedColumn = advancedDataGridView1.Columns[e.ColumnIndex];

            // Check if Status or Transaction column was clicke
            if (clickedColumn.Name.Contains("status") ||
                clickedColumn.Name.Contains("Status") ||
                clickedColumn.Name == "Transaction_number")
            {
                string filePath = "";

                // pick correct file cell
                switch (clickedColumn.Name)
                {
                    case "Civil_Defense_Approval_status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["CivilDefenseFile"].Value?.ToString();
                        break;

                    case "Civil_Defense_Study_status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["CivilDefenseStudyFile"].Value?.ToString();
                        break;

                    case "Environmental_Approval_status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["EnvironmentalFile"].Value?.ToString();
                        break;

                    case "Petroleum_Ministry_Approval_status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["PetroleumFile"].Value?.ToString();
                        break;

                    case "Civil_Aviation_Approval_status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["AviationFile"].Value?.ToString();
                        break;

                    case "Traffic_Study_Status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["TrafficStudyFile"].Value?.ToString();
                        break;

                    case "Model_8_Status":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["Model8File"].Value?.ToString();
                        break;

                    case "Transaction_number":
                        filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["Transaction_numberFile"].Value?.ToString();
                        break;
                }

                if (string.IsNullOrWhiteSpace(filePath))
                {
                    MessageBox.Show("الملف غير موجود أو المسار فارغ.");
                    return;
                }

                // Split multiple files
                string[] files = filePath
                    .Split(new char[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(f => f.Trim())
                    .ToArray();

                // if only 1 file → open directly
                if (files.Length == 1)
                {
                    OpenFile(files[0]);
                    return;
                }

                // MULTIPLE FILES → load matching rows from DB
                List<string> ids = new List<string>();

                // Loop database table
                var docs = documentsTableAdapter.GetData();
                foreach (var file in files)
                {
                    string cleaned = file.Replace("\\", "/").ToLower();

                    foreach (DataRow row in docs.Rows)
                    {
                        string dbPath = row["paths"].ToString().Replace("\\", "/").ToLower();

                        if (dbPath == cleaned)
                        {
                            ids.Add(row["document_id"].ToString());
                        }
                    }
                }

                // Filter table with these IDs
                DataTable filtered = docs.AsEnumerable()
                    .Where(r => ids.Contains(r.Field<string>("document_id")) &&
                     r.projects_fk == advancedDataGridView1.Rows[e.RowIndex].Cells["Project_Id"].Value?.ToString())
                    .CopyToDataTable();

                // Show popup window with correct docs only
                ShowFileListWindow(filtered);
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "Update"
             && e.RowIndex >= 0)
            {
                try
                {
                    Update_Radio.Checked = true;
                    string serial = advancedDataGridView1
                        .Rows[e.RowIndex]
                        .Cells["Project_Id"]   // make sure column name matches exactly
                        .Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(serial))
                    {
                        // Load data into controls
                        LoadProjectData(serial);

                        // Go to TabPage 4
                        guna2TabControl1.SelectedTab = tabPage4;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                if (Update_Radio.Checked)
                {
                    try
                    {
                        if (e.RowIndex >= 0)
                        {
                            string serial = advancedDataGridView1.Rows[e.RowIndex].Cells["project_id"].Value.ToString();
                            serial_number_TB.Text = serial;
                            guna2TextBox1.Text = serial;
                            var x = this.projectsTableAdapter.GetDataByIDProjects(serial);
                            guna2TextBox3.Text = x.First().Name_Projects;
                            LoadProjectData(serial);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
               
                    if (e.RowIndex >= 0)
                    {
                        string serial = advancedDataGridView1.Rows[e.RowIndex].Cells["project_id"].Value.ToString();
                        guna2TextBox1.Text = serial;
                        var x = this.projectsTableAdapter.GetDataByIDProjects(serial);
                        guna2TextBox3.Text = x.First().Name_Projects;

                    }
                
            }
        }
        public class DocInfo
        {
            public string FileName { get; set; }
            public string FullPath { get; set; }
            public int DocumentID { get; set; }
        }
        private void LoadProjectData(string serialNumber)
        {
            bool hasText = !string.IsNullOrWhiteSpace(serialNumber);
            Error_Serial.Visible = !hasText;

            var x = this.projectsTableAdapter.GetDataByIDProjects(serialNumber);

            var gCivil_Defense = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP001");
            var gCivil_Study = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP010");
            var gEnvironmental = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP002");
            var gPetroleum_Ministry = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP003");
            var gCivil_Aviation = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP004");
            var gTraffic_Study = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP005");
            var gModel_8 = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP006");
            var gTransaction_number = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP007");

            if (x == null || x.Count == 0)
            {
                Investment_Name_TB.Text = "";
                Transaction_number_TB.Text = "";

                Governorate_COB.SelectedIndex = 0;
                Land_Name_COB.SelectedIndex = 0;
                Land_ID_COB.SelectedIndex = 0;
                plate_number_COB.SelectedIndex = 0;

                Traffic_Study_COB.SelectedIndex = -1;
                Petroleum_Ministry_COB.SelectedIndex = -1;
                Environmental_COB.SelectedIndex = -1;
                Model_8_COB.SelectedIndex = -1;
                Civil_Aviation_COB.SelectedIndex = -1;
                Civil_Defense_COB.SelectedIndex = -1;
                Civil_Study_COB.SelectedIndex = -1;

                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel2.Controls.Clear();
                flowLayoutPanel3.Controls.Clear();
                flowLayoutPanel4.Controls.Clear();
                flowLayoutPanel5.Controls.Clear();
                flowLayoutPanel6.Controls.Clear();
                flowLayoutPanel7.Controls.Clear();
                flowLayoutPanel9.Controls.Clear();
                return;
            }

            var l = this.landsTableAdapter.GetDataByLandId(x.First().land_fk);
            Investment_Name_TB.Text = x.First().project_name;
            guna2TextBox3.Text = x.First().Name_Projects;
            Name_Pro_TB.Text = x.First().Name_Projects;
            Transaction_number_TB.Text = x.First().Transaction_number;
            Total_stores_TB.Text = x.First().Total_stores.ToString();
            Total_rented_TB.Text = x.First().Total_rented.ToString();
            Total_Not_rented_TB.Text = x.First().Total_Not_rented.ToString();
            Secured_certificate_COM.SelectedItem = x.First().Secured_certificate;
            Architectural_and_Structural_Board_COM.SelectedItem = x.First().Architectural_and_Structural_Board;
            Consultant_Surveying_COM.SelectedItem = x.First().Consultant_Surveying;
            Reconciliation_Form_Stamp_COM.SelectedItem = x.First().Reconciliation_Form_Stamp;
            Governorate_COB.SelectedItem = governorateLookup.TryGetValue(
                x.First().governorate_fk.ToString(), out var govName
            ) ? govName : "غير معروف";

            if (x.First().land_fk == l.First().land_id)
            {
                Land_Name_COB.SelectedItem = l.First().land_name;
                plate_number_COB.SelectedItem = l.First().plate_number;
                AdjustComboBox(Land_Name_COB);
                AdjustComboBox(plate_number_COB);
            }

            Land_ID_COB.SelectedItem = x.First().land_fk;

            // نفس الجزء بتاع تحميل الملفات
            string filePathEnvironmental = gEnvironmental.FirstOrDefault()?.paths?.ToString();
            string filePathCivil_Defense = gCivil_Defense.FirstOrDefault()?.paths?.ToString();
            string filePathCivil_Study = gCivil_Study.FirstOrDefault()?.paths?.ToString();
            string filePathPetroleum_Ministry = gPetroleum_Ministry.FirstOrDefault()?.paths?.ToString();
            string filePathCivil_Aviation = gCivil_Aviation.FirstOrDefault()?.paths?.ToString();
            string filePathTraffic_Study = gTraffic_Study.FirstOrDefault()?.paths?.ToString();
            string filePathModel_8 = gModel_8.FirstOrDefault()?.paths?.ToString();
            string filePathTransaction_number = gTransaction_number.FirstOrDefault()?.paths?.ToString();

            if (!string.IsNullOrWhiteSpace(filePathEnvironmental))
            {
                string FileName = Path.GetFileName(filePathEnvironmental);
                if (!approvalFiles.ContainsKey("Environmental_COB"))
                {
                    approvalFiles["Environmental_COB"] = new List<string>();
                    approvalFiles["Environmental_COB"].Add(filePathEnvironmental);
                }
                AddFileIconToPanel(filePathEnvironmental, FileName, "Environmental_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathPetroleum_Ministry))
            {
                string FileName = Path.GetFileName(filePathPetroleum_Ministry);
                if (!approvalFiles.ContainsKey("Petroleum_Ministry_COB"))
                {
                    approvalFiles["Petroleum_Ministry_COB"] = new List<string>();
                    approvalFiles["Petroleum_Ministry_COB"].Add(filePathPetroleum_Ministry);
                }
                AddFileIconToPanel(filePathPetroleum_Ministry, FileName, "Petroleum_Ministry_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathCivil_Defense))
            {
                string FileName = Path.GetFileName(filePathCivil_Defense);
                if (!approvalFiles.ContainsKey("Civil_Defense_COB")
)
                {
                    approvalFiles["Civil_Defense_COB"] = new List<string>();
                    approvalFiles["Civil_Defense_COB"].Add(filePathCivil_Defense);
                }
                AddFileIconToPanel(filePathCivil_Defense, FileName, "Civil_Defense_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathCivil_Study))
            {
                string FileName = Path.GetFileName(filePathCivil_Study);
                if (!approvalFiles.ContainsKey("Civil_Study_COB"))
                {
                    approvalFiles["Civil_Study_COB"] = new List<string>();
                    approvalFiles["Civil_Study_COB"].Add(filePathCivil_Study);
                }
                AddFileIconToPanel(filePathCivil_Study, FileName, "Civil_Study_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathCivil_Aviation))
            {
                string FileName = Path.GetFileName(filePathCivil_Aviation);
                if (!approvalFiles.ContainsKey("Civil_Aviation_COB"))
                {
                    approvalFiles["Civil_Aviation_COB"] = new List<string>();
                    approvalFiles["Civil_Aviation_COB"].Add(filePathCivil_Aviation);
                }
                AddFileIconToPanel(filePathCivil_Aviation, FileName, "Civil_Aviation_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathTraffic_Study))
            {
                string FileName = Path.GetFileName(filePathTraffic_Study);
                if (!approvalFiles.ContainsKey("Traffic_Study_COB"))
                {
                    approvalFiles["Traffic_Study_COB"] = new List<string>();
                    approvalFiles["Traffic_Study_COB"].Add(filePathTraffic_Study);
                }
                AddFileIconToPanel(filePathTraffic_Study, FileName, "Traffic_Study_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathModel_8))
            {
                string FileName = Path.GetFileName(filePathModel_8);
                if (!approvalFiles.ContainsKey("Model_8_COB"))
                {
                    approvalFiles["Model_8_COB"] = new List<string>();
                    approvalFiles["Model_8_COB"].Add(filePathModel_8);
                }
                AddFileIconToPanel(filePathModel_8, FileName, "Model_8_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathTransaction_number))
            {
                string FileName = Path.GetFileName(filePathTransaction_number);
                if (!approvalFiles.ContainsKey("Transaction_number_TB"))
                {
                    approvalFiles["Transaction_number_TB"] = new List<string>();
                    approvalFiles["Transaction_number_TB"].Add(filePathTransaction_number);
                }
                AddFileIconToPanel(filePathTransaction_number, FileName, "Transaction_number_TB");
            }

            Traffic_Study_COB.SelectedItem = x.First().Traffic_Study_Status.ToString() == "X" ||
                x.First().Traffic_Study_Status.ToString() == "x" ? "X" : "✔";
            Petroleum_Ministry_COB.SelectedItem = x.First().Petroleum_Ministry_Approval_status.ToString() == "X" ||
                x.First().Petroleum_Ministry_Approval_status.ToString() == "x" ? "X" : "✔";
            Environmental_COB.SelectedItem = x.First().Environmental_Approval_status.ToString() == "X" ||
                x.First().Environmental_Approval_status.ToString() == "x" ? "X" : "✔";
            Model_8_COB.SelectedItem = x.First().Model_8_Status.ToString() == "X" ||
                x.First().Model_8_Status.ToString() == "x" ? "X" : "✔";
            Civil_Aviation_COB.SelectedItem = x.First().Civil_Aviation_Approval_status.ToString() == "X" ||
                x.First().Civil_Aviation_Approval_status.ToString() == "x" ? "X" : "✔";
            Civil_Defense_COB.SelectedItem = x.First().Civil_Defense_Approval_status.ToString() == "X" ||
                x.First().Civil_Defense_Approval_status.ToString() =="x"? "X" : "✔";
            Civil_Study_COB.SelectedItem = x.First().Civil_Defense_Study_status.ToString() == "X" ||
                x.First().Civil_Defense_Study_status.ToString() =="x"? "X" : "✔";
        }

        // TextChanged event
     
        // CellClick event
    


        private void serial_number_TB_TextChanged(object sender, EventArgs e)
        {
            //bool hasText = !string.IsNullOrWhiteSpace(serial_number_TB.Text);
            //Error_Serial.Visible = !hasText;
            //var x = this.projectsTableAdapter.GetDataByIDProjects(serial_number_TB.Text);
            
            //var gCivil_Defense = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP001");
            //var gEnvironmental = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP002");
            //var gPetroleum_Ministry = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP003");
            //var gCivil_Aviation = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP004");
            //var gTraffic_Study = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP005");
            //var gModel_8 = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP006");
            //var gTransaction_number = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, "APP007");
            //if (x == null || x.Count == 0)
            //{
            //    Investment_Name_TB.Text = "";
            //    Transaction_number_TB.Text = "";

            //    Governorate_COB.SelectedIndex = 0;
            //    Land_Name_COB.SelectedIndex = 0;
            //    Land_ID_COB.SelectedIndex = 0;
            //    plate_number_COB.SelectedIndex = 0;

            //    Traffic_Study_COB.SelectedIndex = -1;
            //    Petroleum_Ministry_COB.SelectedIndex = -1;
            //    Environmental_COB.SelectedIndex = -1;
            //    Model_8_COB.SelectedIndex = -1;
            //    Civil_Aviation_COB.SelectedIndex = -1;
            //    Civil_Defense_COB.SelectedIndex = -1;

            //    flowLayoutPanel1.Controls.Clear();
            //    flowLayoutPanel2.Controls.Clear();
            //    flowLayoutPanel3.Controls.Clear();
            //    flowLayoutPanel4.Controls.Clear();
            //    flowLayoutPanel5.Controls.Clear();
            //    flowLayoutPanel6.Controls.Clear();
            //    flowLayoutPanel7.Controls.Clear();
            //    return;
            //}
            //var l = this.landsTableAdapter.GetDataBySerial(x.First().land_fk);
            //Investment_Name_TB.Text = x.First().project_name;
            //Transaction_number_TB.Text = x.First().Transaction_number;
            //Governorate_COB.SelectedItem = governorateLookup.TryGetValue
            //(x.First().governorate_fk.ToString(), out var govName)
            //? govName
            //: "غير معروف";
            //if (x.First().land_fk == l.First().land_id)
            //{
            //    Land_Name_COB.SelectedItem = l.First().land_name;
            //    plate_number_COB.SelectedItem = l.First().plate_number;
            //    AdjustComboBox(Land_Name_COB);
            //    AdjustComboBox(plate_number_COB);

            //}
            //Land_ID_COB.SelectedItem = x.First().land_fk;
            //string filePathEnvironmental = gEnvironmental.First().paths?.ToString();
            //string filePathCivil_Defense = gCivil_Defense.First().paths?.ToString();
            //string filePathPetroleum_Ministry = gPetroleum_Ministry.First().paths?.ToString();
            //string filePathCivil_Aviation = gCivil_Aviation.First().paths?.ToString();
            //string filePathTraffic_Study = gTraffic_Study.First().paths?.ToString();
            //string filePathModel_8 = gModel_8.First().paths?.ToString();
            //var firstItem = gTransaction_number.FirstOrDefault();
            //string filePathTransaction_number = firstItem == null
            //    ? null
            //    : firstItem.paths?.ToString();
            //if (!string.IsNullOrWhiteSpace(filePathEnvironmental))
            //{
            //    string FileName = Path.GetFileName(filePathEnvironmental);

            //    if (!approvalFiles.ContainsKey(filePathEnvironmental))
            //    {
            //        approvalFiles["Environmental_COB"] = new List<string>();
            //        approvalFiles["Environmental_COB"].Add(filePathEnvironmental);
            //    }
            //    AddFileIconToPanel(filePathEnvironmental, FileName, "Environmental_COB");
            //}
            //if (!string.IsNullOrWhiteSpace(filePathPetroleum_Ministry))
            //{
            //    string FileName = Path.GetFileName(filePathPetroleum_Ministry);

            //    if (!approvalFiles.ContainsKey(filePathPetroleum_Ministry))
            //    {
            //        approvalFiles["Petroleum_Ministry_COB"] = new List<string>();
            //        approvalFiles["Petroleum_Ministry_COB"].Add(filePathPetroleum_Ministry);

            //    }
            //    AddFileIconToPanel(filePathPetroleum_Ministry, FileName, "Petroleum_Ministry_COB");
            //}
            //if (!string.IsNullOrWhiteSpace(filePathCivil_Defense))
            //{
            //    string FileName = Path.GetFileName(filePathCivil_Defense);

            //    if (!approvalFiles.ContainsKey(filePathCivil_Defense))
            //    {
            //        approvalFiles["Civil_Defense_COB"] = new List<string>();
            //        approvalFiles["Civil_Defense_COB"].Add(filePathCivil_Defense);

            //    }
            //    AddFileIconToPanel(filePathCivil_Defense, FileName, "Civil_Defense_COB");
            //}
            //if (!string.IsNullOrWhiteSpace(filePathCivil_Aviation))
            //{
            //    string FileName = Path.GetFileName(filePathCivil_Aviation);

            //    if (!approvalFiles.ContainsKey(filePathCivil_Aviation))
            //    {

            //        approvalFiles["Civil_Aviation_COB"] = new List<string>();
            //        approvalFiles["Civil_Aviation_COB"].Add(filePathCivil_Aviation);
            //    }
            //    AddFileIconToPanel(filePathCivil_Aviation, FileName, "Civil_Aviation_COB");
            //}
            //if (!string.IsNullOrWhiteSpace(filePathTraffic_Study))
            //{
            //    string FileName = Path.GetFileName(filePathTraffic_Study);

            //    if (!approvalFiles.ContainsKey(filePathTraffic_Study))
            //    {
            //        approvalFiles["Traffic_Study_COB"] = new List<string>();
            //        approvalFiles["Traffic_Study_COB"].Add(filePathTraffic_Study);
            //    }
            //    AddFileIconToPanel(filePathTraffic_Study, FileName, "Traffic_Study_COB");
            //}
            //if (!string.IsNullOrWhiteSpace(filePathModel_8))
            //{
            //    string FileName = Path.GetFileName(filePathModel_8);

            //    if (!approvalFiles.ContainsKey(filePathModel_8))
            //    {
            //        approvalFiles["Model_8_COB"] = new List<string>();
            //        approvalFiles["Model_8_COB"].Add(filePathEnvironmental);
            //    }
            //    AddFileIconToPanel(filePathModel_8, FileName, "Model_8_COB");
            //}
            //if (!string.IsNullOrWhiteSpace(filePathTransaction_number))
            //{
            //    string FileName = Path.GetFileName(filePathTransaction_number);

            //    if (!approvalFiles.ContainsKey(filePathTransaction_number))
            //    {
            //        approvalFiles["Transaction_number_TB"] = new List<string>();
            //        approvalFiles["Transaction_number_TB"].Add(filePathTransaction_number);
            //    }
            //    AddFileIconToPanel(filePathTransaction_number, FileName, "Transaction_number_TB");
            //}
            //Traffic_Study_COB.SelectedItem = x.First().Traffic_Study_Status.ToString() == "X" ? "X" : "✔";
            //Petroleum_Ministry_COB.SelectedItem = x.First().Petroleum_Ministry_Approval_status.ToString() == "X" ? "X" : "✔";
            //Environmental_COB.SelectedItem = x.First().Environmental_Approval_status.ToString() == "X" ? "X" : "✔";
            //Model_8_COB.SelectedItem = x.First().Model_8_Status.ToString() == "X" ? "X" : "✔";
            //Civil_Aviation_COB.SelectedItem = x.First().Civil_Aviation_Approval_status.ToString() == "X" ? "X" : "✔";
            //Civil_Defense_COB.SelectedItem = x.First().Civil_Defense_Approval_status.ToString() == "X" ? "X" : "✔";

           
        }

        private void Investment_Name_TB_TextChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Investment_Name_TB.Text);
            Error_Investment_name.Visible = !hasText;
        }

        private void Governorate_COB_SelectedIndexChanged(object sender, EventArgs e)
        {

            bool hasText = !string.IsNullOrWhiteSpace(Governorate_COB.Text);
            Error_Governorate.Visible = !hasText;
        }

        private void Petroleum_Ministry_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Petroleum_Ministry_COB.Text);
            Error_Petroleum_Ministry.Visible = !hasText;
        }

        private void Traffic_Study_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Traffic_Study_COB.Text);
            Error_Traffic_Study.Visible = !hasText;
        }

        private void Civil_Defense_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Civil_Defense_COB.Text);
            Error_Civil_Defense.Visible = !hasText;
        
        }

        private void Civil_Study_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Civil_Study_COB.Text);
            Error_Civil_Study.Visible = !hasText;
        }

        private void Civil_Aviation_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Civil_Aviation_COB.Text);
            Error_Civil_Aviation.Visible = !hasText;
        }

        private void Model_8_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Model_8_COB.Text);
            Error_Model_8.Visible = !hasText;
        }

        private void Environmental_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool hasText = !string.IsNullOrWhiteSpace(Environmental_COB.Text);
            Error_Environmental.Visible = !hasText;
        }

        private void Transaction_number_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }

        private void serial_number_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }
        private void SyncSelection(string selectedLandId = null, string selectedPlate = null, string selectedName = null)
        {
            try
            {

           
            // نحدد أي Row محتاجينه
            var query = LandData.AsEnumerable();

            if (!string.IsNullOrEmpty(selectedLandId))
                query = query.Where(r => r["land_id"].ToString() == selectedLandId);

            if (!string.IsNullOrEmpty(selectedPlate))
                query = query.Where(r => r["plate_number"].ToString() == selectedPlate);

            if (!string.IsNullOrEmpty(selectedName))
                query = query.Where(r => r["land_name"].ToString() == selectedName);

            var matches = query.ToList();
            if (!matches.Any()) return;

            // لو فيه match وحيد → نختاره كله
            if (matches.Count == 1)
            {
                var row = matches[0];
                Land_ID_COB.SelectedItem = row["land_id"].ToString();
                plate_number_COB.SelectedItem = row["plate_number"].ToString();
                Land_Name_COB.SelectedItem = row["land_name"].ToString();

                string govId = row["governorate_fk"].ToString();
                Governorate_COB.SelectedItem = governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";
            }
            else
            {
                // أكتر من احتمال:
                // - نخلي الـ Plate ثابت لأنه مشترك
                // - نخلي الـ ID و Name يفضلوا مفتوحين للمستخدم

                if (!string.IsNullOrEmpty(selectedPlate))
                    plate_number_COB.SelectedItem = selectedPlate;

                if (!string.IsNullOrEmpty(selectedLandId))
                    Land_ID_COB.SelectedItem = selectedLandId;
                else
                    Land_ID_COB.SelectedItem = matches[0]["land_id"].ToString();

                if (!string.IsNullOrEmpty(selectedName))
                    Land_Name_COB.SelectedItem = selectedName;
                else
                    Land_Name_COB.SelectedItem = matches[0]["land_name"].ToString();
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,"Error");
            }
        }
        private void Land_ID_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Land_ID_COB.SelectedItem != null)
                SyncSelection(selectedLandId: Land_ID_COB.SelectedItem.ToString());
            //string selectedLandID = Land_ID_COB.SelectedItem?.ToString();
            //if (string.IsNullOrEmpty(selectedLandID)) return;

            //var match = LandData.AsEnumerable()
            //    .FirstOrDefault(row => row["land_id"].ToString() == selectedLandID);

            //if (match != null)
            //{
            //    string plate = match["plate_number"].ToString();
            //    string name = match["land_name"].ToString();
            //    string govId = match["governorate_fk"].ToString();
            //    string governorateName = governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";

            //    // Set selections only
            //    plate_number_COB.SelectedItem = plate;
            //    Land_Name_COB.SelectedItem = name;
            //    Governorate_COB.SelectedItem = governorateName;
            //    AdjustComboBox(Land_Name_COB);
            //    AdjustComboBox(plate_number_COB); 
            //    AdjustComboBox(Governorate_COB);
            //}
        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            fillterProjects = "";
            advancedDataGridView1.CleanSort();
            advancedDataGridView1.CleanFilter();
            advancedDataGridView1.DataSource = null;
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            LoadProjectData();
            LoadNotes();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();
            tabPage4.Text = "";

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            // ------------------------------------------------------
            // 1) User Confirmation
            // ------------------------------------------------------
            DialogResult result = MessageBox.Show(
                "هل أنت متأكد من أنك تريد تعديل هذه البيانات؟",
                "تأكيد التعديل",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            // ------------------------------------------------------
            // 2) Validate Project Exists
            // ------------------------------------------------------
            var projectData = this.projectsTableAdapter.GetDataByIDProjects(serial_number_TB.Text);
            if (projectData == null || projectData.Count == 0)
            {
                ShowAlert("لا يوجد هذا البيان للتعديل", AlertForm.AlertType.Error);
                return;
            }

            // ------------------------------------------------------
            // 3) Fetch Governorate ID
            // ------------------------------------------------------
            var igover = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text);

            // ------------------------------------------------------
            // 4) Update Project Information
            // ------------------------------------------------------

            this.projectsTableAdapter.UpdateQuery(
                Investment_Name_TB.Text,
                igover.First().governorate_id,
                Civil_Defense_COB.Text,
                Environmental_COB.Text,
                Traffic_Study_COB.Text,
                Model_8_COB.Text,
                Petroleum_Ministry_COB.Text,
                Civil_Aviation_COB.Text,
                Land_ID_COB.Text,
                Transaction_number_TB.Text,
                guna2DateTimePicker1.Value.ToString(),
                int.Parse(Total_stores_TB.Text),
                int.Parse(Total_rented_TB.Text),
                int.Parse(Total_Not_rented_TB.Text),
                Secured_certificate_COM.Text,
                Architectural_and_Structural_Board_COM.Text,
                Reconciliation_Form_Stamp_COM.Text,
                Consultant_Surveying_COM.Text,
                Name_Pro_TB.Text,
                Civil_Study_COB.Text,
                serial_number_TB.Text

            );

            // ------------------------------------------------------
            // 5) Approval Mapping Table
            // ------------------------------------------------------
            Dictionary<string, string> approvalMap = new Dictionary<string, string>()
    {
        { "Petroleum_Ministry_COB", "APP003" },
        { "Civil_Defense_COB",      "APP001" },
        { "Civil_Study_COB",      "APP010" },
        { "Model_8_COB",            "APP006" },
        { "Traffic_Study_COB",      "APP005" },
        { "Civil_Aviation_COB",     "APP004" },
        { "Environmental_COB",      "APP002" },
        { "Transaction_number_TB",  "APP007" }
    };

            // ------------------------------------------------------
            // 6) For Each Approval → Sync Documents
            // ------------------------------------------------------

            foreach (var entry in approvalFiles)
            {
                string approvalKey = entry.Key;
                List<string> paths = entry.Value;

                if (!approvalMap.ContainsKey(approvalKey))
                    continue;

                string approvalID = approvalMap[approvalKey];

                // Load existing documents for this approval/project
                var existingDocs = this.documentsTableAdapter.GetDataByProjectDoc(serial_number_TB.Text, approvalID);
                int existingCount = existingDocs.Count;
                int selectedCount = paths.Count;

                // Normalize IDs
                var allDocuments = this.documentsTableAdapter.GetData();
                int lastDocumentId = allDocuments
                    .Select(d => int.TryParse(d.document_id, out int parsed) ? parsed : 0)
                    .DefaultIfEmpty(0)
                    .Max();

                int updateCount = Math.Min(existingCount, selectedCount);

                // ------------------------------------------------------
                // 6A — Update existing records
                // ------------------------------------------------------
                for (int i = 0; i < updateCount; i++)
                {
                    string path = paths[i];
                    byte[] defaultValue = new byte[] { 0x00 };

                    this.documentsTableAdapter.UpdateDocumentPath(
                        path,
                        defaultValue,
                        existingDocs[i].document_id
                    );
                }

                // ------------------------------------------------------
                // 6B — Insert new records (if more files selected)
                // ------------------------------------------------------
                for (int i = existingCount; i < selectedCount; i++)
                {
                    string path = paths[i];
                    byte[] defaultValue = new byte[] { 0x00 };

                    lastDocumentId++;

                    this.documentsTableAdapter.Insert(
                        lastDocumentId.ToString(),
                        approvalID,
                        serial_number_TB.Text,
                        path,
                        defaultValue
                    );
                }

                // ------------------------------------------------------
                // 6C — Remove extra previous documents (if fewer selected)
                // ------------------------------------------------------
                for (int i = selectedCount; i < existingCount; i++)
                {
                    this.documentsTableAdapter.DeleteQuery1(existingDocs[i].document_id);
                }
            }

            // ------------------------------------------------------
            // 7) Reload and Cleanup
            // ------------------------------------------------------
            LoadProjectData();
            ShowAlert("تم التعديل بنجاح", AlertForm.AlertType.Success);
            try
            {
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string mainFolder = Path.Combine(desktopPath, "ملفات و مستندات الطرح العام");
                string projectFolder = Path.Combine(mainFolder, serial_number_TB.Text);

                // إنشاء الفولدر لو مش موجود
                if (!Directory.Exists(projectFolder))
                    Directory.CreateDirectory(projectFolder);


                foreach (var entry in approvalFiles)
                {
                    foreach (string sourcePath in entry.Value)
                    {
                        if (File.Exists(sourcePath))
                        {
                            string fileName = Path.GetFileName(sourcePath);
                            string destinationPath = Path.Combine(projectFolder, fileName);

                            File.Copy(sourcePath, destinationPath, true); // true = overwrite لو موجود
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ أثناء نسخ الملفات:\n" + ex.Message, AlertForm.AlertType.Error);
            }
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";

            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();

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
            // Show confirmation dialog
            DialogResult result = MessageBox.Show(
                "هل أنت متأكد أنك تريد حذف المشاريع المحددة وجميع مستنداتها؟",
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
                    string projectId = row.Cells["Project_Id"].Value.ToString();
                    string CivilDefenseFileValue = row.Cells["CivilDefenseFile"].Value.ToString();
                    string CivilDefenseStudyFileValue = row.Cells["CivilDefenseStudyFile"].Value.ToString();
                    string EnvironmentalFileValue = row.Cells["EnvironmentalFile"].Value.ToString();
                    string PetroleumFileValue = row.Cells["PetroleumFile"].Value.ToString();
                    string AviationFileValue = row.Cells["AviationFile"].Value.ToString();
                    string TrafficStudyFileValue = row.Cells["TrafficStudyFile"].Value.ToString();
                    string Model8FileValue = row.Cells["Model8File"].Value.ToString();
                    string Transaction_numberValue = row.Cells["Transaction_numberFile"].Value.ToString();

                    if (!string.IsNullOrWhiteSpace(CivilDefenseFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP001", projectId);

                    if (!string.IsNullOrWhiteSpace(CivilDefenseStudyFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP010", projectId);

                    if (!string.IsNullOrWhiteSpace(EnvironmentalFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP002", projectId);

                    if (!string.IsNullOrWhiteSpace(PetroleumFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP003", projectId);

                    if (!string.IsNullOrWhiteSpace(AviationFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP004", projectId);

                    if (!string.IsNullOrWhiteSpace(TrafficStudyFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP005", projectId);

                    if (!string.IsNullOrWhiteSpace(Model8FileValue))
                        this.documentsTableAdapter.DeleteQuery("APP006", projectId);

                    if (!string.IsNullOrWhiteSpace(Transaction_numberValue))
                        this.documentsTableAdapter.DeleteQuery("APP007", projectId);

                    this.projectsTableAdapter.DeleteQuery(projectId);
                    deleted++;
                }
            }
            foreach (DataGridViewRow row in advancedDataGridView2.Rows)
            {
                if (Convert.ToBoolean(row.Cells["Select"].Value ?? false))
                {
                    string document_id = row.Cells["document_id"].Value.ToString();

                    this.documentsTableAdapter.DeleteQuery1(document_id);

                    deleted++;
                }
            }
            if (deleted > 0)
            {
                LoadProjectData();
                ShowAlert($"{deleted} صف تم حذفه بنجاح", AlertForm.AlertType.Success);
            }

            
            else
            {
                ShowAlert("لم يتم تحديد أي صفوف للحذف", AlertForm.AlertType.Warning);
            }
            UpdateRowCount();
        }

        private void skyButton2_Click_1(object sender, EventArgs e)
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

            // --- ✅ ترتيب الأعمدة حسب DisplayIndex ---
            var visibleColumns = advancedDataGridView1.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible && c.Name.ToLower() != "select")
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            int colCount = visibleColumns.Count;

            // --- ✅ Add Title Row ---
            var titleRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, colCount]];
            titleRange.Merge();

            // تنسيقات العنوان
            titleRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            titleRange.Value = fileNameTextBox.Text;

            // ✅ Header Row
            int excelCol = 1;
            foreach (var gridCol in visibleColumns)
            {
                var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[2, excelCol];
                cell.Value = gridCol.HeaderText;
                cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                cell.Font.Bold = true;
                excelCol++;
            }

            // ✅ Write data (Row 3 onwards)
            int excelRow = 3;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                excelCol = 1;
                foreach (var gridCol in visibleColumns)
                {
                    var value = row.Cells[gridCol.Index].Value;
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


        private void guna2CircleButton2_Click(object sender, EventArgs e)
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
                parentForm.Close(); // or parentForm.Hide();
            }
        }

        private void dungeonLabel14_Click(object sender, EventArgs e)
        {

        }

        private void Total_stores_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }

        private void Total_rented_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }

        private void Total_Not_rented_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }

        private void Add_Radio_CheckedChanged(object sender, EventArgs e)
        {
            if (Add_Radio.Checked)
            {
                guna2TabControl1.SelectedTab = tabPage4;

            }
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button2.Visible = !Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Enabled = !Add_Radio.Checked;
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();

        }
        private void Update_Radio_CheckedChanged(object sender, EventArgs e)
        {
            if (Update_Radio.Checked)
            {
                guna2TabControl1.SelectedTab = tabPage4;
            }
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button1.Visible = !Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button1.Enabled = !Update_Radio.Checked;
        }
        private Dictionary<string, string> columnMap = new Dictionary<string, string>
        {
            { "اسم المشروع / نشاط", "project_name" },
            { "اسم قطعة الأرض", "LandName" },
            { "اسم المحافظة", "GovernorateName" },
            { "Name_Projects","اسم المكان" },
            { "رقم اللوحة", "PlateNumber" },
            { "رقم قطعة الأرض", "Land_Number" },
            { "مسلسل القطعة الكود", "Land_Id" },
            { "مسلسل المشروع", "Project_Id" },
            { "مؤجر", "Total_rented" },
            { "غير مؤجر", "Total_Not_rented" },
            { "حالة موافقة الحماية المدنية", "Civil_Defense_Approval_status" },
            { "حالة دراسة الحماية المدنية", "Civil_Defense_Study_status" },
            { "حالة موافقة البيئة", "Environmental_Approval_status" },
            { "حالة موافقة وزارة البترول", "Petroleum_Ministry_Approval_status" },
            { "حالة موافقة الطيران المدني", "Civil_Aviation_Approval_status" },
            { "حالة الدراسة المرورية", "Traffic_Study_Status" },
            { "حالة نموذج 8 أو 10", "Model_8_Status" },
            { "رقم المعاملة", "Transaction_number" },
            { "تاريخ انتهاء العقد", "Contract_expiry_date" },
            { "اجمالي محلات", "Total_stores" },
            { "الشهادة المؤمنه", "Secured_Certificate" },
            { "شركة ادارة المحطة", "Architectural_and_Structural_Board" },
            { "ختم نموذج التصالح", "Reconciliation_Form_Stamp" },
            { "الرفع المساحي الاستشاري", "Consultant_Surveying" },
            { "المكتب الاستشاري", "consulting_Office" },
    };
        private void skyButton4_Click(object sender, EventArgs e)
        {

            DataTable original =new DataTable();

            if (advancedDataGridView1.DataSource is BindingSource bs)
            {
                original = ((DataView)bs.List).ToTable();
            }
            else if (advancedDataGridView1.DataSource is DataView dv)
            {
                original = dv.ToTable();
            }
            else if (advancedDataGridView1.DataSource is DataTable dt)
            {
                original = dt.Copy();
            }
            else
            {
                ShowAlert("Unsupported DataSource type.",AlertForm.AlertType.Error);
            }
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




        }


        private void advancedDataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            advancedDataGridView1.Invalidate();
        }

        private void Report_TB_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2ImageButton6_Click_1(object sender, EventArgs e)
        {
            Files("عقود محلات");
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            var documents = this.documentsTableAdapter.GetData();

            int lastDocumentId = 0;

            if (documents != null && documents.Count > 0)
            {
                lastDocumentId = documents
                    .Select(d =>
                    {
                        int parsed;
                        return int.TryParse(d.document_id, out parsed) ? parsed : 0;
                    })
                    .OrderBy(id => id)   // ترتيب كأرقام
                    .Last();             // آخر ID بعد الترتيب
            }
            else
            {
                lastDocumentId = -1; // أو 0 حسب احتياجك
            }

            int idNext = lastDocumentId;
            foreach (var item in AllfilePaths)
            {
                byte[] fileData = File.ReadAllBytes(@item);
                byte[] defaultValue = new byte[] { 0x00 };
                if (item.Contains("كروكيات"))
                {
                    this.documentsTableAdapter.Insert(
                        (++idNext).ToString(),
                        "APP009",
                        guna2TextBox1.Text,
                        item,
                        defaultValue
                        );
                }
                else {
                    this.documentsTableAdapter.Insert(
                        (++idNext).ToString(),
                        "APP008",
                        guna2TextBox1.Text,
                        item,
                        defaultValue
                        );
                }
            }
            if (AllfilePaths.Count >0)
            {
                loadDocs();
                UpdateRowCount();
                ShowAlert("تم الحفظ بنجاح", AlertForm.AlertType.Success);
                AllfilePaths = new List<string>();
            }
            else
            {
                ShowAlert("يرجي أضافة مستند واحد علي الاقل", AlertForm.AlertType.Warning);
                AllfilePaths = new List<string>();
            }
        }

        private void advancedDataGridView2_Scroll(object sender, ScrollEventArgs e)
        {
            advancedDataGridView2.Invalidate();
        }

        private void advancedDataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ensure click is not on header row
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var clickedColumn = advancedDataGridView2.Columns[e.ColumnIndex];

            // Check if the clicked column is the "pdf" column
            if (clickedColumn.Name.Contains("MarketsFile"))
            {
                string filePath = "";
                // Get the file path from the cell value
                if (clickedColumn.Name == "MarketsFile") { filePath = advancedDataGridView2.Rows[e.RowIndex].Cells["MarketsFile"].Value?.ToString(); }
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

        private void guna2ImageButton8_Click(object sender, EventArgs e)
        {
            flowLayoutPanel8.Controls.Clear();

            if (approvalFiles.ContainsKey("عقود محلات"))
            {
                approvalFiles.Remove("عقود محلات");
            }
         AllfilePaths= new List<string>();
        }

        private void advancedDataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (advancedDataGridView2.Columns[e.ColumnIndex].Name == "Select" && e.RowIndex >= 0)
            {
                var cell = advancedDataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex];
                bool isChecked = Convert.ToBoolean(cell.Value ?? false);
                cell.Value = !isChecked;
            }
        }

        private void advancedDataGridView2_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            var grid = sender as Zuby.ADGV.AdvancedDataGridView;
            Properties.Settings.Default.LastFilter = (advancedDataGridView2.DataSource as BindingSource)?.Filter;
            Properties.Settings.Default.LastFilterUI = grid.FilterString; // The visible string
            Properties.Settings.Default.Save();
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView2.DataSource;
            b.Filter = advancedDataGridView2.FilterString;

            // Update the row count
            if (Road_Radio.Checked)
            {
                UpdateRowCount();
            }
            else if (Station_Radio.Checked)
            {
                int documentsCount = 0;

                foreach (DataGridViewRow row in advancedDataGridView2.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        documentsCount++;
                    }
                }
                DocumentsCountLabel.Text = $"عدد المستندات المطلوب تسجيلها: {documentsCount}";
            }
        }         
        private void guna2Shapes3_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            

        }

        private void guna2Button6_Click(object sender, EventArgs e)
        {
           
        }

        private void Road_Radio_CheckedChanged(object sender, EventArgs e)
        {
            if (Road_Radio.Checked)
            {
                loadDocs();
                UpdateRowCount();
            }

        }

        private void Station_Radio_CheckedChanged(object sender, EventArgs e)
        {
            if (Station_Radio.Checked)
            {

           
            loadInvestmentsWithoutDocs();
            int documentsCount = 0;

            foreach (DataGridViewRow row in advancedDataGridView2.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    documentsCount++;
                }
            }
            DocumentsCountLabel.Text = $"عدد المستندات المطلوب تسجيلها: {documentsCount}";
            }
        }

        private void skyButton5_Click(object sender, EventArgs e)
        {
            TrueFunction();
            if (advancedDataGridView2.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات", AlertForm.AlertType.Error);
                return;
            }

            string fileName = guna2TextBox4.Text.Trim();
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

            // --- ✅ ترتيب الأعمدة حسب DisplayIndex ---
            var visibleColumns = advancedDataGridView2.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible && c.Name.ToLower() != "select")
                .OrderBy(c => c.DisplayIndex)
                .ToList();

            int colCount = visibleColumns.Count;

            // --- ✅ Add Title Row ---
            var titleRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, colCount]];
            titleRange.Merge();

            // تنسيقات العنوان
            titleRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            titleRange.Value = guna2TextBox4.Text;

            // ✅ Header Row
            int excelCol = 1;
            foreach (var gridCol in visibleColumns)
            {
                var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[2, excelCol];
                cell.Value = gridCol.HeaderText;
                cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                cell.Font.Bold = true;
                excelCol++;
            }

            // ✅ Write data (Row 3 onwards)
            int excelRow = 3;
            foreach (DataGridViewRow row in advancedDataGridView2.Rows)
            {
                if (row.IsNewRow) continue;

                excelCol = 1;
                foreach (var gridCol in visibleColumns)
                {
                    var value = row.Cells[gridCol.Index].Value;
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

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void guna2TabControl1_Click(object sender, EventArgs e)
        {

        }

        private void guna2TabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index == guna2TabControl1.TabPages.IndexOf(tabPage4))
            {
                // Do nothing → skip drawing this tab header
                return;
            }

            // Draw the other tabs normally
            e.Graphics.DrawString(
                guna2TabControl1.TabPages[e.Index].Text,
                guna2TabControl1.Font,
                Brushes.Black,
                e.Bounds.X + 3,
                e.Bounds.Y + 3
            );
        }

        private void guna2TabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage4&& Add_Radio.Checked== false && Update_Radio.Checked == false )
                e.Cancel = true;
            else if (e.TabPage == tabPage4 && (Add_Radio.Checked == true|| Update_Radio.Checked==true))
            {
                e.Cancel = false;
            }
            //if (e.TabPage != tabPage4)
            //{
            //    Update_Radio.Checked = false;
            //    Add_Radio.Checked = false;
            //}
        }

        private void guna2CirclePictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void advancedDataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string columnName = advancedDataGridView1.Columns[e.ColumnIndex].Name;
                if (columnName.Equals("select", StringComparison.OrdinalIgnoreCase))
                    return;
                string newValue = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                string Project_Id = advancedDataGridView1.Rows[e.RowIndex].Cells["Project_Id"].Value?.ToString();

                if (string.IsNullOrEmpty(Project_Id)) return;

                // Get the specific row(s) from adapter
                var table = projectsTableAdapter.GetDataByIDProjects(Project_Id);

                if (table.Rows.Count > 0)
                {
                    var row = table[0]; // get the first row

                    // Update the field value
                    if (string.IsNullOrEmpty(newValue))
                        row[columnName] = DBNull.Value;
                    else
                        row[columnName] = newValue;

                    // Save back to database
                    projectsTableAdapter.Update(row);


                }
                else
                {
                    MessageBox.Show("لم يتم التسجيل بشكل صحيح");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("حدث خطأ أثناء الحفظ: " + ex.Message, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private DataTable originalData;
        private void dreamButton13_Click(object sender, EventArgs e)
        {
            if (originalData != null)
            {
                DataView view = new DataView(originalData);
                view.RowFilter = "[Secured_Certificate] IS NOT NULL AND [Secured_Certificate] <> ''";
                advancedDataGridView1.DataSource = view;
                UpdateRowCount();
            }
        }

        private void guna2ImageButton10_Click(object sender, EventArgs e)
        {
            Files("عقود محلات");
        }

        private void advancedDataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            var grid = advancedDataGridView1;
            var row = grid.Rows[e.RowIndex];

            if (row.IsNewRow) return;

            if (fillterProjects == "")
            {
            string transaction = row.Cells["Transaction_number"].Value?.ToString()?.Trim() ?? "";
                if (transaction == "✔")
                {
                    string[] orangeColumns =
                      {
                    "Civil_Defense_Approval_status",
                    "Civil_Defense_Study_status",
                    "Environmental_Approval_status",
                    "Petroleum_Ministry_Approval_status",
                    "Civil_Aviation_Approval_status",
                    "Traffic_Study_Status",
                    "Model_8_Status"
                };

                    foreach (string colName in orangeColumns)
                    {
                        if (grid.Columns.Contains(colName))
                        {
                            var cell = row.Cells[colName];
                            string cellValue = cell.Value?.ToString()?.Trim() ?? "";

                            if (cellValue == "X")
                            {
                                cell.Style.BackColor = Color.Orange;
                                cell.Style.ForeColor = Color.Black;
                            }
                        }
                    }

                }
            }
            string notes = row.Cells["Secured_Certificate"].Value?.ToString()?.Trim() ?? "";
            // RED CONDITION
            if (!string.IsNullOrEmpty(notes) && notes.Contains("الغاء"))
            {
                row.DefaultCellStyle.BackColor = Color.Red;
                row.DefaultCellStyle.ForeColor = Color.White;
            }
            // YELLOW CONDITION
            else if (!string.IsNullOrEmpty(notes) && notes.Contains("اضافة"))
            {
                row.DefaultCellStyle.BackColor = Color.Yellow;
                row.DefaultCellStyle.ForeColor = Color.Black;
            }
            
            else
            {
                // Reset (important when filtering or refreshing)
                row.DefaultCellStyle.BackColor = grid.DefaultCellStyle.BackColor;
                row.DefaultCellStyle.ForeColor = grid.DefaultCellStyle.ForeColor;
            }
            
        }



        private void dreamButton14_Click(object sender, EventArgs e)
        {

        }



        private void Add_Radio_Click(object sender, EventArgs e)
        {
        }
        public void LoadNotes()
        {
            var query = from n in this.note_pageTableAdapter.GetData()
                        where n.Page_Name == "PRO"
                        select new
                        {
                            n.Id,
                            Code = n.IsCodeNull() ? 0 : n.Code,
                            Note_Page = n.IsNote_PageNull() ? "" : n.Note_Page,
                            DateTime = n.IsDateTimeNull() ? (DateTime?)null : n.DateTime,
                            status = n.IsStatusNull() ? "" : n.Status
                        };
            var OrderNoteId = query.OrderBy(n => n.Id).ToList();
            DataTable table = ToDataTable(OrderNoteId);
            BindingSource Binding = new BindingSource();
            Binding.DataSource = table;
            advancedDataGridView3.DataSource = Binding.DataSource;
        }
        private void guna2GradientCircleButton1_Click(object sender, EventArgs e)
        {
            var notes = this.note_pageTableAdapter.GetData();
            this.note_pageTableAdapter.Insert(++notes.Last().Id, "", "PRO", DateTime.Today, "جديد", 0);
            LoadNotes();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            GenerativePanalFlow();
            LoadColumnsIntoCheckedListBox();
        }
        HashSet<string> NotesColumns = new HashSet<string>
        {
            "Id",
            "Code",
            "Note_Page",
            "DateTime",
            "status"

        };
        private TextBox editingTextBox = null;
        private void advancedDataGridView3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                string columnName = advancedDataGridView3.Columns[e.ColumnIndex].Name;
                DataGridViewCell cell =
                  advancedDataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex];

                string NoteId = advancedDataGridView3.Rows[e.RowIndex]
                                .Cells["Id"].Value?.ToString();

                if (string.IsNullOrEmpty(NoteId))
                    return;
                // NORMAL TEXT / NUMERIC COLUMNS
                // =================================================
                string newValue = editingTextBox?.Text ?? cell.Value?.ToString();
                editingTextBox = null;

                if (NotesColumns.Contains(columnName))
                {

                    var table = note_pageTableAdapter.GetDataByID(int.Parse(NoteId));
                    if (table.Rows.Count == 0) return;

                    var row = table[0];

                    if (string.IsNullOrEmpty(newValue))
                        row[columnName] = DBNull.Value;
                    else
                        row[columnName] = newValue.Replace("\u200E", "");

                    note_pageTableAdapter.Update(row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                   "حدث خطأ أثناء الحفظ:\n" + ex.Message,
                   "خطأ",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error
               );
            }
        }

        private string GetNextProjectId()
        {
            var table = projectsTableAdapter.GetData();

            int maxId = 0;
            foreach (var row in table)
            {
                if (int.TryParse(row.project_id, out int id))
                    if (id > maxId) maxId = id;
            }

            return (maxId + 1).ToString();
        }
        private void guna2CircleButton3_Click(object sender, EventArgs e)
        {

            try
            {
                // 1️⃣ إنشاء صف جديد في جدول lands
                var projectTable = projectsTableAdapter.GetData();
                var newRow = projectTable.NewprojectsRow();

                // 2️⃣ توليد land_fk جديد (حسب طريقتك)
                newRow.project_id = GetNextProjectId();   // ⬅️ دالة هنكتبها تحت
                newRow.project_name = " ";
                newRow.governorate_fk = "0";
                newRow.Total_Not_rented = 0;
                newRow.Contract_expiry_date = DateTime.Now;
                newRow.Civil_Defense_Approval_status = "X";
                newRow.Civil_Defense_Study_status = "X";
                newRow.Environmental_Approval_status = "X";
                newRow.Traffic_Study_Status = "X";
                newRow.Model_8_Status = "X";
                newRow.Petroleum_Ministry_Approval_status = "X";
                newRow.Civil_Aviation_Approval_status = "X";
                newRow.land_fk = "0";
                newRow.Transaction_number = "";
                newRow.Total_stores = 0;
                newRow.Total_rented = 0;
                newRow.Secured_certificate = "";
                newRow.Architectural_and_Structural_Board = "";
                newRow.Reconciliation_Form_Stamp = "";
                newRow.Consultant_Surveying = "";
                newRow.Name_Projects = "";


                //newRow.Notes = "تم اضافة قطعة ارض جديدة";
                projectTable.AddprojectsRow(newRow);
                projectsTableAdapter.Update(projectTable);
                LoadProjectData();
                ArabicColumnGrid();
                UpdateRowCount();
                ApplyGuna2StyleToGrid(advancedDataGridView1);
                GenerativePanalFlow();
                LoadColumnsIntoCheckedListBox();

                // 3️⃣ إضافة الصف إلى الـ DataGridView
                DataTable gridTable = null;

                if (advancedDataGridView1.DataSource is BindingSource bs)
                {
                    if (bs.DataSource is DataView dv)
                        gridTable = dv.Table;
                    else if (bs.DataSource is DataTable dt)
                        gridTable = dt;
                }
                else if (advancedDataGridView1.DataSource is DataTable dt)
                {
                    gridTable = dt;
                }

                if (gridTable == null)
                {
                    MessageBox.Show("DataSource غير مدعوم");
                    return;
                }
                DataRow gridRow = gridTable.NewRow();
                gridRow["project_id"] = int.Parse(newRow.project_id);
                gridRow["project_name"] = "";
                //gridRow["total_area"] = 0;

                gridTable.Rows.Add(gridRow);

                //4️⃣ تحديد الصف الجديد تلقائيًا
                int rowIndex = advancedDataGridView1.Rows.Count - 1;

                if (rowIndex >= 0)
                {
                    advancedDataGridView1.ClearSelection(); // نشيل أي تحديد قديم

                    var row = advancedDataGridView1.Rows[rowIndex];

                    if (row.Visible)
                    {
                        row.Selected = true; // 👈 تحديد الصف بالكامل

                        // optional: يخلي الصف يظهر في الشاشة
                        advancedDataGridView1.FirstDisplayedScrollingRowIndex = rowIndex;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "حدث خطأ أثناء إضافة صف جديد:\n" + ex.Message,
                    "خطأ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        class PageState
        {
            public int SelectedTabIndex { get; set; }
            public string FilterString { get; set; }
            public string SortString { get; set; }
            public DataTable Data { get; set; }
        }

        List<PageState> history = new List<PageState>();
        int currentIndex = -1;

        bool isRestoring = false;
        private void RestoreState(PageState state)
        {
            isRestoring = true;

            guna2TabControl1.SelectedIndex = state.SelectedTabIndex;

            if (state.Data != null)
            {
                originalData = state.Data.Copy();

                DataView view = new DataView(originalData);
                view.RowFilter = state.FilterString;

                BindingSource bs = new BindingSource();
                bs.DataSource = view;
                bs.Sort = state.SortString;

                advancedDataGridView1.DataSource = bs;
            }

            advancedDataGridView1.Refresh();
            UpdateRowCount();

            isRestoring = false;
        }
        private void ApplyFilterInternal(string filter)
        {
            if (originalData == null) return;

            DataView view = new DataView(originalData);
            view.RowFilter = filter;

            BindingSource bs = new BindingSource();
            bs.DataSource = view;

            advancedDataGridView1.DataSource = bs;

            UpdateRowCount();
        }

        private void SaveState()
        {
            if (isRestoring) return;

            BindingSource bs = null;
            DataView dv = null;

            if (advancedDataGridView1.DataSource is BindingSource)
            {
                bs = (BindingSource)advancedDataGridView1.DataSource;
                dv = bs.DataSource as DataView;
            }
            else if (advancedDataGridView1.DataSource is DataTable dt)
            {
                dv = dt.DefaultView;
            }

            var state = new PageState
            {
                SelectedTabIndex = guna2TabControl1.SelectedIndex,
                FilterString = dv?.RowFilter,
                SortString = bs?.Sort,
                Data = originalData?.Copy()
            };

            if (currentIndex < history.Count - 1)
                history.RemoveRange(currentIndex + 1, history.Count - currentIndex - 1);

            history.Add(state);
            currentIndex++;
        }


        private void guna2CircleButton6_Click(object sender, EventArgs e)
        {
            if (currentIndex <= 0) return;

            currentIndex--;

            RestoreState(history[currentIndex]);
            UpdateRowCount();
        }

        private void guna2CircleButton7_Click(object sender, EventArgs e)
        {
            if (currentIndex >= history.Count - 1) return;

            currentIndex++;

            RestoreState(history[currentIndex]);
            UpdateRowCount();
        }

        private void guna2TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
           if(guna2TabControl1.SelectedTab!=tabPage5)
            {
                dreamButton22.Visible = false;
            }
            else { dreamButton22.Visible = true; }
            if (isRestoring) return;

            SaveState();
        }
        private void AddFiterToButtonDreamButton(string ch)
        {
            fillterProjects = ch;

            DataTable converted = null;
            try
            {
                TrueFunction();

                // Step 1: Load all data once
                var projects = projectsTableAdapter.GetData().ToList();
                var governorates = governorateTableAdapter.GetData().ToList();
                var lands = landsTableAdapter.GetData().ToList();
                var documents = documentsTableAdapter.GetData().ToList();
                var approvals = approvalsTableAdapter.GetData().ToList();
                var investments = investmentsTableAdapter.GetData().ToList();

                // Step 2: Pre-join documents + approvals to avoid repeating lookups
                var docsWithApprovals = (
                    from d in documents
                    join a in approvals on d.approvals_fk equals a.approval_id
                    select new
                    {
                        ProjectId = d.projects_fk,
                        ApprovalName = a.approvals,
                        Path = d.paths
                    }
                ).ToList();

                // Step 3: Create main query (pure in-memory join)
                var query =
                    from p in projects
                    join g in governorates on p.governorate_fk equals g.governorate_id
                    join l in lands on p.land_fk equals l.land_id
                    let inv = investments.Where(x => x.land_fk == l.land_id)
                    select new
                    {

                        Name_Projects = p.IsName_ProjectsNull() ? "" : p.Name_Projects,

                        land_fk = p.Island_fkNull() ? "" : p.land_fk,

                        Project_Id = p.project_id,

                        Address = l.IsAddressNull()?"":l.Address,
                        GovernorateName = g.IsgovernorateNull()?"":g.governorate,


                        //consulting_Office = l.consulting_Office,

                        //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                       

                       
                        Contract_expiry_date = p.Contract_expiry_date,
                        Secured_Certificate = p.IsSecured_certificateNull() ? "" : p.Secured_certificate
                    };

                var joinedList = query.OrderBy(r => int.TryParse(r.Project_Id, out var n) ? n : int.MaxValue).Where(r => r.land_fk.Contains(ch)).ToList();
                DataTable original = ToDataTable(joinedList);

                // Step 4: Convert Contract_expiry_date column
                DataTable tempConverted = original.Clone();
                tempConverted.Columns["Contract_expiry_date"].DataType = typeof(DateTime);

                foreach (DataRow row in original.Rows)
                {
                    var newRow = tempConverted.NewRow();
                    foreach (DataColumn col in original.Columns)
                    {
                        if (col.ColumnName == "Contract_expiry_date")
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
                    tempConverted.Rows.Add(newRow);
                }

                converted = tempConverted;

                // Step 5: Format DateTime columns
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                }

                // Step 6: Bind data to DataGridView
                advancedDataGridView1.DataSource = converted;

                if (!advancedDataGridView1.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn
                    {
                        HeaderText = "تحديد",
                        Name = "Select",
                        Width = 60,
                        ReadOnly = false,
                        TrueValue = true,
                        FalseValue = false
                    };
                    advancedDataGridView1.Columns.Add(checkBoxColumn);
                }
                var getAccesUserFunction = accessTableAdapter.GetDataByOFAccessFunctionUserId(UserId, 4, 5);

                if (getAccesUserFunction.Count == 1)
                {
                    if (!advancedDataGridView1.Columns.Contains("Update"))
                    {
                        DataGridViewButtonColumn btnUpdate = new DataGridViewButtonColumn
                        {
                            HeaderText = "تعديل",
                            Name = "Update",
                            Text = "تعديل",
                            Tag = "Function:Edit",
                            UseColumnTextForButtonValue = true,
                            DisplayIndex = 0,
                            Width = 80
                        };

                        advancedDataGridView1.Columns.Add(btnUpdate);
                    }

                }
                advancedDataGridView1.Columns["Select"].DisplayIndex = 0;
                originalData = converted;
                loadDocs();
                LoadNotes();
                SaveState();
                UpdateRowCount();
                GenerativePanalFlow();
                ApplyGuna2StyleToGrid(advancedDataGridView1);
                ApplyGuna2StyleToGrid(advancedDataGridView2);
                ApplyGuna2StyleToGrid(advancedDataGridView3);
                ArabicColumnGrid();
                //LoadColumnsIntoCheckedListBox();
                //checkedListBox1.Items.Clear();
                //checkedListBox1.Items.Add("تحديد", true);
                //checkedListBox1.Items.Add("تعديل", true);
                //checkedListBox1.Items.Add("اسم المكان", true);
                //checkedListBox1.Items.Add("مسلسل القطعة الكود", true);
                //checkedListBox1.Items.Add("اسم المحافظة", true);
                //checkedListBox1.Items.Add("ملاحظات", true);
                

            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            finally
            {
                FalseFunction();
            }
        }
        string fillterProjects = "";
        private void dreamButton14_Click_1(object sender, EventArgs e)
        {
            fillterProjects = "";
            advancedDataGridView1.CleanSort();
            advancedDataGridView1.CleanFilter();
            advancedDataGridView1.DataSource = null;
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            AddFiterToButtonDreamButton("#");
            LoadNotes();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();
            tabPage4.Text = "";
        }

        private void dreamButton19_Click(object sender, EventArgs e)
        {
            fillterProjects = "";
            advancedDataGridView1.CleanSort();
            advancedDataGridView1.CleanFilter();
            advancedDataGridView1.DataSource = null;
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            AddFiterToButtonDreamButton("&");
            LoadNotes();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();
            tabPage4.Text = "";
        }

        private void AddFiterToButtonDreamButton2()
        {
            DataTable converted = null;
            try
            {
                TrueFunction();

                // Step 1: Load all data once
                var projects = projectsTableAdapter.GetData().ToList();
                var governorates = governorateTableAdapter.GetData().ToList();
                var lands = landsTableAdapter.GetData().ToList();
                var documents = documentsTableAdapter.GetData().ToList();
                var approvals = approvalsTableAdapter.GetData().ToList();
                var investments = investmentsTableAdapter.GetData().ToList();

                // Step 2: Pre-join documents + approvals to avoid repeating lookups
                var docsWithApprovals = (
                    from d in documents
                    join a in approvals on d.approvals_fk equals a.approval_id
                    select new
                    {
                        ProjectId = d.projects_fk,
                        ApprovalName = a.approvals,
                        Path = d.paths
                    }
                ).ToList();

                // Step 3: Create main query (pure in-memory join)
                var query =
                    from p in projects
                    join g in governorates on p.governorate_fk equals g.governorate_id
                    join l in lands on p.land_fk equals l.land_id
                    let inv = investments.Where(x => x.land_fk == l.land_id)
                    select new
                    {
                        project_name = p.project_name,
                        LandName = l.Island_nameNull() ? "" : l.land_name,
                        Name_Projects = p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                        Address = l.IsAddressNull() ? "" : l.Address,
                        GovernorateName = g.IsgovernorateNull() ? "" : g.governorate,
                        PlateNumber = l.Isplate_numberNull() ? "" : l.plate_number,
                        Land_Number = l.Island_numberNull() ? "" : l.land_number,
                        land_fk = p.Island_fkNull() ? "" : p.land_fk,
                        Project_Id = p.project_id,
                        area = l.Istotal_areaNull() ? "" : l.total_area,
                        //Total_stores = inv.Where(i => i.Visable_Value == true).Sum(i => i.IsShops_CountNull() ? 0 : i.Shops_Count),

                        //                    Total_rented =
                        //inv.Where(i =>
                        //    (!i.IsActivity_NameNull() && i.Activity_Name != "لا يوجد") &&
                        //    i.IsRental_StatusNull() && i.Visable_Value == true
                        //)
                        //.Sum(i => i.IsShops_CountNull() ? 0 : i.Shops_Count),

                        //                    Total_Not_rented =
                        //inv.Where(i =>
                        //    (!i.IsActivity_NameNull() && i.Activity_Name == "لا يوجد") &&
                        //    (i.IsRental_StatusNull()||i.Rental_Status== "غير مؤجر (*)") && i.Visable_Value == true
                        //)
                        //.Sum(i => i.IsShops_CountNull() ? 0 : i.Shops_Count),

                        consulting_Office = l.consulting_Office,

                        Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                        //Reconciliation_Form_Stamp = p.IsReconciliation_Form_StampNull()?"":p.Reconciliation_Form_Stamp,
                        Consultant_Surveying = p.Consultant_Surveying,

                        Civil_Defense_Approval_status = p.Civil_Defense_Approval_status,
                        CivilDefenseFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة الحماية المدنية")
                .Select(d => d.Path)),

                        Civil_Defense_Study_status = p.Civil_Defense_Study_status,
                        CivilDefenseStudyFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "دراسة حماية مدنيه")
                .Select(d => d.Path)),

                        Environmental_Approval_status = p.Environmental_Approval_status,
                        EnvironmentalFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة البيئة")
                .Select(d => d.Path)),

                        Petroleum_Ministry_Approval_status = p.Petroleum_Ministry_Approval_status,
                        PetroleumFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة وزارة البترول")
                .Select(d => d.Path)),

                        Civil_Aviation_Approval_status = p.Civil_Aviation_Approval_status,
                        AviationFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "موافقة الطيران المدني")
                .Select(d => d.Path)),

                        Traffic_Study_Status = p.Traffic_Study_Status,
                        TrafficStudyFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "الدراسة المرورية")
                .Select(d => d.Path)),

                        Model_8_Status = p.Model_8_Status,
                        Model8File = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "نموذج 8 أو 10")
                .Select(d => d.Path)),

                        Transaction_number = p.Transaction_number,
                        Transaction_numberFile = string.Join(" , ",
            docsWithApprovals
                .Where(d => d.ProjectId == p.project_id &&
                            d.ApprovalName == "رقم المعاملة")
                .Select(d => d.Path)),
                        Contract_expiry_date = p.Contract_expiry_date,
                        Secured_Certificate = p.IsSecured_certificateNull() ? "" : p.Secured_certificate
                    };

                var joinedList = query.OrderBy(r => int.TryParse(r.Project_Id, out var n) ? n : int.MaxValue)
                    .Where(r => !r.land_fk.Contains("#")&& r.Project_Id!="0"&& !r.land_fk.Contains("&") && !r.land_fk.Contains("$")&& r.land_fk!="215"&& r.land_fk != "216"&& r.Secured_Certificate != "تم الغاء قطعة الارض").ToList();
                DataTable original = ToDataTable(joinedList);

                // Step 4: Convert Contract_expiry_date column
                DataTable tempConverted = original.Clone();
                tempConverted.Columns["Contract_expiry_date"].DataType = typeof(DateTime);

                foreach (DataRow row in original.Rows)
                {
                    var newRow = tempConverted.NewRow();
                    foreach (DataColumn col in original.Columns)
                    {
                        if (col.ColumnName == "Contract_expiry_date")
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
                    tempConverted.Rows.Add(newRow);
                }

                converted = tempConverted;

                // Step 5: Format DateTime columns
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                }

                // Step 6: Bind data to DataGridView
                advancedDataGridView1.DataSource = converted;

                if (!advancedDataGridView1.Columns.Contains("Select"))
                {
                    DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn
                    {
                        HeaderText = "تحديد",
                        Name = "Select",
                        Width = 60,
                        ReadOnly = false,
                        TrueValue = true,
                        FalseValue = false
                    };
                    advancedDataGridView1.Columns.Add(checkBoxColumn);
                }
                var getAccesUserFunction = accessTableAdapter.GetDataByOFAccessFunctionUserId(UserId, 4, 5);

                if (getAccesUserFunction.Count == 1)
                {
                    if (!advancedDataGridView1.Columns.Contains("Update"))
                    {
                        DataGridViewButtonColumn btnUpdate = new DataGridViewButtonColumn
                        {
                            HeaderText = "تعديل",
                            Name = "Update",
                            Text = "تعديل",
                            Tag = "Function:Edit",
                            UseColumnTextForButtonValue = true,
                            DisplayIndex = 0,
                            Width = 80
                        };

                        advancedDataGridView1.Columns.Add(btnUpdate);
                    }

                }
                advancedDataGridView1.Columns["Select"].DisplayIndex = 0;
                originalData = converted;
                loadDocs();
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            finally
            {
                FalseFunction();
            }
        }
        private void dreamButton23_Click(object sender, EventArgs e)
        {
           
            fillterProjects = "";
            advancedDataGridView1.CleanSort();
            advancedDataGridView1.CleanFilter();
            advancedDataGridView1.DataSource = null;
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            LoadProjectData();
            LoadNotes();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();
            tabPage4.Text = "";

        }

        private void dreamButton24_Click(object sender, EventArgs e)
        {
            fillterProjects = "";
            advancedDataGridView1.CleanSort();
            advancedDataGridView1.CleanFilter();
            advancedDataGridView1.DataSource = null;
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            AddFiterToButtonDreamButton("$");
            LoadNotes();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            Governorate_COB.SelectedItem = -1;
            Investment_Name_TB.Text = "";
            Civil_Defense_COB.SelectedItem = -1;
            Civil_Study_COB.SelectedItem = -1;
            Environmental_COB.SelectedItem = -1;
            Traffic_Study_COB.SelectedItem = -1;
            Model_8_COB.SelectedItem = -1;
            Petroleum_Ministry_COB.SelectedItem = -1;
            Civil_Aviation_COB.SelectedItem = -1;
            Land_ID_COB.SelectedItem = -1;
            Transaction_number_TB.Text = "";
            Total_stores_TB.Text = "";
            Total_rented_TB.Text = "";
            Total_Not_rented_TB.Text = "";
            Secured_certificate_COM.SelectedItem = -1;
            Architectural_and_Structural_Board_COM.SelectedItem = -1;
            Reconciliation_Form_Stamp_COM.SelectedItem = -1;
            Consultant_Surveying_COM.SelectedItem = -1;
            serial_number_TB.Text = "";
            approvalFiles = new Dictionary<string, List<string>>();
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel2.Controls.Clear();
            flowLayoutPanel3.Controls.Clear();
            flowLayoutPanel4.Controls.Clear();
            flowLayoutPanel5.Controls.Clear();
            flowLayoutPanel6.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            flowLayoutPanel9.Controls.Clear();
            tabPage4.Text = "";
        }


    }
}
