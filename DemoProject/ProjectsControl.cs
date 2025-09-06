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
            Model_8_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Traffic_Study_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Civil_Aviation_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
            Environmental_COB.SelectedIndexChanged += COB_SelectedIndexChanged;
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
        private void ProjectsControl_Load(object sender, EventArgs e)
        {
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
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
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
        }
        public void ArabicColumnGrid()
        {
            // Set Arabic headers manually
            advancedDataGridView1.Columns["Land_Id"].HeaderText = "مسلسل القطعة";
            advancedDataGridView1.Columns["Project_Id"].HeaderText = "مسلسل المشروع";
            advancedDataGridView1.Columns["Project_Id"].Visible = false;
            advancedDataGridView1.Columns["PlateNumber"].HeaderText = "رقم اللوحة";
            advancedDataGridView1.Columns["LandName"].HeaderText = "اسم قطعة الأرض";
            advancedDataGridView1.Columns["Land_Number"].HeaderText = "رقم قطعة الأرض";
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
            advancedDataGridView1.Columns["Transaction_numberFile"].HeaderText = "ملف رقم المعاملة";
            advancedDataGridView1.Columns["Contract_expiry_date"].HeaderText = "تاريخ انتهاء العقد";
            advancedDataGridView1.Columns["Total_stores"].HeaderText = "اجمالي محلات";
            advancedDataGridView1.Columns["Total_rented"].HeaderText = "مؤجر";
            advancedDataGridView1.Columns["Total_Not_rented"].HeaderText = "غير مؤجر";
            advancedDataGridView1.Columns["Secured_Certificate"].HeaderText = "الشهادة المؤمنه";
            advancedDataGridView1.Columns["Architectural_and_Structural_Board"].HeaderText = "لوحة المعماري والانشائي";
            advancedDataGridView1.Columns["Reconciliation_Form_Stamp"].HeaderText = "ختم نموذج التصالح";
            advancedDataGridView1.Columns["Consultant_Surveying"].HeaderText = "الرفع المساحي الاستشاري";
            advancedDataGridView1.Columns["consulting_Office"].HeaderText = "المكتب الاستشاري";
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
        public void LoadProjectData()
        {
            DataTable converted = null;
            try
            {
                TrueFunction();

                // --- Part 1: Fetch and join data ---
                var query = from p in projectsTableAdapter.GetData()
                            join g in governorateTableAdapter.GetData() on p.governorate_fk equals g.governorate_id
                            join l in landsTableAdapter.GetData() on p.land_fk equals l.land_id
                            select new
                            {
                                Land_Id = l.land_id,
                                PlateNumber = l.plate_number,
                                Land_Number = l.land_number,
                                Project_Id = p.project_id,
                                LandName = l.land_name,
                                ProjectName = p.project_name,
                                GovernorateName = g.governorate,
                                Total_stores = p.Total_stores,
                                Total_rented = p.Total_rented,
                                Total_Not_rented = p.Total_Not_rented,
                                consulting_Office = l.consulting_Office,
                                Secured_Certificate = p.Secured_certificate,
                                Architectural_and_Structural_Board = p.Architectural_and_Structural_Board,
                                Reconciliation_Form_Stamp = p.Reconciliation_Form_Stamp,
                                Consultant_Surveying = p.Consultant_Surveying,
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
                                Transaction_number = p.Transaction_number,
                                Transaction_numberFile = (
                                    from d in documentsTableAdapter.GetData()
                                    join a in approvalsTableAdapter.GetData() on d.approvals_fk equals a.approval_id
                                    where d.projects_fk == p.project_id && a.approvals == "رقم المعاملة"
                                    select d.paths
                                ).FirstOrDefault(),
                                Contract_expiry_date = p.Contract_expiry_date
                            };

                var joinedList = query.OrderBy(r => int.TryParse(r.Project_Id, out var n) ? n : int.MaxValue)
                .ToList();
                DataTable original = ToDataTable(joinedList);

                // --- Convert string date column to DateTime type ---
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
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                    }
                }
                // --- Bind data to DataGridView ---
                advancedDataGridView1.DataSource = converted;

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

        public void LoadColumnsIntoCheckedListBox()
        {
            // Clear previous items
            checkedListBox1.Items.Clear();

            // Make sure the DataGridView has a DataSource
            if (advancedDataGridView1.DataSource == null) return;

            // Loop through the columns and add their HeaderText or Name
            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {//test
                string columnName = column.Name?.ToLower() ?? "";
                if (columnName.Contains("file") ||
                    columnName.Contains("transaction_number")||
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
            var lastDocumentId = 0;
            var documents = this.documentsTableAdapter.GetData();
            if (documents != null && documents.Count > 0)
            {
                 lastDocumentId = int.Parse(documents.Last().document_id.ToString());
                // استخدم lastDocumentId هنا
            }
            else
            {
                // مفيش بيانات - ممكن تدي قيمة افتراضية
                 lastDocumentId = -1; // أو 0 حسب اللي محتاجه
            }
            var row = this.projectsTableAdapter.GetData()
            .OrderByDescending(r => Convert.ToInt32(r.project_id))
            .FirstOrDefault();

            int lastId = row != null ? Convert.ToInt32(row.project_id) : 0;
            int newId = lastId + 1;
            var idNext = int.Parse(lastDocumentId.ToString());
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
                Architectural_and_Structural_Board_COM.Text,
                Consultant_Surveying_COM.Text,
                Reconciliation_Form_Stamp_COM.Text,
                Secured_certificate_COM.Text
                );
            foreach (var entry in approvalFiles)
            {
                string approvalID = entry.Key;
                if (approvalID == "Petroleum_Ministry_COB") { approvalID = "APP003"; }
                if (approvalID == "Civil_Defense_COB") { approvalID = "APP001"; }
                if (approvalID == "Model_8_COB") { approvalID = "APP006"; }
                if (approvalID == "Traffic_Study_COB") { approvalID = "APP005"; }
                if (approvalID == "Civil_Aviation_COB") { approvalID = "APP004"; }
                if (approvalID == "Environmental_COB") { approvalID = "APP002"; }
                if (approvalID == "Transaction_number_TB") { approvalID = "APP007"; }
                List<string> paths = entry.Value;
                foreach (var path in paths)
                {
                    this.documentsTableAdapter.Insert(
                        (++idNext).ToString(), 
                        approvalID,
                        newId.ToString(),
                        path 
                    );
                }
            }
            LoadProjectData();
            ShowAlert("تمت الاضافه بنجاح",AlertForm.AlertType.Success);
            UpdateRowCount();
        }
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
                approvalFiles.Remove("Petroleum_Ministry_COB");
            }
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            flowLayoutPanel2.Controls.Clear();
            if (approvalFiles.ContainsKey("Civil_Defense_COB"))
            {
                approvalFiles.Remove("Civil_Defense_COB");
            }
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            if (approvalFiles.ContainsKey("Model_8_COB"))
            {
                approvalFiles.Remove("Model_8_COB");
            }
        }

        private void guna2ImageButton7_Click(object sender, EventArgs e)
        {
            flowLayoutPanel4.Controls.Clear();
            if (approvalFiles.ContainsKey("Traffic_Study_COB"))
            {
                approvalFiles.Remove("Traffic_Study_COB");
            }
        }

        private void guna2ImageButton11_Click(object sender, EventArgs e)
        {
            flowLayoutPanel6.Controls.Clear();
            if (approvalFiles.ContainsKey("Civil_Aviation_COB"))
            {
                approvalFiles.Remove("Civil_Aviation_COB");
            }
        }

        private void guna2ImageButton9_Click(object sender, EventArgs e)
        {
            flowLayoutPanel5.Controls.Clear();
            if (approvalFiles.ContainsKey("Environmental_COB"))
            {
                approvalFiles.Remove("Environmental_COB");
            }
        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            flowLayoutPanel7.Controls.Clear();
            if (approvalFiles.ContainsKey("Transaction_number_TB"))
            {
                approvalFiles.Remove("Transaction_number_TB");
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
            HashSet<string> uniqueLands = new HashSet<string>();
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Visible && !row.IsNewRow)
                {
                    var value = row.Cells["Land_Id"].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        uniqueLands.Add(value);
                    }
                }
            }
            rowCountLabel.Text = $"عدد الانشطة: {count}";
            RowCountSeries.Text = $"عدد المشاريع: {uniqueLands.Count}";
        }
        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            var grid = sender as Zuby.ADGV.AdvancedDataGridView;
            Properties.Settings.Default.LastSort = grid.SortString;
            Properties.Settings.Default.Save();
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
        }

        private void guna2ImageButton4_Click(object sender, EventArgs e)
        {
            Files("Transaction_number_TB");
        }

        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ensure click is not on header row
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var clickedColumn = advancedDataGridView1.Columns[e.ColumnIndex];

            // Check if the clicked column is the "pdf" column
            if (clickedColumn.Name.Contains("Status")|| clickedColumn.Name== "Transaction_number")
            {
                string filePath="";
                // Get the file path from the cell value
                if (clickedColumn.Name== "CivilDefenseStatus") {  filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["CivilDefenseFile"].Value?.ToString(); }
                if (clickedColumn.Name == "EnvironmentalStatus") {  filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["EnvironmentalFile"].Value?.ToString(); }
                if (clickedColumn.Name == "PetroleumStatus") { filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["PetroleumFile"].Value?.ToString(); }
                if (clickedColumn.Name == "AviationStatus") { filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["AviationFile"].Value?.ToString(); }
                if (clickedColumn.Name == "TrafficStudyStatus") { filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["TrafficStudyFile"].Value?.ToString(); }
                if (clickedColumn.Name == "Model8Status") { filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["Model8File"].Value?.ToString(); }
                if(clickedColumn.Name== "Transaction_number") { filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["Transaction_numberFile"].Value?.ToString(); }
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
                try
                {
                    if (e.RowIndex >= 0)
                    {
                        string serial = advancedDataGridView1.Rows[e.RowIndex].Cells["project_id"].Value.ToString();
                        serial_number_TB.Text = serial;
                        LoadProjectData(serial);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }
        private void LoadProjectData(string serialNumber)
        {
            bool hasText = !string.IsNullOrWhiteSpace(serialNumber);
            Error_Serial.Visible = !hasText;

            var x = this.projectsTableAdapter.GetDataByIDProjects(serialNumber);

            var gCivil_Defense = this.documentsTableAdapter.GetDataByProjectDoc(serialNumber, "APP001");
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

                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel2.Controls.Clear();
                flowLayoutPanel3.Controls.Clear();
                flowLayoutPanel4.Controls.Clear();
                flowLayoutPanel5.Controls.Clear();
                flowLayoutPanel6.Controls.Clear();
                flowLayoutPanel7.Controls.Clear();
                return;
            }

            var l = this.landsTableAdapter.GetDataBySerial(x.First().land_fk);
            Investment_Name_TB.Text = x.First().project_name;
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
            string filePathPetroleum_Ministry = gPetroleum_Ministry.FirstOrDefault()?.paths?.ToString();
            string filePathCivil_Aviation = gCivil_Aviation.FirstOrDefault()?.paths?.ToString();
            string filePathTraffic_Study = gTraffic_Study.FirstOrDefault()?.paths?.ToString();
            string filePathModel_8 = gModel_8.FirstOrDefault()?.paths?.ToString();
            string filePathTransaction_number = gTransaction_number.FirstOrDefault()?.paths?.ToString();

            if (!string.IsNullOrWhiteSpace(filePathEnvironmental))
            {
                string FileName = Path.GetFileName(filePathEnvironmental);
                if (!approvalFiles.ContainsKey(filePathEnvironmental))
                {
                    approvalFiles["Environmental_COB"] = new List<string>();
                    approvalFiles["Environmental_COB"].Add(filePathEnvironmental);
                }
                AddFileIconToPanel(filePathEnvironmental, FileName, "Environmental_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathPetroleum_Ministry))
            {
                string FileName = Path.GetFileName(filePathPetroleum_Ministry);
                if (!approvalFiles.ContainsKey(filePathPetroleum_Ministry))
                {
                    approvalFiles["Petroleum_Ministry_COB"] = new List<string>();
                    approvalFiles["Petroleum_Ministry_COB"].Add(filePathPetroleum_Ministry);
                }
                AddFileIconToPanel(filePathPetroleum_Ministry, FileName, "Petroleum_Ministry_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathCivil_Defense))
            {
                string FileName = Path.GetFileName(filePathCivil_Defense);
                if (!approvalFiles.ContainsKey(filePathCivil_Defense))
                {
                    approvalFiles["Civil_Defense_COB"] = new List<string>();
                    approvalFiles["Civil_Defense_COB"].Add(filePathCivil_Defense);
                }
                AddFileIconToPanel(filePathCivil_Defense, FileName, "Civil_Defense_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathCivil_Aviation))
            {
                string FileName = Path.GetFileName(filePathCivil_Aviation);
                if (!approvalFiles.ContainsKey(filePathCivil_Aviation))
                {
                    approvalFiles["Civil_Aviation_COB"] = new List<string>();
                    approvalFiles["Civil_Aviation_COB"].Add(filePathCivil_Aviation);
                }
                AddFileIconToPanel(filePathCivil_Aviation, FileName, "Civil_Aviation_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathTraffic_Study))
            {
                string FileName = Path.GetFileName(filePathTraffic_Study);
                if (!approvalFiles.ContainsKey(filePathTraffic_Study))
                {
                    approvalFiles["Traffic_Study_COB"] = new List<string>();
                    approvalFiles["Traffic_Study_COB"].Add(filePathTraffic_Study);
                }
                AddFileIconToPanel(filePathTraffic_Study, FileName, "Traffic_Study_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathModel_8))
            {
                string FileName = Path.GetFileName(filePathModel_8);
                if (!approvalFiles.ContainsKey(filePathModel_8))
                {
                    approvalFiles["Model_8_COB"] = new List<string>();
                    approvalFiles["Model_8_COB"].Add(filePathEnvironmental);
                }
                AddFileIconToPanel(filePathModel_8, FileName, "Model_8_COB");
            }
            if (!string.IsNullOrWhiteSpace(filePathTransaction_number))
            {
                string FileName = Path.GetFileName(filePathTransaction_number);
                if (!approvalFiles.ContainsKey(filePathTransaction_number))
                {
                    approvalFiles["Transaction_number_TB"] = new List<string>();
                    approvalFiles["Transaction_number_TB"].Add(filePathTransaction_number);
                }
                AddFileIconToPanel(filePathTransaction_number, FileName, "Transaction_number_TB");
            }

            Traffic_Study_COB.SelectedItem = x.First().Traffic_Study_Status.ToString() == "X" ? "X" : "✔";
            Petroleum_Ministry_COB.SelectedItem = x.First().Petroleum_Ministry_Approval_status.ToString() == "X" ? "X" : "✔";
            Environmental_COB.SelectedItem = x.First().Environmental_Approval_status.ToString() == "X" ? "X" : "✔";
            Model_8_COB.SelectedItem = x.First().Model_8_Status.ToString() == "X" ? "X" : "✔";
            Civil_Aviation_COB.SelectedItem = x.First().Civil_Aviation_Approval_status.ToString() == "X" ? "X" : "✔";
            Civil_Defense_COB.SelectedItem = x.First().Civil_Defense_Approval_status.ToString() == "X" ? "X" : "✔";
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
            // advancedDataGridView1.FilterString = Properties.Settings.Default.LastFilter;
            //advancedDataGridView1.SortString = Properties.Settings.Default.LastSort;
            loadcomboxes();
            LoadProjectData();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
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
            var igover = this.governorateTableAdapter.GetDataByGovernorate(Governorate_COB.Text);
            var x = this.projectsTableAdapter.GetDataByIDProjects(serial_number_TB.Text);          
            if (x == null || x.Count == 0)
            {
                ShowAlert("لا يوجد هذا البيان للتعديل", AlertForm.AlertType.Error);
                return;
            }
            this.projectsTableAdapter.UpdateQuery(Investment_Name_TB.Text,
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
                serial_number_TB.Text
                );
            foreach (var entry in approvalFiles)
            {
                string approvalID = entry.Key;
                if (approvalID == "Petroleum_Ministry_COB") { approvalID = "APP003"; }
                if (approvalID == "Civil_Defense_COB") { approvalID = "APP001"; }
                if (approvalID == "Model_8_COB") { approvalID = "APP006"; }
                if (approvalID == "Traffic_Study_COB") { approvalID = "APP005"; }
                if (approvalID == "Civil_Aviation_COB") { approvalID = "APP004"; }
                if (approvalID == "Environmental_COB") { approvalID = "APP002"; }
                if (approvalID == "Transaction_number_TB") { approvalID = "APP007"; }
                List<string> paths = entry.Value;
                foreach (var path in paths)
                {
                    this.documentsTableAdapter.UpdateQuery(
                        approvalID,
                        serial_number_TB.Text,
                        path
                    );
                }
            }
            LoadProjectData();
            ShowAlert("تم التعديل بنجاح", AlertForm.AlertType.Success);
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
                    string EnvironmentalFileValue = row.Cells["EnvironmentalFile"].Value.ToString();
                    string PetroleumFileValue = row.Cells["PetroleumFile"].Value.ToString();
                    string AviationFileValue = row.Cells["AviationFile"].Value.ToString();
                    string TrafficStudyFileValue = row.Cells["TrafficStudyFile"].Value.ToString();
                    string Model8FileValue = row.Cells["Model8File"].Value.ToString();
                    string Transaction_numberValue = row.Cells["Transaction_numberFile"].Value.ToString();

                    if (!string.IsNullOrWhiteSpace(CivilDefenseFileValue))
                        this.documentsTableAdapter.DeleteQuery("APP001", projectId);

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
            { "مسلسل القطعة", "Land_Id" },
            { "مسلسل المشروع", "Project_Id" },
            { "رقم اللوحة", "PlateNumber" },
            { "اسم قطعة الأرض", "LandName" },
            { "رقم قطعة الأرض", "Land_Number" },
            { "اسم المشروع", "ProjectName" },
            { "اسم المحافظة", "GovernorateName" },
            { "حالة موافقة الحماية المدنية", "CivilDefenseStatus" },
            { "حالة موافقة البيئة", "EnvironmentalStatus" },
            { "حالة موافقة وزارة البترول", "PetroleumStatus" },
            { "حالة موافقة الطيران المدني", "AviationStatus" },
            { "حالة الدراسة المرورية", "TrafficStudyStatus" },
            { "حالة نموذج 8 أو 10", "Model8Status" },
            { "رقم المعاملة", "Transaction_number" },
            { "تاريخ انتهاء العقد", "Contract_expiry_date" },
            { "اجمالي محلات", "Total_stores" },
            { "مؤجر", "Total_rented" },
            { "غير مؤجر", "Total_Not_rented" },
            { "الشهادة المؤمنه", "Secured_Certificate" },
            { "لوحة المعماري والانشائي", "Architectural_and_Structural_Board" },
            { "ختم نموذج التصالح", "Reconciliation_Form_Stamp" },
            { "الرفع المساحي الاستشاري", "Consultant_Surveying" },
            { "المكتب الاستشاري", "consulting_Office" }
        };
        private void skyButton4_Click(object sender, EventArgs e)
        {
            reportViewer1.Visible = true;
            reportViewer1.LocalReport.DataSources.Clear();
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

        private void advancedDataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            advancedDataGridView1.Invalidate();
        }
    }
}
