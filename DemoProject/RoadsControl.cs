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
    public partial class RoadsControl : UserControl
    {
        public long UserId;
        public void SetUserData(string user, long id)
        {

            this.UserId = id;    // Or store it in a field/property
        }
        public RoadsControl()
        {
            InitializeComponent();
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

        private void skyButton3_Click(object sender, EventArgs e)
        {

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
        public static DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            var properties = typeof(T).GetProperties()
                .Where(p => p.CanRead &&
                            p.PropertyType.Namespace == "System" && // Avoid complex nested types
                            p.DeclaringType == typeof(T))          // Skip inherited properties like DataRow stuff
                .ToArray();

            foreach (var prop in properties)
            {
                dataTable.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (var item in items)
            {
                var values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }
        public void translateToArabic()
        {
            advancedDataGridView1.Columns["council_of_ministers_decision_Date"].HeaderText = "تاريخ قرار مجلس الوزراء";
            advancedDataGridView1.Columns["franchise_Contract_Duration"].HeaderText = "مدة عقد منح الالتزام";
            advancedDataGridView1.Columns["contract_signing_date"].HeaderText = "تاريخ توقع العقد";
            advancedDataGridView1.Columns["road_id"].HeaderText = "مسلسل";
            advancedDataGridView1.Columns["road_name"].HeaderText = "اسم الطريق";
            advancedDataGridView1.Columns["Administrative_Affiliation"].HeaderText = "التبعية الادارية";
            advancedDataGridView1.Columns["Financial_Affiliation"].HeaderText = "التبعية المالية";
            advancedDataGridView1.Columns["Toll_Booth_Count"].HeaderText = "عدد منافذ";
            advancedDataGridView1.Columns["contract_type"].HeaderText = "نوع العقد";
            advancedDataGridView1.Columns["contract_status"].HeaderText = "حالة العقد";
            advancedDataGridView1.Columns["council_of_ministers_decision_Image"].HeaderText = "صورة قرار مجلس الوزراء";
            advancedDataGridView1.Columns["council_of_ministers_decision_Image"].Visible = false;
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
                // --- Part 1: Fetch & transform data in background ---
                var result = await Task.Run(() =>
                {
                    var query = roadsTableAdapter.GetData();
                    var joinedList = query.ToList();

                    // Convert to DataTable
                    DataTable original = ToDataTable(joinedList);

                    // Clone & fix Date columns
                    DataTable converted = original.Clone();
                    converted.Columns["council_of_ministers_decision_Date"].DataType = typeof(DateTime);
                    converted.Columns["franchise_Contract_Duration"].DataType = typeof(DateTime);
                    converted.Columns["contract_signing_date"].DataType = typeof(DateTime);

                    foreach (DataRow row in original.Rows)
                    {
                        var newRow = converted.NewRow();
                        foreach (DataColumn col in original.Columns)
                        {
                            if (col.ColumnName == "council_of_ministers_decision_Date" ||
                                col.ColumnName == "franchise_Contract_Duration" ||
                                col.ColumnName == "contract_signing_date")
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

                    return converted;
                });

                // --- Part 2: Bind to UI (on UI thread) ---
                BindingSource bindingSource = new BindingSource();
                bindingSource.DataSource = result;
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

                // Call translations after binding
                translateToArabic();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "خطأ غير متوقع", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            FalseFunction();
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
            guna2Button3.Location = new Point(headerRight - 5, headerTop);
        }
        private async void _ٌRoadsControl_Load(object sender, EventArgs e)
        {
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            await LoadSourceDataAsync();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
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

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_ItemCheck_1(object sender, ItemCheckEventArgs e)
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
   
        private void guna2Button1_Click(object sender, EventArgs e)
        {
            var data = this.roadsTableAdapter.GetData().FindByroad_id(serial_number_TB.Text);
            if (data !=null)
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

            if (string.IsNullOrWhiteSpace(Road_Name_TB.Text))
            {
                Error_road.Visible = true;
                Road_Name_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Toll_Bath_Count_TB.Text))
            {
                Error_Toll.Visible = true;
                Toll_Bath_Count_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Administrative_Affiliation_TB.Text))
            {
                Error_Administrative.Visible = true;
                Administrative_Affiliation_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Financial_Affiliation_TB.Text))
            {
                Error_Financial.Visible = true;
                Financial_Affiliation_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(council_of_ministers_decision_TB.Text))
            {
                council_of_ministers_decision_Error.Visible = true;
                contract_signing_date_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(number_of_exits_TB.Text))
            {
                number_of_exits_Error.Visible = true;
                number_of_exits_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(contract_status_COB.Text))
            {
                Error_contract_status.Visible = true;
                contract_status_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(contract_type_COB.Text))
            {
                Error_contract_type.Visible = true;
                contract_type_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(road_length_TB.Text))
            {
                road_length_Error.Visible = true;
                road_length_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(road_authority_percentage_TB.Text))
            {
                road_authority_percentage_Error.Visible = true;
                road_authority_percentage_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(company_percentage_TB.Text))
            {
                company_percentage_Error.Visible = true;
                company_percentage_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(number_of_operation_staff_TB.Text))
            {
                number_of_operation_staff_Error.Visible = true;
                number_of_operation_staff_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(road_length_including_branches_TB.Text))
            {
                road_length_including_branches_Error.Visible = true;
                road_length_including_branches_TB.Focus();
                hasError = true;
            }
            if (hasError)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
                return;
            }
            else
            {
                int lastId = int.Parse(this.roadsTableAdapter.GetData()
               .OrderByDescending(r => Convert.ToInt32(r.road_id))
               .First()
               .road_id);

                int newId = lastId + 1;
                roadsTableAdapter.Insert(
                    newId.ToString(),
                    Road_Name_TB.Text,
                    Toll_Bath_Count_TB.Text,
                    Administrative_Affiliation_TB.Text,
                    Financial_Affiliation_TB.Text,
                    contract_type_COB.Text,
                    contract_status_COB.Text,
                    council_of_ministers_decision_TB.Text,
                    council_of_ministers_decision_Date_TB.Value,
                    contract_signing_date_TB.Value,
                    franchise_Contract_Duration_TB.Value,
                    number_of_exits_TB.Text,
                    road_length_TB.Text,
                    road_length_including_branches_TB.Text,
                    number_of_operation_staff_TB.Text,
                    company_percentage_TB.Text,
                    road_authority_percentage_TB.Text,
                    approvalFiles.TryGetValue("contract_type_COB", out var filePathcontract) ? filePathcontract.ToString() : null,
                    approvalFiles.TryGetValue("council_of_ministers_decision_TB", out var filePathdecision) ? filePathdecision.ToString() : null
                    );
                var query = this.roadsTableAdapter.GetData();
                advancedDataGridView1.DataSource = query;
                ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
                UpdateRowCount();
            }

        }

        private void serial_number_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(serial_number_TB.Text))
            {
                Error_Serial.Visible = false;
            }
            var x = this.roadsTableAdapter.GetDataByID(serial_number_TB.Text);
            if (x == null || x.Count==0)
            {
                // No data found — clear all fields
               // serial_number_TB.Text = "";
                Road_Name_TB.Text = "";
                Toll_Bath_Count_TB.Text = "";
                Administrative_Affiliation_TB.Text = "";
                Financial_Affiliation_TB.Text = "";
                council_of_ministers_decision_TB.Text = "";
                number_of_exits_TB.Text = "";
                road_length_TB.Text = "";
                company_percentage_TB.Text = "";
                road_authority_percentage_TB.Text = "";
                number_of_operation_staff_TB.Text = "";
                road_length_including_branches_TB.Text = "";
                contract_status_COB.SelectedIndex = -1;
                contract_type_COB.SelectedIndex = -1;
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel2.Controls.Clear();
                return;
            }
            serial_number_TB.Text = x.First().road_id;
            Road_Name_TB.Text = x.First().road_name;
            Toll_Bath_Count_TB.Text = x.First().Toll_Booth_Count;
            Administrative_Affiliation_TB.Text = x.First().Administrative_Affiliation;
            Financial_Affiliation_TB.Text = x.First().Financial_Affiliation;
            council_of_ministers_decision_TB.Text = x.First().council_of_ministers_decision;
            number_of_exits_TB.Text = x.First().number_of_exits;
            road_length_TB.Text = x.First().road_length;
            company_percentage_TB.Text = x.First().company_percentage;
            road_authority_percentage_TB.Text = x.First().road_authority_percentage;
            number_of_operation_staff_TB.Text = x.First().number_of_operation_staff;
            road_length_including_branches_TB.Text = x.First().road_length_including_branches;
            contract_status_COB.Text = x.First().contract_status.ToString() == "X" ? "X" : "✔";
            contract_type_COB.Text = x.First().contract_type.ToString() == "X" ? "X" : "✔"; ;
            string filePathContract = x.First().contract_image?.ToString();
            string filePathCouncil = x.First().council_of_ministers_decision_Image?.ToString();
            if (!string.IsNullOrWhiteSpace(filePathContract))
            {
                string contractFileName = Path.GetFileName(filePathContract);

                if (!approvalFiles.ContainsKey(filePathContract))
                {
                    approvalFiles["contract_type_COB"] = new List<string>();
                    approvalFiles["contract_type_COB"].Add(filePathContract);
                }
                AddFileIconToPanel(filePathContract, contractFileName, "contract_type_COB");
            }

            // Add council_of_ministers_decision_Image to panel 2
            if (!string.IsNullOrWhiteSpace(filePathCouncil))
            {
                string councilFileName = Path.GetFileName(filePathCouncil);

                if (!approvalFiles.ContainsKey(filePathCouncil))
                {
                    approvalFiles["council_of_ministers_decision_TB"] = new List<string>();
                    approvalFiles["council_of_ministers_decision_TB"].Add(filePathCouncil);

                }
                AddFileIconToPanel(filePathCouncil, councilFileName, "council_of_ministers_decision_TB");
            }

        }

        private void Road_Name_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Road_Name_TB.Text))
            {
                Error_road.Visible = false;
            }
        }

        private void Toll_Bath_Count_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Toll_Bath_Count_TB.Text))
            {
                Error_Toll.Visible = false;
            }
        }

        private void Administrative_Affiliation_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Administrative_Affiliation_TB.Text))
            {
                Error_Administrative.Visible = false;
            }
        }

        private void Financial_Affiliation_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(Financial_Affiliation_TB.Text))
            {
                Error_Financial.Visible = false;
            }
        }

        private void contract_signing_date_TB_TextChanged(object sender, EventArgs e)
        {
           /* if (!string.IsNullOrWhiteSpace(contract_signing_date_TB.Text))
            {
                Error_contract_signing.Visible = false;
            }*/
        }

        private void contract_status_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(contract_status_COB.Text))
            {
                Error_contract_status.Visible = false;
            }
        }

        private void contract_type_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(contract_type_COB.Text))
            {
                Error_contract_type.Visible = false;
            }
        }

        private void franchise_Contract_Duration_TB_TextChanged(object sender, EventArgs e)
        {
           /* if (!string.IsNullOrWhiteSpace(franchise_Contract_Duration_TB.Text))
            {
                Error_franchise_Contract.Visible = false;
            }*/
        }

        private void skyButton2_Click_1(object sender, EventArgs e)
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

        private void serial_number_TB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // Block the input
            }
        }

        private void skyButton3_Click_1(object sender, EventArgs e)
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

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void nightLabel1_Click(object sender, EventArgs e)
        {

        }

        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            Files("council_of_ministers_decision_TB");
        }
        private Dictionary<string, List<string>> approvalFiles = new Dictionary<string, List<string>>();// at class level
        public void Files(string approvalKey)
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
                case "contract_type_COB":
                    flowLayoutPanel2.Controls.Add(container);
                    break;
                case "council_of_ministers_decision_TB":
                    flowLayoutPanel1.Controls.Add(container);
                    break;
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

        private void guna2ImageButton4_Click(object sender, EventArgs e)
        {
            Files("contract_type_COB");
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            if (approvalFiles.ContainsKey("council_of_ministers_decision_TB"))
            {
                approvalFiles.Remove("council_of_ministers_decision_TB");
            }
        }

        private void guna2ImageButton3_Click(object sender, EventArgs e)
        {
            flowLayoutPanel2.Controls.Clear();
            if (approvalFiles.ContainsKey("contract_type_COB"))
            {
                approvalFiles.Remove("contract_type_COB");
            }
        }

        private void council_of_ministers_decision_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(council_of_ministers_decision_TB.Text))
            {
                council_of_ministers_decision_Error.Visible = false;
            }
        }

        private void number_of_exits_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(number_of_exits_TB.Text))
            {
                number_of_exits_Error.Visible = false;
            }
        }

        private void road_length_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(road_length_TB.Text))
            {
                road_length_Error.Visible = false;
            }
        }

        private void road_authority_percentage_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(road_authority_percentage_TB.Text))
            {
                road_authority_percentage_Error.Visible = false;
            }

        }

        private void company_percentage_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(company_percentage_TB.Text))
            {
                company_percentage_Error.Visible = false;
            }

        }

        private void number_of_operation_staff_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(number_of_operation_staff_TB.Text))
            {
                number_of_operation_staff_Error.Visible = false;
            }
        }

        private void road_length_including_branches_TB_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(road_length_including_branches_TB.Text))
            {
                road_length_including_branches_Error.Visible = false;
            }
        }

        private async void guna2Button2_Click(object sender, EventArgs e)
        {
            var x = this.roadsTableAdapter.GetDataByID(serial_number_TB.Text);
            if (x == null || x.Count == 0)
            {
                ShowAlert("No data to update it", AlertForm.AlertType.Error);
                return;
            }

            string filePathcouncil = (approvalFiles != null && approvalFiles.Count > 0)
                           ? approvalFiles["council_of_ministers_decision_TB"].ToString()
                           : null;
            string filePathContract = (approvalFiles != null && approvalFiles.Count > 0)
                           ? approvalFiles["contract_type_COB"].ToString()
                           : null;

            this.roadsTableAdapter.UpdateQuery(
                Road_Name_TB.Text,
                Toll_Bath_Count_TB.Text,
                Administrative_Affiliation_TB.Text,
                Financial_Affiliation_TB.Text,
                contract_type_COB.Text,
                contract_status_COB.Text,
                council_of_ministers_decision_TB.Text,
                contract_signing_date_TB.Value.ToString(),
                contract_signing_date_TB.Value.ToString(),
                franchise_Contract_Duration_TB.Value.ToString(),
                number_of_exits_TB.Text,
                road_length_TB.Text,
                road_length_including_branches_TB.Text,
                number_of_operation_staff_TB.Text,
                company_percentage_TB.Text,
                road_authority_percentage_TB.Text,
                filePathContract,
                filePathcouncil,
                serial_number_TB.Text
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
            DialogResult result = MessageBox.Show(
                "هل أنت متأكد من أنك تريد حذف الطرق المحددة؟",
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
                    string roadId = row.Cells["road_id"].Value.ToString();
                    this.roadsTableAdapter.DeleteQuery(roadId);
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

        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;

            // Update the row count
            UpdateRowCount();
        }

        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            BindingSource b = new BindingSource();
            b.DataSource = advancedDataGridView1.DataSource;
            b.Filter = advancedDataGridView1.FilterString;
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
