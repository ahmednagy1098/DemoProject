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
using System.Globalization;
using System.Text.RegularExpressions;
using System.Diagnostics;
namespace DemoProject
{
    public partial class InvestmentsControl : UserControl
    {
        public long UserId;
        public string UserName;
        public string Investment_type;
        private DataTable originalData;
        private bool _skipCloseConfirmation = false;
        public void SetUserData(string user, long id, string investment_type)
        {
            this.UserName = user;
            this.UserId = id;    // Or store it in a field/property
            this.Investment_type = investment_type;
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
            HighlightExpiryCells();
        }

        private void UpdateRowCount()
        {
            int count = 0; // total visible rows
            HashSet<string> uniqueBridgeNeighborhoods = new HashSet<string>(); // to track unique bridge areas
            HashSet<string> Roads_ID_Projects = new HashSet<string>();
            // Example: show counts in labels
            rowCountLabel.Text = $"عدد الصفوف: {count.ToString()}";
            if (SessionData.Investment_Type == "استثمارات على الطرق")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                        string investmentId = row.Cells["investments_id"].Value?.ToString();
                        string Project_ID = row.Cells["Project_Id"].Value?.ToString();

                        // Check if it's a bridge (ID starts with #)
                        if (!string.IsNullOrEmpty(investmentId) && investmentId.StartsWith("&"))
                        {
                            // Add neighborhood to HashSet (duplicates ignored automatically)
                            if (!string.IsNullOrEmpty(Project_ID))
                                Roads_ID_Projects.Add(Project_ID);
                        }
                    }
                }
                int countRoads = Roads_ID_Projects.Count;
                labelBridges.Text = $"عدد الطرق:{countRoads.ToString()}";
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";


            }
            if (SessionData.Investment_Type == "اسفل كباري")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                        string investmentId = row.Cells["investments_id"].Value?.ToString();
                        string neighborhood = row.Cells["Name_Projects"].Value?.ToString();

                        // Check if it's a bridge (ID starts with #)
                        if (!string.IsNullOrEmpty(investmentId) && investmentId.StartsWith("#"))
                        {
                            // Add neighborhood to HashSet (duplicates ignored automatically)
                            if (!string.IsNullOrEmpty(neighborhood))
                                uniqueBridgeNeighborhoods.Add(neighborhood);
                        }
                    }
                }
                int countBridges = uniqueBridgeNeighborhoods.Count;
                labelBridges.Text = $"عدد الكباري:{countBridges.ToString()}";
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
            }
            if (SessionData.Investment_Type == "مواقف")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                        string investmentId = row.Cells["investments_id"].Value?.ToString();
                        string neighborhood = row.Cells["Name_Projects"].Value?.ToString();

                        // Check if it's a bridge (ID starts with #)
                        if (!string.IsNullOrEmpty(investmentId) && investmentId.StartsWith("م"))
                        {
                            // Add neighborhood to HashSet (duplicates ignored automatically)
                            if (!string.IsNullOrEmpty(neighborhood))
                                uniqueBridgeNeighborhoods.Add(neighborhood);
                        }
                    }
                }
                int countBridges = uniqueBridgeNeighborhoods.Count;
                labelBridges.Text = $"عدد المواقف:{countBridges.ToString()}";
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
            }
            if (SessionData.Investment_Type == "مناطق تنموية")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                        string investmentId = row.Cells["investments_id"].Value?.ToString();
                        string neighborhood = row.Cells["Name_Projects"].Value?.ToString();

                        // Check if it's a bridge (ID starts with #)
                        if (!string.IsNullOrEmpty(investmentId) && investmentId.StartsWith("@"))
                        {
                            // Add neighborhood to HashSet (duplicates ignored automatically)
                            if (!string.IsNullOrEmpty(neighborhood))
                                uniqueBridgeNeighborhoods.Add(neighborhood);
                        }
                    }
                }
                int countBridges = uniqueBridgeNeighborhoods.Count;
                labelBridges.Text = $"عدد المناطق:{countBridges.ToString()}";
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
            }
            if (SessionData.Investment_Type == "مولات")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                        string investmentId = row.Cells["investments_id"].Value?.ToString();
                        string neighborhood = row.Cells["Dependent_neighborhood"].Value?.ToString();

                        // Check if it's a bridge (ID starts with #)
                        if (!string.IsNullOrEmpty(investmentId) && investmentId.Contains("MA"))
                        {
                            // Add neighborhood to HashSet (duplicates ignored automatically)
                            if (!string.IsNullOrEmpty(neighborhood))
                                uniqueBridgeNeighborhoods.Add(neighborhood);
                        }
                    }
                }
                int countBridges = uniqueBridgeNeighborhoods.Count;
                labelBridges.Text = $"عدد المولات:{countBridges.ToString()}";
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
            }
            if (SessionData.Investment_Type == "محلات شل اوت خارج")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                    }
                }
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
                labelBridges.Text = "";
            }
            if (SessionData.Investment_Type == "" || SessionData.Investment_Type == null)
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;
                    }
                }

                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
                labelBridges.Text = "";
            }
            if (SessionData.Investment_Type == "محلات شل اوت داخل")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                    }
                }
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
                labelBridges.Text = "";
            }
            if (SessionData.Investment_Type == "كل الاستثمارات")
            {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        count++;

                        string investmentId = row.Cells["investments_id"].Value?.ToString();
                        string Project_ID = row.Cells["Name_Projects"].Value?.ToString();

                        // Check if it's a bridge (ID starts with #)
                        if (!string.IsNullOrEmpty(investmentId) )
                        {
                            // Add neighborhood to HashSet (duplicates ignored automatically)
                            if (!string.IsNullOrEmpty(Project_ID))
                                Roads_ID_Projects.Add(Project_ID);
                        }
                    }
                }
                int countRoads = Roads_ID_Projects.Count;
                labelBridges.Text = $"عدد الاماكن:{countRoads.ToString()}";
                rowCountLabel.Text = $"عدد الصفوف : {count.ToString()}";
                bigLabel2.Visible = false;
            }
            CalculateTotalShops();
            CalculateUniqueNeighborhoods();
            CalculateRentedPDFS();
            HighlightExpiryCells();
        }
        private void CalculateTotalShops()
        {
            int total = 0;

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Cells["Shops_Count"].Value != null &&
                    int.TryParse(row.Cells["Shops_Count"].Value.ToString(), out int value))
                {
                    total += value;
                }
            }
            if (SessionData.Investment_Type != "مولات")
            {
                bigLabel1.Text = $"إجمالي عدد المحلات: {total}";
            }
            else
            {
                bigLabel1.Text = $"إجمالي عدد وحدات: {total}";
            }
            
        }
        private ToolTip sharedToolTip = new ToolTip(); // One tooltip shared across form
        private List<string> mixedStationsCache = new List<string>(); // To store mixed stations for the click event

        private void CalculateUniqueNeighborhoods()
        {
            // Store unique station names and map of statuses per station (based on visible rows)
            HashSet<string> uniqueNeighborhoods = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, HashSet<string>> stationStatusMap = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (!row.Visible || row.IsNewRow)
                    continue;

                string neighborhood = row.Cells["Name_Projects"].Value?.ToString()?.Trim();
                string rentalStatus = row.Cells["rental_status"].Value?.ToString()?.Trim();

                if (string.IsNullOrEmpty(neighborhood))
                    continue;

                uniqueNeighborhoods.Add(neighborhood);

                // Build status map for visible data
                if (!stationStatusMap.ContainsKey(neighborhood))
                    stationStatusMap[neighborhood] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (!string.IsNullOrEmpty(rentalStatus))
                    stationStatusMap[neighborhood].Add(rentalStatus);
            }

            // Find mixed-status stations (stations that appear in multiple rental states)
            var mixedStations = stationStatusMap
                .Where(kv => kv.Value.Count > 1)
                .Select(kv => kv.Key)
                .ToList();

            // Keep them for the click event
            mixedStationsCache = mixedStations;

            // Case: محلات شل اوت داخل / خارج
            if (SessionData.Investment_Type == "محلات شل اوت داخل" || SessionData.Investment_Type == "محلات شل اوت خارج")
            {
                var x = this.invesments_ListTableAdapter.GetData();
                int totalRows = x.Rows.Count;

                bigLabel2.Text = $"اجمالي عدد المحطات: {totalRows}";
                labelBridges.Visible = true;
                labelBridges.Text = $"محطات بها استثمار: {uniqueNeighborhoods.Count} ⓘ"; // add small info icon text
                labelBridges.ForeColor = Color.FromArgb(33, 150, 243);
                // ✅ Always clear old tooltip firs
                
                // ✅ Attach click event only once
                labelBridges.Cursor = Cursors.Hand;
                labelBridges.Click -= labelBridges_Click; // avoid duplicate bindings
                labelBridges.Click += labelBridges_Click;
            }

            // Case: استثمارات على الطرق / اسفل كباري
            if (SessionData.Investment_Type == "استثمارات على الطرق" ||
                SessionData.Investment_Type == "اسفل كباري" ||
                SessionData.Investment_Type == "مولات" ||
                SessionData.Investment_Type == "مواقف" ||
                SessionData.Investment_Type == "مناطق تنموية")
            {
                bigLabel2.Visible = false;
            }
        }



        private void CalculateRentedPDFS()
        {
            int rentedCount = 0; // total rows with Rental_Status = "مؤجر"
            if (SessionData.Investment_Type == "محلات شل اوت داخل" ||
                SessionData.Investment_Type == "محلات شل اوت خارج"
                && SessionData.Investment_Type != "كل الاستثمارات") {
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        string rentalStatus = row.Cells["Rental_Status"].Value?.ToString()?.Trim();

                        // ✅ Count rows that have `Rental_Status = "مؤجر"
                        if (rentalStatus == "مؤجر")
                            rentedCount++;
                    }
                }
                bigLabel3.Text = $"عدد العقود: {rentedCount}";
            }

            if (SessionData.Investment_Type != "محلات شل اوت داخل" &&
                SessionData.Investment_Type != "محلات شل اوت خارج"
                && SessionData.Investment_Type != "كل الاستثمارات")
            {
                HashSet<string> countedActivities = new HashSet<string>();
                 rentedCount = 0;

                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && !row.IsNewRow)
                    {
                        string rentalStatus = row.Cells["Rental_Status"].Value?.ToString()?.Trim();
                        string activityName = row.Cells["Activity_Name"].Value?.ToString()?.Trim();

                        // Only count if rented AND activityName is not null/empty AND not already counted
                        if (rentalStatus == "مؤجر" && !string.IsNullOrEmpty(activityName) && !countedActivities.Contains(activityName))
                        {
                            rentedCount++;
                            countedActivities.Add(activityName); // mark this activity as counted
                        }
                    }
                }

                bigLabel3.Text = $"عدد العقود: {rentedCount}";
            }

            if (SessionData.Investment_Type == "كل الاستثمارات")
            {
                rentedCount = 0;

                // HashSet for non "شل اوت" activities (to avoid duplicates)
                HashSet<string> countedActivities = new HashSet<string>();

                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (!row.Visible || row.IsNewRow)
                        continue;

                    string investment_type = row.Cells["investment_type"].Value?.ToString()?.Trim();
                    string rentalStatus = row.Cells["Rental_Status"].Value?.ToString()?.Trim();
                    string activityName = row.Cells["Activity_Name"].Value?.ToString()?.Trim();

                    if (rentalStatus != "مؤجر")
                        continue;

                    // ✅ Case 1: Shell Out investments → count every rented row
                    if (investment_type == "محلات شل اوت داخل" ||
                        investment_type == "محلات شل اوت خارج")
                    {
                        rentedCount++;
                    }
                    // ✅ Case 2: Other investments → count unique Activity_Name only
                    else
                    {
                        if (!string.IsNullOrEmpty(activityName) &&
                            !countedActivities.Contains(activityName))
                        {
                            rentedCount++;
                            countedActivities.Add(activityName);
                        }
                    }
                }

                bigLabel3.Text = $"عدد العقود: {rentedCount}";
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
        // 146 - 147 - 148 - 145 -176 - 406 -235- 236- 1-2-3-4-340-341-
        // 782-783-784 -785 -786 -787 (833 NEED TO BE CHECKED)
        private void LoadRoadsToCheckedListBox()
        {
            if (SessionData.Investment_Type != "مناطق تنموية") {
                var allProjects = projectsTableAdapter.GetData();

                // Get all project names where ID starts with &
                var roads = allProjects
                    .Where(p => !p.Island_fkNull() && p.project_id.StartsWith("&"))
                    .Where(p => !p.IsName_ProjectsNull())
                    .Select(p => p.Name_Projects.Trim())
                    .Distinct()
                    .OrderBy(name => name)
                    .ToList();

                checkedListBox2.Items.Clear();
                foreach (var road in roads)
                    checkedListBox2.Items.Add(road, false);
            }
            if (SessionData.Investment_Type == "مناطق تنموية")
            {
                checkedListBox2.Items.Clear();
                checkedListBox2.Items.Add("مول", false);
                checkedListBox2.Items.Add("غاز", false);
                checkedListBox2.Items.Add("محلات", false);
                checkedListBox2.Items.Add("ملاعب", false);
                checkedListBox2.Items.Add("ساحة", false);
                checkedListBox2.Items.Add("موقف", false);
                checkedListBox2.Items.Add("ملاهي", false);
                checkedListBox2.Items.Add("حدائق", false);
                checkedListBox2.Items.Add("عمارات", false);
                checkedListBox2.Items.Add("مركز صحي", false);
                checkedListBox2.Items.Add("فضاء", false);
            }
        }
        private string GetNameWithoutParentheses(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "";

            int idx = value.IndexOf("(");
            return idx > 0 ? value.Substring(0, idx).Trim() : value.Trim();
        }

        private string GetTextInsideParentheses(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "";

            int start = value.IndexOf("(");
            int end = value.IndexOf(")");

            if (start >= 0 && end > start)
                return value.Substring(start + 1, end - start - 1).Trim();

            return "";
        }
        private int GetYearsDifference(DateTime start, DateTime end)
        {
            if (end < start)
                return 0;

            int years = end.Year - start.Year;

            // لو تاريخ النهاية قبل يوم/شهر البداية في نفس السنة → السنة لم تكتمل
            if (end < start.AddYears(years))
                years--;

            return years;
        }

        private string GetDateDifference(DateTime start, DateTime end)
        {
            if (end < start)
                return "";

            int years = end.Year - start.Year;
            int months = end.Month - start.Month;
            int days = end.Day - start.Day;

            if (days < 0)
            {
                months--;
                var prevMonth = end.AddMonths(-1);
                days += DateTime.DaysInMonth(prevMonth.Year, prevMonth.Month);
            }

            if (months < 0)
            {
                years--;
                months += 12;
            }

            // ✅ تقريب السنة لو الفرق يوم واحد فقط
            var totalDays = (end - start).TotalDays;
            if (years == 0 && totalDays >= 364 && totalDays <= 366)
            {
                years = 1;
                months = 0;
                days = 0;
            }

            return $"{years} سنة {months} شهر {days} يوم";
        }

        public void LoadInvestmentData() // شل اوت خارج - مواقف - مناطق تنموية - كباري
        {
            try
            {
                if (SessionData.Investment_Type != "محلات شل اوت خارج" &&
                    SessionData.Investment_Type != "كل الاستثمارات" && 
                    SessionData.Investment_Type != "مواقف" &&
                   SessionData.Investment_Type != "مناطق تنموية")
                {
                    skyButton6.Visible = true;

                    TrueFunction();

                    // --- Part 1: Fetch and join data ---
                    var query =
                      from I in investmentsTableAdapter.GetData()
                      join g in governorateTableAdapter.GetData()
                          on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                      from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                      join l in landsTableAdapter.GetData()
                          on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                      from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                      join p in projectsTableAdapter.GetData()
                          on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                      from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                      join d in documentsTableAdapter.GetData()
                      on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                      from d in dj
                          .Where(x => x.approvals_fk == "APP009")
                          .DefaultIfEmpty()
                      where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                      select new
                      {
                         
                          Project_Id = p == null ? "" : p.project_id,
                          investments_id = I.investments_id,
                          Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*"))? " " : (l == null ? "" : l.land_id),
                          PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                          Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                          GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                          investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                          //Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                          Name_Projects = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetNameWithoutParentheses(I.Dependent_neighborhood),

                          investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                          //lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                          //Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                          //City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                          //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull()?"": p.Architectural_and_Structural_Board,
                          Dependent_neighborhood = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetTextInsideParentheses(I.Dependent_neighborhood),
                          Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                          Location = I.IsLocationNull() ? "" : I.Location,
                          Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                          //Description_Drawing_Place = I.IsDescription_Drawing_PlaceNull() ? "" : I.Description_Drawing_Place,
                          Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                          Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                          Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                          Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                          Rental_Status =
                            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                    ? (
                                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                          ? "غير مؤجر (فارغ)"
                                          : "منتظر العقد"
                                      )
                                    : "مؤجر"),
                          Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                          Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                          Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                          Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                          Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                          Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                          OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                          ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                          Notes = I.IsNotesNull() ? "" : I.Notes,
                          Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                      };

                    var joinedList = query
                     .Where(r => r.Visible == true)
                     .GroupBy(r => r.investments_id)
                     .Select(g => g.First()) // pick one of each investment_id
                     .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                     .ToList();
                    DataTable original = ToDataTable(joinedList);

                    // --- Convert string date columns to DateTime ---
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
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.ValueType == typeof(DateTime))
                        {
                            col.DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }

                    // --- Bind data to DataGridView ---
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
                    
                    originalData = converted;
                }

                if (SessionData.Investment_Type == "محلات شل اوت خارج" && SessionData.Investment_para1 == "شل اوت الطرق")
                {

                    skyButton6.Visible = false;
                    TrueFunction();

                    // --- Part 1: Fetch and join data ---
                    var query =
                      from I in investmentsTableAdapter.GetData()
                      join g in governorateTableAdapter.GetData()
                          on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                      from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                      join l in landsTableAdapter.GetData()
                          on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                      from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                      join p in projectsTableAdapter.GetData()
                          on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                      from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                      join d in documentsTableAdapter.GetData()
                      on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                      from d in dj
                          .Where(x => x.approvals_fk == "APP009")
                          .DefaultIfEmpty()
                      where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                              && (!I.IsDependent_neighborhoodNull() && I.Dependent_neighborhood.Contains("*"))
                      select new
                      {
                          Project_Id = p == null ? "" : p.project_id,
                          investments_id = I.investments_id,
                          Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                          PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                          Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                          lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                          Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                          City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                          Address = l.IsAddressNull() ? "" : l.Address,
                          GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                          investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                          Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                          investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                          Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                          Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                          Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                          Location = I.IsLocationNull() ? "" : I.Location,
                          Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                          //Description_Drawing_Place = I.IsDescription_Drawing_PlaceNull() ? "" : I.Description_Drawing_Place,
                          Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                          Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                          Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                          Rental_expiry_date = 
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                          Rental_Status =
                            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                    ? (
                                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                          ? "غير مؤجر (فارغ)"
                                          : "منتظر العقد"
                                      )
                                    : "مؤجر"),

                          Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                          Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                          Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                          Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                          Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                          Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,           
                          OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                          ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                          Notes = I.IsNotesNull() ? "" : I.Notes,
                          Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                      };

                    var joinedList = query
                    .Where(r => r.Visible == true) // 👈 only show visible rows
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
                    DataTable original = ToDataTable(joinedList);

                    // --- Convert string date columns to DateTime ---
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
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.ValueType == typeof(DateTime))
                        {
                            col.DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }

                    // --- Bind data to DataGridView ---
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
                    
                    originalData = converted;
                }

                if (SessionData.Investment_Type == "محلات شل اوت خارج" && SessionData.Investment_para1 == "شل اوت المدن")
                {

                    skyButton6.Visible = false;
                    TrueFunction();

                    // --- Part 1: Fetch and join data ---
                    var query =
                      from I in investmentsTableAdapter.GetData()
                      join g in governorateTableAdapter.GetData()
                          on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                      from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                      join l in landsTableAdapter.GetData()
                          on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                      from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                      join p in projectsTableAdapter.GetData()
                          on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                      from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                      join d in documentsTableAdapter.GetData()
                      on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                      from d in dj
                          .Where(x => x.approvals_fk == "APP009")
                          .DefaultIfEmpty()
                      where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                              && (!I.IsDependent_neighborhoodNull() && !I.Dependent_neighborhood.Contains("*"))
                      select new
                      {
                          Project_Id = p == null ? "" : p.project_id,
                          investments_id = I.investments_id,
                          Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                          PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                          Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                          lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                          Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                          City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                          Address = l.IsAddressNull() ? "" : l.Address,
                          GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                          investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                          Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                          investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,   
                          Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                          Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                          Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                          Location = I.IsLocationNull() ? "" : I.Location,
                          Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                          Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                          Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                          Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                          Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                          Rental_Status =
                        !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                            ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                            : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                ? (
                                    string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                    string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                      ? "غير مؤجر (فارغ)"
                                      : "منتظر العقد"
                                  )
                                : "مؤجر"),

                          Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                          Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                          Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                          Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                          Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                          Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,                                               
                          OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                          ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                          Notes = I.IsNotesNull() ? "" : I.Notes,
                          Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                      };

                    var joinedList = query
                    .Where(r => r.Visible == true) // 👈 only show visible rows
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
                    DataTable original = ToDataTable(joinedList);

                    // --- Convert string date columns to DateTime ---
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
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.ValueType == typeof(DateTime))
                        {
                            col.DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }

                    // --- Bind data to DataGridView ---
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
                    //advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    originalData = converted;
                }

                if (SessionData.Investment_Type == "محلات شل اوت خارج" && SessionData.Investment_para1 == "الكل")
                {
                    TrueFunction();
                    skyButton6.Visible = false;
                    // --- Part 1: Fetch and join data ---
                    var query =
                      from I in investmentsTableAdapter.GetData()
                          // --- Join lands first ---
                      join l in landsTableAdapter.GetData()
                          on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                      from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                          // --- Then join governorate based on the land’s governorate_fk, not the investment’s ---
                      join g in governorateTableAdapter.GetData()
                          on (l == null || l.Isgovernorate_fkNull() ? "-1" : l.governorate_fk) equals g.governorate_id into gj
                      from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate (from land)
                      join p in projectsTableAdapter.GetData()
                          on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                      from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                      join d in documentsTableAdapter.GetData()
                      on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                      from d in dj
                          .Where(x => x.approvals_fk == "APP009")
                          .DefaultIfEmpty()
                      where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                      select new
                      {
                          Project_Id = p == null ? "" : p.project_id,
                          investments_id = I.investments_id,
                          Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                          PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                          Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                          investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                          lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                          Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                          City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                          Address = l.IsAddressNull() ? "" : l.Address,
                          GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                          investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                          Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,                       
                          Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                          Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                          Location = I.IsLocationNull() ? "" : I.Location,
                          Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                          Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                          Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                          Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                          Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                          Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                          Rental_Status =
                        !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                            ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                            : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                ? (
                                    string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                    string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                      ? "غير مؤجر (فارغ)"
                                      : "منتظر العقد"
                                  )
                                : "مؤجر"),

                          Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                          Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                          Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                          Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                          Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                          Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,  
                          OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                          ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                          Notes = I.IsNotesNull() ? "" : I.Notes,
                          Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                      };

                    var joinedList = query
                    .Where(r => r.Visible == true)
                    .GroupBy(r => r.investments_id)
                    .Select(g => g.First()) // pick one of each investment_id
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
                    DataTable original = ToDataTable(joinedList);
                    // --- Convert string date columns to DateTime ---
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
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.ValueType == typeof(DateTime))
                        {
                            col.DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }
                    // --- Bind data to DataGridView ---
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
                    //advancedDataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    originalData = converted;
                }

                if (SessionData.Investment_Type == "مواقف")
                {
                    skyButton6.Visible = true;

                    TrueFunction();

                    // --- Part 1: Fetch and join data ---
                    var query =
                      from I in investmentsTableAdapter.GetData()
                      join g in governorateTableAdapter.GetData()
                          on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                      from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                      join l in landsTableAdapter.GetData()
                          on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                      from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                      join p in projectsTableAdapter.GetData()
                          on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                      from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                      join d in documentsTableAdapter.GetData()
                      on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                      from d in dj
                          .Where(x => x.approvals_fk == "APP009")
                          .DefaultIfEmpty()
                      where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                      select new
                      {

                          Project_Id = p == null ? "" : p.project_id,
                          investments_id = I.investments_id,
                          Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                          PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                          Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                          GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                          investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                          //Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                          Name_Projects = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetNameWithoutParentheses(I.Dependent_neighborhood),

                          investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                          //lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                          //Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                          //City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                          //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull()?"": p.Architectural_and_Structural_Board,
                          Dependent_neighborhood = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetTextInsideParentheses(I.Dependent_neighborhood),
                          Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                          Location = I.IsLocationNull() ? "" : I.Location,
                          Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                          //Description_Drawing_Place = I.IsDescription_Drawing_PlaceNull() ? "" : I.Description_Drawing_Place,
                          Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                          Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                          Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                          Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                          Rental_Status =
                            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                    ? (
                                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                          ? "غير مؤجر (فارغ)"
                                          : "منتظر العقد"
                                      )
                                    : "مؤجر"),
                          Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                          Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                          Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                          Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                          Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                          Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                          OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                          ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                          Notes = I.IsNotesNull() ? "" : I.Notes,
                          Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                      };

                    var joinedList = query
                     .Where(r => r.Visible == true)
                     .GroupBy(r => r.investments_id)
                     .Select(g => g.First()) // pick one of each investment_id
                     .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                     .ToList();
                    DataTable original = ToDataTable(joinedList);

                    // --- Convert string date columns to DateTime ---
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
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.ValueType == typeof(DateTime))
                        {
                            col.DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }

                    // --- Bind data to DataGridView ---
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

                    originalData = converted;
                }

                if (SessionData.Investment_Type == "مناطق تنموية")
                {
                    skyButton6.Visible = false;
                    checkedListBox2.Visible = true;
                    var documents = documentsTableAdapter.GetData().ToList();
                    var approvals = approvalsTableAdapter.GetData().ToList();
                    
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
                    TrueFunction();

                    // --- Part 1: Fetch and join data ---
                    var query =
                      from I in investmentsTableAdapter.GetData()
                      join g in governorateTableAdapter.GetData()
                          on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                      from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                      join l in landsTableAdapter.GetData()
                          on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                      from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                      join p in projectsTableAdapter.GetData()
                          on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                      from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                      join d in documentsTableAdapter.GetData()
                      on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                      from d in dj
                          .Where(x => x.approvals_fk == "APP009")
                          .DefaultIfEmpty()
                      where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                      select new
                      {

                          Project_Id = p == null ? "" : p.project_id,
                          investments_id = I.investments_id,
                          Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                          PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                          Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                          land_name = l == null || l.Island_nameNull() ? "" : l.land_name,
                          Model_8_Status = p.Model_8_Status,
                          Model8File = string.Join(" , ",
                            docsWithApprovals
                                .Where(d => d.ProjectId == p.project_id &&
                                            d.ApprovalName == "نموذج 8 أو 10")
                                .Select(d => d.Path)),


                          Civil_Defense_Approval_status = p.Civil_Defense_Approval_status,
                          CivilDefenseFile = string.Join(" , ",
                            docsWithApprovals
                                .Where(d => d.ProjectId == p.project_id &&
                                            d.ApprovalName == "موافقة الحماية المدنية")
                                .Select(d => d.Path)),
                          GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                          investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                          //Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                          Name_Projects = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetNameWithoutParentheses(I.Dependent_neighborhood),

                          investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                          //lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                          //Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                          //City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                          //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull()?"": p.Architectural_and_Structural_Board,
                          Dependent_neighborhood = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetTextInsideParentheses(I.Dependent_neighborhood),
                          Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                          Location = I.IsLocationNull() ? "" : I.Location,
                          Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                          //Description_Drawing_Place = I.IsDescription_Drawing_PlaceNull() ? "" : I.Description_Drawing_Place,
                          Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                          Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                          Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                          Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                          Rental_Status =
                            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                    ? (
                                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                          ? "غير مؤجر (فارغ)"
                                          : "منتظر العقد"
                                      )
                                    : "مؤجر"),
                          Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                          Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                          Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                          Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                          Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                          Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                          OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                          ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                          Notes = I.IsNotesNull() ? "" : I.Notes,
                          Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                      };

                    var joinedList = query
                     .Where(r => r.Visible == true)
                     .GroupBy(r => r.investments_id)
                     .Select(g => g.First()) // pick one of each investment_id
                     .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                     .ToList();
                    DataTable original = ToDataTable(joinedList);

                    // --- Convert string date columns to DateTime ---
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
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.ValueType == typeof(DateTime))
                        {
                            col.DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }

                    // --- Bind data to DataGridView ---
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

                    originalData = converted;
                    LoadRoadsToCheckedListBox();
                }
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            FalseFunction();
        }
        public void LoadInvestmentDataForRoads()
        {
            try
            {
                TrueFunction();

                // --- Part 1: Fetch and join data ---
                var query =
                  from I in investmentsTableAdapter.GetData()
                  join g in governorateTableAdapter.GetData()
                      on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                  from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                  join l in landsTableAdapter.GetData()
                      on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                  from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                  join p in projectsTableAdapter.GetData()
                      on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                  from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                  join d in documentsTableAdapter.GetData()
                  on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                  from d in dj
                      .Where(x => x.approvals_fk == "APP009")
                      .DefaultIfEmpty()
                     where
                           (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                  select new
                  {
                      Project_Id = p == null ? "" : p.project_id,
                      investments_id = I.investments_id,
                      Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                      PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                      Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                      Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                      GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                      investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                      Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                      investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                      //lAND_Dependent_neighborhood=l.IsDependent_neighborhoodNull()?"":l.Dependent_neighborhood,
                      //Dependent_road=l.IsDependent_roadNull()?"":l.Dependent_road,
                      //City_Name=l.IsCity_NameNull()?"":l.City_Name,

                      //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                      Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                      Location = I.IsLocationNull() ? "" : I.Location,
                      Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                      Description_Drawing_Place = d == null || d.IspathsNull()
                      ? ""
                      : System.IO.Path.GetFileName(d.paths),
                      Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                      Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                      Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                      Rental_Status =
                        (I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                            ? (
                                string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                  ? "غير مؤجر (فارغ)"
                                  : "منتظر العقد"
                              )
                            : "مؤجر",
                      Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                      Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                      Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                      Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                      Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                      Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                      OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                      ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                      Notes = I.IsNotesNull() ? "" : I.Notes,
                      Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                  };

                var joinedList = query
                .Where(r => r.Visible == true) // 👈 only show visible rows
                .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                .ToList();
                if (SessionData.Investment_para1 == "الكل")
                {
                    joinedList = joinedList
                        .Where(r => r.investments_id.StartsWith("&") || r.Dependent_neighborhood == "*")
                        .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                        .ToList();
                }
                else
                {
                    // For other filters, keep your default sort and selection
                    joinedList = joinedList
                        .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                        .ToList();
                }

                DataTable original = ToDataTable(joinedList);

                // --- Convert string date columns to DateTime ---
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
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                    }
                }

                // --- Bind data to DataGridView ---
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
                originalData = converted; // after you build your full DataTable
                //advancedDataGridView1.DataSource = new BindingSource { DataSource = originalData };
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            FalseFunction();
          
        }

        public void LoadInvestmentDataMarketsIn()
        {
            try
            {
                TrueFunction();

                // --- Part 1: Fetch and join data ---
                var query =
                   from I in investmentsTableAdapter.GetData()

                       // --- Join lands first ---
                   join l in landsTableAdapter.GetData()
                       on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                   from l in lj.DefaultIfEmpty() // LEFT JOIN for lands

                       // --- Then join governorate based on the land’s governorate_fk, not the investment’s ---
                   join g in governorateTableAdapter.GetData()
                       on (l == null || l.Isgovernorate_fkNull() ? "-1" : l.governorate_fk) equals g.governorate_id into gj
                   from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate (from land)

                       // --- Then join projects ---
                   join p in projectsTableAdapter.GetData()
                       on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                   from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                   join d in documentsTableAdapter.GetData()
                    on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                   from d in dj
                       .Where(x => x.approvals_fk == "APP009")
                       .DefaultIfEmpty()
                   where string.IsNullOrEmpty(SessionData.Investment_Type)
                         || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type)

                   select new
                   {
                       Project_Id = p == null ? "" : p.project_id,
                       investments_id = I.investments_id,
                       Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                       PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                       Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                       lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                       Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                       City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                       Address = l.IsAddressNull() ? "" : l.Address,
                       GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                       investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                       Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                   // Name_Projects = I == null || I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                      investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                       //Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                       Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                       Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                       Location = I.IsLocationNull() ? "" : I.Location,
                       Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                       Description_Drawing_Place = d == null || d.IspathsNull()
                      ? ""
                      : System.IO.Path.GetFileName(d.paths),
                       Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                       Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                       Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                       Rental_Status =
            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                    ? (
                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                          ? "غير مؤجر (فارغ)"
                          : "منتظر العقد"
                      )
                    : "مؤجر"),

                      Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                      Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                      Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                      Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                      Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                      Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                      OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                      ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                       Notes = I.IsNotesNull() ? "" : I.Notes,
                       Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                   };

                var joinedList = query
                .Where(r => r.Visible == true)
                .GroupBy(r => r.investments_id)
                .Select(g => g.First()) // or aggregate values if needed
                .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                .ToList();
                DataTable original = ToDataTable(joinedList);

                // --- Convert string date columns to DateTime ---
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
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                    }
                }

                // --- Bind data to DataGridView ---
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
                originalData = converted;
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            FalseFunction();
        }
        public void LoadInvestmentDataMALLS()
        {
            try
            {
                TrueFunction();

                // --- Part 1: Fetch and join data ---
                var query =
                   from I in investmentsTableAdapter.GetData()

                       // --- Join lands first ---
                   join l in landsTableAdapter.GetData()
                       on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                   from l in lj.DefaultIfEmpty() // LEFT JOIN for lands

                       // --- Then join governorate based on the land’s governorate_fk, not the investment’s ---
                   join g in governorateTableAdapter.GetData()
                       on (l == null || l.Isgovernorate_fkNull() ? "-1" : l.governorate_fk) equals g.governorate_id into gj
                   from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate (from land)

                       // --- Then join projects ---
                   join p in projectsTableAdapter.GetData()
                       on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                   from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                   join d in documentsTableAdapter.GetData()
                    on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                   from d in dj
                       .Where(x => x.approvals_fk == "APP009")
                       .DefaultIfEmpty()
                   where string.IsNullOrEmpty(SessionData.Investment_Type)
                         || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type)

                   select new
                   {
                       Project_Id = p == null ? "" : p.project_id,
                       investments_id = I.investments_id,
                       Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                       PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                       Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                       lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                       Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                       City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                       Address = l.IsAddressNull() ? "" : l.Address,
                       GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                       investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                       Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : 
                       p.Name_Projects.StartsWith("شل اوت")?
                       I.Dependent_neighborhood:p.Name_Projects,
                       // Name_Projects = I == null || I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                       investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                       Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                       //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                       Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                       Location = I.IsLocationNull() ? "" : I.Location,
                       Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                       Description_Drawing_Place = d == null || d.IspathsNull()
                      ? ""
                      : System.IO.Path.GetFileName(d.paths),
                       Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                       Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                       Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                       Rental_Status =
            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                    ? (
                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                          ? "غير مؤجر (فارغ)"
                          : "منتظر العقد"
                      )
                    : "مؤجر"),

                       Rental_value = I.IsRental_valueNull() ? 0m : I.Rental_value,
                       Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                       Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                       Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                       Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                       Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                       OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                       ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                       Notes = I.IsNotesNull() ? "" : I.Notes,
                       Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                   };

                var joinedList = query
                .Where(r => r.Visible == true)
                .GroupBy(r => r.investments_id)
                .Select(g => g.First()) // or aggregate values if needed
                .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                .ToList();
                DataTable original = ToDataTable(joinedList);

                // --- Convert string date columns to DateTime ---
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
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                    }
                }

                // --- Bind data to DataGridView ---
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
                originalData = converted;
               
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
            }
            ApplyUnitFilter();
            advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            FalseFunction();
        }
        private Dictionary<string, string> columnMap = new Dictionary<string, string>
        {
            
            { "مسلسل القطعة", "Land_Id" },
            { "نوع العقد", "Location" },
            {"المكان","investment_type"},
            { "رقم اللوحة", "PlateNumber" },
            { "مسلسل كل محل", "investments_id" },
            { "رقم قطعة الأرض", "Land_Number" },
            // { "اسم المشروع", "ProjectName" }, // commented out because you disabled it
            { "اسم المحافظة", "GovernorateName" },
            { "الحي التابع", "Dependent_neighborhood" },
            { "نوع النشاط", "Activity_Type" },
            { "صورة العقد", "Activity_Name" },
            { "رقم المكان", "Place_number" },
            { "رقم مذكرة العرض", "Offer_memorandum_number" },
            { "رقم العقد", "Contract_number" },
            { "القيمة الايجارية الحالية", "Rental_value" },
            { "تاريخ توقيع العقد", "Contract_start_date" },
            { "تاريخ انتهاء العقد", "Contract_expiry_date" },
            {"اسم الاستثمار","investment_name"},
            {"اسم المكان","Name_Projects"},
            {"موقف التأجير","Rental_Status"}
            
            // hidden columns like OfferfilePaths and ContractFilePaths are excluded
        };


        public void ArabicColumnGrid()
        {
            if (SessionData.Investment_Type != "محلات شل اوت داخل" && SessionData.Investment_Type != "مولات")
            {
                // Set Arabic headers manually
                advancedDataGridView1.Columns["Notes"].HeaderText = "ملاحظات";
                advancedDataGridView1.Columns["Contract_terms"].HeaderText = "شروط العقد";         
                advancedDataGridView1.Columns["Land_Id"].HeaderText = "مسلسل القطعة";
                advancedDataGridView1.Columns["Location"].HeaderText = "نوع العقد";
                advancedDataGridView1.Columns["PlateNumber"].HeaderText = "رقم اللوحة";
                advancedDataGridView1.Columns["investments_id"].HeaderText = "مسلسل كل محل";
                advancedDataGridView1.Columns["Land_Number"].HeaderText = "رقم قطعة الأرض";
                advancedDataGridView1.Columns["investment_name"].HeaderText = "اسم الاستثمار";
                advancedDataGridView1.Columns["Name_Projects"].HeaderText = "اسم المكان";
                advancedDataGridView1.Columns["Rental_Status"].HeaderText = "موقف التأجير";
               
                if (SessionData.Investment_Type != "كل الاستثمارات")
                {
                    advancedDataGridView1.Columns["Dependent_neighborhood"].HeaderText = "الحي التابع";
                }
                if (SessionData.Investment_Type == "مناطق تنموية") {
                    advancedDataGridView1.Columns["land_name"].HeaderText = "اسم قطعة الارض";
                    advancedDataGridView1.Columns["Model_8_Status"].HeaderText = "نموذج 8 او 10";
                    advancedDataGridView1.Columns["Model8File"].HeaderText = "ملف نموذج 8 او 10";
                    advancedDataGridView1.Columns["Civil_Defense_Approval_status"].HeaderText = "الحماية المدنية";
                    advancedDataGridView1.Columns["CivilDefenseFile"].HeaderText = "ملف الحماية المدنية";
                    advancedDataGridView1.Columns["Model8File"].Visible = false;
                    advancedDataGridView1.Columns["CivilDefenseFile"].Visible = false;
                }
                //advancedDataGridView1.Columns["Dependent_neighborhood"].HeaderText = "الحي التابع";    
                //advancedDataGridView1.Columns["ProjectName"].HeaderText = "اسم المشروع";
                advancedDataGridView1.Columns["GovernorateName"].HeaderText = "اسم المحافظة";
                if (SessionData.Investment_Type == "محلات شل اوت خارج" || SessionData.Investment_Type == "كل الاستثمارات")
                {
                    advancedDataGridView1.Columns["lAND_Dependent_neighborhood"].HeaderText = "الحي التابع";
                    advancedDataGridView1.Columns["Dependent_road"].HeaderText = "الطريق التابع";
                    advancedDataGridView1.Columns["City_Name"].HeaderText = "مدينة";
                    advancedDataGridView1.Columns["Address"].HeaderText = "العنوان";
                    if (SessionData.Investment_Type == "محلات شل اوت خارج")
                    {
                        advancedDataGridView1.Columns["Architectural_and_Structural_Board"].HeaderText = "شركة ادارة المحطة";
                    }
                }
                advancedDataGridView1.Columns["Activity_Type"].HeaderText = "نوع النشاط";
                advancedDataGridView1.Columns["Activity_Name"].HeaderText = "صورة العقد";
                advancedDataGridView1.Columns["Description_Drawing_Place"].HeaderText = "كروكي المكان";
                advancedDataGridView1.Columns["Place_number"].HeaderText = "رقم المكان";
                advancedDataGridView1.Columns["Offer_memorandum_number"].HeaderText = "رقم مذكرة العرض";
                advancedDataGridView1.Columns["Contract_number"].HeaderText = "رقم العقد";
                advancedDataGridView1.Columns["Rental_value"].HeaderText = "القيمة الايجارية الحالية";
                advancedDataGridView1.Columns["Contract_start_date"].HeaderText = "تاريخ توقيع العقد";
                advancedDataGridView1.Columns["Contract_expiry_date"].HeaderText = "تاريخ انتهاء العقد";
                advancedDataGridView1.Columns["Rental_expiry_date"].HeaderText = "نهاية المده";
                advancedDataGridView1.Columns["investment_type"].HeaderText = "المكان";
                advancedDataGridView1.Columns["OfferfilePaths"].HeaderText = "مستند رقم مذكرة العرض";
                advancedDataGridView1.Columns["ContractFilePaths"].HeaderText = "مستند رقم العقد";
                advancedDataGridView1.Columns["OfferfilePaths"].Visible = false;
                advancedDataGridView1.Columns["ContractFilePaths"].Visible = false;
                advancedDataGridView1.Columns["Project_Id"].Visible = false;
                advancedDataGridView1.Columns["investment_type"].Visible = false;

                advancedDataGridView1.Columns["Visible"].Visible = false;
                advancedDataGridView1.Columns["Shops_Count"].HeaderText = "عدد المحلات";
                

            }
            else
            {
                // Set Arabic headers manually
                if (SessionData.Investment_Type == "محلات شل اوت داخل")
                {
                    advancedDataGridView1.Columns["Architectural_and_Structural_Board"].HeaderText = "شركة ادارة المحطة";
                }
                if (SessionData.Investment_Type == "مولات")
                {
                    advancedDataGridView1.Columns["Dependent_neighborhood"].HeaderText = "اسم المول";
                }
                advancedDataGridView1.Columns["Notes"].HeaderText = "ملاحظات";
                advancedDataGridView1.Columns["Contract_terms"].HeaderText = "شروط العقد";
                advancedDataGridView1.Columns["Land_Id"].HeaderText = "مسلسل القطعة";
                advancedDataGridView1.Columns["Location"].HeaderText = "نوع العقد";
                advancedDataGridView1.Columns["PlateNumber"].HeaderText = "رقم اللوحة";
                advancedDataGridView1.Columns["investments_id"].HeaderText = "مسلسل كل محل";
                advancedDataGridView1.Columns["Land_Number"].HeaderText = "رقم قطعة الأرض";
                advancedDataGridView1.Columns["investment_name"].HeaderText = "اسم الاستثمار";
                advancedDataGridView1.Columns["Name_Projects"].HeaderText = "اسم المكان";
                advancedDataGridView1.Columns["Rental_Status"].HeaderText = "موقف التأجير";
                //advancedDataGridView1.Columns["ProjectName"].HeaderText = "اسم المشروع";
                advancedDataGridView1.Columns["Address"].HeaderText = "العنوان";
                
                advancedDataGridView1.Columns["Description_Drawing_Place"].HeaderText = "كروكي المكان";
                advancedDataGridView1.Columns["GovernorateName"].HeaderText = "اسم المحافظة";
                advancedDataGridView1.Columns["lAND_Dependent_neighborhood"].HeaderText = "الحي التابع";
                advancedDataGridView1.Columns["Dependent_road"].HeaderText = "الطريق التابع";
                advancedDataGridView1.Columns["City_Name"].HeaderText = "مدينة";
                advancedDataGridView1.Columns["Activity_Type"].HeaderText = "نوع النشاط";
                advancedDataGridView1.Columns["Activity_Name"].HeaderText = "صورة العقد";
                advancedDataGridView1.Columns["Place_number"].HeaderText = "رقم المكان";
                advancedDataGridView1.Columns["Offer_memorandum_number"].HeaderText = "رقم مذكرة العرض";
                advancedDataGridView1.Columns["Contract_number"].HeaderText = "رقم العقد";
                advancedDataGridView1.Columns["Rental_value"].HeaderText = "القيمة الايجارية الحالية";
                advancedDataGridView1.Columns["Contract_start_date"].HeaderText = "تاريخ توقيع العقد";
                advancedDataGridView1.Columns["Contract_expiry_date"].HeaderText = "تاريخ انتهاء العقد";
                advancedDataGridView1.Columns["Rental_expiry_date"].HeaderText = "نهاية المده";
                advancedDataGridView1.Columns["investment_type"].HeaderText = "المكان";
                advancedDataGridView1.Columns["OfferfilePaths"].HeaderText = "مستند رقم مذكرة العرض";
                advancedDataGridView1.Columns["ContractFilePaths"].HeaderText = "مستند رقم العقد";
                advancedDataGridView1.Columns["OfferfilePaths"].Visible = false;
                advancedDataGridView1.Columns["ContractFilePaths"].Visible = false;
                advancedDataGridView1.Columns["Project_Id"].Visible = false;
                advancedDataGridView1.Columns["investment_type"].Visible = false;
                advancedDataGridView1.Columns["Shops_Count"].HeaderText = "عدد المحلات";
                advancedDataGridView1.Columns["Visible"].Visible = false;
            }
          
          

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
            var governorateNames = GoverData.AsEnumerable()
             .Select(r =>
             {
                 string govId = r["governorate_id"].ToString();
                 return governorateLookup.ContainsKey(govId) ? governorateLookup[govId] : "غير معروف";
             })
             .Distinct()
             .ToList();
            Activity = this.activity_LookupTableAdapter.GetData();
            activityLookup = Activity.AsEnumerable().
                ToDictionary(r => r["Id"].ToString(), r => r["Activity_Type"].ToString());
            Land_Name_COB.DataSource = landNames;
            plate_number_COB.DataSource = plateNumbers;
            Land_ID_COB.DataSource = landIds;
            Governorate_COB.DataSource = governorateNames;
            Activity_Type_TB1.DataSource = activityLookup.Values.ToList();
            AdjustComboBox(Land_Name_COB);
            AdjustComboBox(plate_number_COB);
            AdjustComboBox(Land_ID_COB);
            AdjustComboBox(Governorate_COB);
            AdjustComboBox(Activity_Type_TB);
            StyleComboBox(Activity_Type_TB1);
        }
        private void StyleNormalComboBox(ComboBox combo, Label label)
        {
            // Enable typing
            combo.DropDownStyle = ComboBoxStyle.DropDown;
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;

            // Flat look
            combo.FlatStyle = FlatStyle.Flat;
            combo.BackColor = Color.White;
            combo.ForeColor = Color.FromArgb(68, 88, 112);
            combo.Font = new System.Drawing.Font("Segoe UI", 10F);

            // Remove ugly border
            combo.Region = new Region(new System.Drawing.Rectangle(0, 0, combo.Width, combo.Height - 2));

            // Label (like Guna2 floating label)
            label.Font = new System.Drawing.Font("Segoe UI", 9F, FontStyle.Regular);
            label.ForeColor = Color.Gray;
            label.TextAlign = ContentAlignment.MiddleRight;
            label.BringToFront();

            // Paint underline under the combo
            combo.Paint += (s, e) =>
            {
                using (Pen p = new Pen(Color.FromArgb(94, 148, 255), 2))
                {
                    e.Graphics.DrawLine(p, 0, combo.Height - 1, combo.Width, combo.Height - 1);
                }
            };
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
        public void LoadGovernorates()
        {
            GoverData = governorateTableAdapter.GetData();

            // governorateLookup dictionary
            governorateLookup = GoverData.AsEnumerable()
                .ToDictionary(r => r["governorate_id"].ToString(), r => r["governorate"].ToString());

            // كل المحافظات بالاسم
            var governorateNames = GoverData.AsEnumerable()
                .Select(r => r["governorate"].ToString())
                .Distinct()
                .ToList();

            Governorate_COB.DataSource = governorateNames;
        }
        private void PositionHeaderButton()
        {
            int headerRight = guna2TabControl1.Left + guna2TabControl1.Width - guna2Button3.Width - 5;
            int headerTop = guna2TabControl1.Top;
            guna2Button3.Location = new Point(headerRight - 5, headerTop);
            Add_Radio.Location = new Point(headerRight - 680, headerTop);
            Update_Radio.Location = new Point(headerRight - 730, headerTop);
        }
        private void StyleComboBox(ComboBox combo)
        {
            combo.DropDownStyle = ComboBoxStyle.DropDown; // typing enabled
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;
            combo.FlatStyle = FlatStyle.Flat;
            combo.Font = new System.Drawing.Font("Segoe UI", 10F);
            combo.ForeColor = Color.FromArgb(68, 88, 112);
            combo.BackColor = Color.White;
            combo.RightToLeft = RightToLeft.Yes; // Arabic support
        }
        private void StyleGunaComboBox(Guna.UI2.WinForms.Guna2ComboBox combo)
        {
            combo.DropDownStyle = ComboBoxStyle.DropDown; // typing enabled
            combo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            combo.AutoCompleteSource = AutoCompleteSource.ListItems;

            combo.DrawMode = DrawMode.OwnerDrawFixed;
            combo.BorderThickness = 0;
            combo.Style = Guna.UI2.WinForms.Enums.TextBoxStyle.Material;

            combo.FocusedState.BorderColor = Color.FromArgb(94, 148, 255);
            combo.HoverState.BorderColor = Color.FromArgb(94, 148, 255);
            combo.BackColor = Color.White;
            combo.ForeColor = Color.FromArgb(68, 88, 112);
        }
        string IsAllData = "";
        string NoPara1 = "";
        private void LoadInvestmentDataForSelectedRoads()
        {
           
            if (SessionData.Investment_Type == "كل الاستثمارات") {
                IsAllData = SessionData.Investment_Type;
                NoPara1 = "كل الاستثمارات";
                SessionData.Investment_para1 = "محلات علي الطرق فقط";
                SessionData.Investment_Type = "استثمارات على الطرق";
            }
            // --- Get selected roads from CheckedListBox ---
            var selectedRoads = checkedListBox2.CheckedItems
                .Cast<string>()
                .Select(r => r.Trim())
                .ToHashSet();

            // --- Fetch all required data ---
            var allInvestments = investmentsTableAdapter.GetData();
            var allGovernorates = governorateTableAdapter.GetData();
            var allLands = landsTableAdapter.GetData();
            var allProjects = projectsTableAdapter.GetData();

            // --- Build main query ---
            var query =
                from I in allInvestments
                join g in allGovernorates
                    on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                join l in allLands
                    on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                join p in allProjects
                    on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                join d in documentsTableAdapter.GetData()
                on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                from d in dj
                    .Where(x => x.approvals_fk == "APP009")
                    .DefaultIfEmpty()
                where
                   (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                    &&
                    // --- Road filter: if none selected, show all roads ---
                    (
                        selectedRoads.Count == 0
                        || (p != null && !p.IsName_ProjectsNull() && selectedRoads.Contains(p.Name_Projects.Trim()))
                    )
                select new
                {
                    Project_Id = p == null ? "" : p.project_id,
                    investments_id = I.investments_id,
                    Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*") || l.land_id.StartsWith("$"))
                        ? " "
                        : (l == null ? "" : l.land_id),
                    investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                    Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                    investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                    Location = I.IsLocationNull() ? "" : I.Location,
                    PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                    Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                    GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                    Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                    Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                    Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                   Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                    Description_Drawing_Place = d == null || d.IspathsNull()
                      ? ""
                      : System.IO.Path.GetFileName(d.paths),
                    Rental_Status =
                        (I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                            ? (
                                string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                    ? "غير مؤجر (فارغ)"
                                    : "منتظر العقد"
                              )
                            : "مؤجر",
                    Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                    Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                    Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                    Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                    Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                    Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                    Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                    Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                    Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                    OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                    ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File
                };

            // --- Filter visible and sort ---
            var joinedList = query
                .Where(r => r.Visible == true)
                .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                .ToList();

            // --- Handle “الكل” case separately ---
            if (SessionData.Investment_para1 == "الكل")
            {
                joinedList = joinedList
                    .Where(r => r.investments_id.StartsWith("&") || r.Dependent_neighborhood == "*")
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
            }
            else
            {
                joinedList = joinedList
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
            }

            // --- Convert to DataTable and bind ---
            DataTable original = ToDataTable(joinedList);
            advancedDataGridView1.DataSource = original;
            if (IsAllData != "")
            {
                SessionData.Investment_Type = IsAllData;
                SessionData.Investment_para1 = NoPara1;
                NoPara1="";
                IsAllData = "";
            }
            UpdateRowCount();
        }
        private void LoadInvestmentDataForSelectedRoads2()
        {

            if (SessionData.Investment_Type == "كل الاستثمارات")
            {
                IsAllData = SessionData.Investment_Type;
                NoPara1 = "كل الاستثمارات";
                SessionData.Investment_para1 = "محلات علي الطرق فقط";
                SessionData.Investment_Type = "استثمارات على الطرق";
            }
            // --- Get selected roads from CheckedListBox ---
            var selectedTypes = checkedListBox2.CheckedItems
             .Cast<string>()
             .Select(x => x.Trim())
             .ToList();
            bool isAllSelected = selectedTypes.Contains("الكل");

            // --- Fetch all required data ---
            var allInvestments = investmentsTableAdapter.GetData();
            var allGovernorates = governorateTableAdapter.GetData();
            var allLands = landsTableAdapter.GetData();
            var allProjects = projectsTableAdapter.GetData();
            var documents = documentsTableAdapter.GetData().ToList();
            var approvals = approvalsTableAdapter.GetData().ToList();
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
            // --- Build main query ---
            var query =
                from I in allInvestments
                join g in allGovernorates
                    on I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk equals g.governorate_id into gj
                from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                join l in allLands
                    on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                from l in lj.DefaultIfEmpty() // LEFT JOIN for lands
                join p in allProjects
                    on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                join d in documentsTableAdapter.GetData()
                on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                from d in dj
                    .Where(x => x.approvals_fk == "APP009")
                    .DefaultIfEmpty()
                where
                   (string.IsNullOrEmpty(SessionData.Investment_Type)
                            || (!I.Isinvestment_typeNull() && I.investment_type == SessionData.Investment_Type))
                    &&
                    // --- Road filter: if none selected, show all roads ---
                    (
                        isAllSelected
                        || selectedTypes.Count == 0
                        || (l != null && !l.Island_nameNull() &&
                            selectedTypes.Any(type =>
                        l.land_name.IndexOf(type, StringComparison.OrdinalIgnoreCase) >= 0)
                    ))
                select new
                {
                    Project_Id = p == null ? "" : p.project_id,
                    investments_id = I.investments_id,
                    Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                    PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                    Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                    land_name = l == null || l.Island_nameNull() ? "" : l.land_name,
                    Model_8_Status = p.Model_8_Status,
                    Model8File = string.Join(" , ",
                            docsWithApprovals
                                .Where(d => d.ProjectId == p.project_id &&
                                            d.ApprovalName == "نموذج 8 أو 10")
                                .Select(d => d.Path)),


                    Civil_Defense_Approval_status = p.Civil_Defense_Approval_status,
                    CivilDefenseFile = string.Join(" , ",
                            docsWithApprovals
                                .Where(d => d.ProjectId == p.project_id &&
                                            d.ApprovalName == "موافقة الحماية المدنية")
                                .Select(d => d.Path)),
                    GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                    investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                    //Name_Projects = p == null || p.IsName_ProjectsNull() ? "" : p.Name_Projects,
                    Name_Projects = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetNameWithoutParentheses(I.Dependent_neighborhood),

                    investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                    //lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                    //Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                    //City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                    //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull()?"": p.Architectural_and_Structural_Board,
                    Dependent_neighborhood = I.IsDependent_neighborhoodNull()
                          ? ""
                          : GetTextInsideParentheses(I.Dependent_neighborhood),
                    Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                    Location = I.IsLocationNull() ? "" : I.Location,
                    Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                    //Description_Drawing_Place = I.IsDescription_Drawing_PlaceNull() ? "" : I.Description_Drawing_Place,
                    Description_Drawing_Place = d == null || d.IspathsNull()
                          ? ""
                          : System.IO.Path.GetFileName(d.paths),
                    Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                    Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                    Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                    Rental_Status =
                            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                    ? (
                                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                          ? "غير مؤجر (فارغ)"
                                          : "منتظر العقد"
                                      )
                                    : "مؤجر"),
                    Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                    Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                    Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                    Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                    Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                    Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                    OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                    ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                    Notes = I.IsNotesNull() ? "" : I.Notes,
                    Contract_terms = I.IsContract_termsNull() ? "" : I.Contract_terms
                };

            // --- Filter visible and sort ---
            var joinedList = query
                .Where(r => r.Visible == true)
                .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                .ToList();

            // --- Handle “الكل” case separately ---
            if (SessionData.Investment_para1 == "الكل")
            {
                joinedList = joinedList
                    .Where(r => r.investments_id.StartsWith("@") || r.Dependent_neighborhood == "*")
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
            }
            else
            {
                joinedList = joinedList
                    .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                    .ToList();
            }

            // --- Convert to DataTable and bind ---
            DataTable original = ToDataTable(joinedList);
            advancedDataGridView1.DataSource = original;
            if (IsAllData != "")
            {
                SessionData.Investment_Type = IsAllData;
                SessionData.Investment_para1 = NoPara1;
                NoPara1 = "";
                IsAllData = "";
            }
            UpdateRowCount();
        }
        private void HighlightExpiryCells()
        {
            DateTime today = DateTime.Today;

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.Cells["Contract_expiry_date"].Value != null &&
                    DateTime.TryParse(row.Cells["Contract_expiry_date"].Value.ToString(), out DateTime expiryDate))
                {
                    var cell = row.Cells["Contract_expiry_date"];
                    cell.Style.BackColor = Color.White;
                    cell.Style.ForeColor = Color.Black;

                    if (expiryDate < today)
                    {
                        cell.Style.BackColor = Color.Red;
                        cell.Style.ForeColor = Color.White;
                    }
                    else if (expiryDate <= today.AddMonths(1))
                    {
                        cell.Style.BackColor = Color.Yellow;
                        cell.Style.ForeColor = Color.Black;
                    }
                }
            }
        }
        public void LoadInvestmentDataAll()
        {

                // --- Part 1: Fetch and join data ---
                var query =
                  from I in investmentsTableAdapter.GetData()

                      // --- Join lands first ---
                  join l in landsTableAdapter.GetData()
                      on I.Island_fkNull() ? "-1" : I.land_fk equals l.land_id into lj
                  from l in lj.DefaultIfEmpty() // LEFT JOIN for lands

                      // --- Then join governorate based on the land’s governorate_fk, not the investment’s ---
                  join g in governorateTableAdapter.GetData()
                  on (
                      I.investments_id.StartsWith("&")
                          ? (I.Isgovernorate_fkNull() ? "-1" : I.governorate_fk)
                          : (l == null || l.Isgovernorate_fkNull() ? "-1" : l.governorate_fk)
                     ) equals g.governorate_id into gj
                              from g in gj.DefaultIfEmpty() // LEFT JOIN for governorate
                      // --- Then join projects ---
                  join p in projectsTableAdapter.GetData()
                      on (l == null ? "-1" : l.land_id) equals p.land_fk into pj
                  from p in pj.DefaultIfEmpty() // LEFT JOIN for projects
                  join d in documentsTableAdapter.GetData()
                  on p == null ? "-1" : p.project_id equals d.projects_fk into dj
                  from d in dj
                  .Where(x => x.approvals_fk == "APP009")
                  .DefaultIfEmpty()
                  select new
                  {

                      Project_Id = p == null ? "" : p.project_id,
                      investments_id = I.investments_id,
                      Land_Id = l != null && (l.land_id.StartsWith("%") || l.land_id.StartsWith("*")) ? " " : (l == null ? "" : l.land_id),
                      PlateNumber = l == null || l.Isplate_numberNull() ? "" : l.plate_number,
                      Land_Number = l == null || l.Island_numberNull() ? "" : l.land_number,
                      lAND_Dependent_neighborhood = l.IsDependent_neighborhoodNull() ? "" : l.Dependent_neighborhood,
                      Dependent_road = l.IsDependent_roadNull() ? "" : l.Dependent_road,
                      City_Name = l.IsCity_NameNull() ? "" : l.City_Name,
                      Address = l.IsAddressNull() ? "" : l.Address,
                      GovernorateName = g == null || g.IsgovernorateNull() ? "" : g.governorate,
                      investment_type = I.Isinvestment_typeNull() ? "" : I.investment_type,
                      Name_Projects =
                        (p == null || p.IsName_ProjectsNull() || string.IsNullOrWhiteSpace(p.Name_Projects))
                        ? (I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood)
                        : p.Name_Projects,
                      //Name_Projects = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                      investment_name = I.Isinvestment_nameNull() ? "" : I.investment_name,
                      //Dependent_neighborhood = I.IsDependent_neighborhoodNull() ? "" : I.Dependent_neighborhood,
                      //Architectural_and_Structural_Board = p.IsArchitectural_and_Structural_BoardNull() ? "" : p.Architectural_and_Structural_Board,
                      Activity_Type = I.IsActivity_TypeNull() ? "" : I.Activity_Type,
                      Location = I.IsLocationNull() ? "" : I.Location,
                      Activity_Name = I.IsActivity_NameNull() ? "" : I.Activity_Name,
                      Description_Drawing_Place = d == null || d.IspathsNull()
                      ? ""
                      : System.IO.Path.GetFileName(d.paths),
                      Contract_start_date = I.IsContract_start_dateNull() ? (DateTime?)null : I.Contract_start_date,
                      Contract_expiry_date = I.IsContract_expiry_dateNull() ? (DateTime?)null : I.Contract_expiry_date,
                      Rental_expiry_date =
                            (I.IsContract_start_dateNull() || I.IsContract_expiry_dateNull())
                                ? ""
                                : GetDateDifference(I.Contract_start_date, I.Contract_expiry_date),
                      Rental_value = I.IsRental_valueNull() ? 0 : I.Rental_value,
                      Rental_Status =
                            !string.IsNullOrEmpty(I.IsRental_StatusNull() ? "" : I.Rental_Status)
                                ? I.Rental_Status // keep existing status (e.g., "غير مؤجر تم الفسخ")
                                : ((I.IsActivity_NameNull() ? "" : I.Activity_Name) == "لا يوجد"
                                    ? (
                                        string.IsNullOrEmpty(I.Isinvestment_nameNull() ? "" : I.investment_name) ||
                                        string.IsNullOrEmpty(I.IsActivity_TypeNull() ? "" : I.Activity_Type)
                                          ? "غير مؤجر (فارغ)"
                                          : "منتظر العقد"
                                      )
                                    : "مؤجر"),
                      Shops_Count = I.IsShops_CountNull() ? 0 : I.Shops_Count,
                      Visible = I.IsVisable_ValueNull() ? false : I.Visable_Value,
                      Place_number = I.IsPlace_numberNull() ? "" : I.Place_number,
                      Offer_memorandum_number = I.IsOffer_memorandum_numberNull() ? "" : I.Offer_memorandum_number,
                      Contract_number = I.IsContract_numberNull() ? "" : I.Contract_number,
                      OfferfilePaths = I.IsOffer_memorandum_number_FileNull() ? "" : I.Offer_memorandum_number_File,
                      ContractFilePaths = I.IsContract_number_FileNull() ? "" : I.Contract_number_File,
                      Notes = I.IsNotesNull()?"":I.Notes,
                      Contract_terms = I.IsContract_termsNull()?"":I.Contract_terms
                  };

                var joinedList = query
                 .Where(r => r.Visible == true)
                 .GroupBy(r => r.investments_id)
                 .Select(g => g.First()) // pick one of each investment_id
                 .OrderBy(r => int.TryParse(r.investments_id, out var n) ? n : int.MaxValue)
                 .ToList();
                DataTable original = ToDataTable(joinedList);

                // --- Convert string date columns to DateTime ---
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
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.ValueType == typeof(DateTime))
                    {
                        col.DefaultCellStyle.Format = "dd/MM/yyyy";
                    }
                }

                // --- Bind data to DataGridView ---
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
                originalData = converted;
            
        }
        private void ApplyUnitFilter()
        {
            if (originalData == null) return;

            bool oneUnit = OneUnitCHB.Checked;
            bool multiUnit = MultyUnitsCHB.Checked;

            DataTable dt = originalData.Copy();

            // Count Name_Projects repetitions
            var nameGroups = dt.AsEnumerable()
                .Where(r => !string.IsNullOrWhiteSpace(r.Field<string>("Name_Projects")))
                .GroupBy(r => r.Field<string>("Name_Projects"))
                .ToDictionary(g => g.Key, g => g.Count());

            // Filtering logic
            IEnumerable<DataRow> filteredRows = dt.AsEnumerable();

            if (!oneUnit && multiUnit)
            {
                // Only single units
                filteredRows = filteredRows
                    .Where(r => nameGroups[r.Field<string>("Name_Projects")] == 1);
            }
            else if (oneUnit && !multiUnit)
            {
                // Only multi units
                filteredRows = filteredRows
                    .Where(r => nameGroups[r.Field<string>("Name_Projects")] > 1);
            }
            // else both checked → show all

            DataTable result =
                filteredRows.Any() ? filteredRows.CopyToDataTable() : dt.Clone();

            advancedDataGridView1.DataSource = result;
           
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);

            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
        }

        private void InvestmentsControl_Load(object sender, EventArgs e)
        {
            
            guna2TextBox2.Text = $"{SessionData.Investment_Type}";
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            Add_Radio.Parent = guna2TabControl1.Parent;
            Add_Radio.BringToFront();
            Update_Radio.Parent = guna2TabControl1.Parent;
            Update_Radio.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            if (SessionData.Investment_Type == "كل الاستثمارات")
            {
                checkedListBox2.Visible = false;
                skyButton5.Visible = true;
                skyButton4.Visible = true;
                skyButton6.Visible = true;
                OneUnitCHB.Visible = false;
                MultyUnitsCHB.Visible = false;
                OneUnitCHB.Checked = false;
                MultyUnitsCHB.Checked = false;
                guna2Panel5.Visible = false;
                LoadInvestmentDataAll();

                // checkedListBox2.Visible = true;
                //LoadRoadsToCheckedListBox();
            }

            if (SessionData.Investment_Type == "استثمارات على الطرق")
            {
                checkedListBox2.Visible = true;
                skyButton5.Visible = false;
                skyButton4.Visible = false;
                skyButton6.Visible = false;

                OneUnitCHB.Visible = false;
                MultyUnitsCHB.Visible = false;
                OneUnitCHB.Checked = false;
                MultyUnitsCHB.Checked = false;
                guna2Panel5.Visible = false;
                LoadInvestmentDataForRoads();
    

                LoadRoadsToCheckedListBox();
            }


            if (SessionData.Investment_Type == "محلات شل اوت داخل")
            {
                checkedListBox2.Visible = false;
                dreamButton4.Visible = true;
                skyButton5.Visible = true;
                skyButton4.Visible = false;
                skyButton6.Visible = false;

                OneUnitCHB.Visible = false;
                MultyUnitsCHB.Visible = false;
                OneUnitCHB.Checked = false;
                MultyUnitsCHB.Checked = false;
                guna2Panel5.Visible = false;
                LoadInvestmentDataMarketsIn();
             
            }

            if (SessionData.Investment_Type == "مولات")
            {
                checkedListBox2.Visible = false;
                dreamButton4.Visible = true;
                skyButton5.Visible = false;
                skyButton4.Visible = false;
                skyButton6.Visible = false;
                OneUnitCHB.Visible = true;
                MultyUnitsCHB.Visible = true;
                OneUnitCHB.Checked = true;
                MultyUnitsCHB.Checked = true;
                guna2Panel5.Visible = true;
                LoadInvestmentDataMALLS();
             
            }

            if (SessionData.Investment_Type != "محلات شل اوت داخل" && SessionData.Investment_Type != "مولات"
                && SessionData.Investment_Type != "استثمارات على الطرق" 
                && SessionData.Investment_Type != "كل الاستثمارات")
            {
                skyButton5.Visible = false;
                checkedListBox2.Visible = false;
                skyButton4.Visible = false;
                OneUnitCHB.Visible = false;
                MultyUnitsCHB.Visible = false;
                OneUnitCHB.Checked = false;
                MultyUnitsCHB.Checked = false;
                guna2Panel5.Visible = false;
                LoadInvestmentData();
         

            }

            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            
            ApplyGuna2StyleToGrid(advancedDataGridView2);
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
            guna2CircleButton3.Tag = "Function:Add";
            guna2Button3.Tag = "Function:Delete";
            guna2Button4.Tag = "Page:Settings";
            guna2TextBox1.Tag = "Page:Settings";
            dungeonLabel18.Tag = "Page:Settings";
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
            this.advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            advancedDataGridView1.Columns["Project_Id"].Visible = false;
            HighlightExpiryCells();
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
            checkedListBox1.Items.Add("اختيار الكل", true);
            // Make sure the DataGridView has a DataSource
            if (advancedDataGridView1.DataSource == null) return;

            // Loop through the columns and add their HeaderText or Name
            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                // ❌ Skip unwanted columns
                if (column.Name == "OfferfilePaths" 
                    || column.Name == "ContractFilePaths"  
                    || column.Name== "Project_Id" 
                    || column.Name == "Visible"
                    || (column.Name == "investment_type" && SessionData.Investment_Type != "كل الاستثمارات")
                    || column.Name == "Dependent_neighborhood"
                    || (column.Name == "Model8File")
                    || (column.Name == "CivilDefenseFile")
                    )
                    continue;

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
                
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fullPath = Path.Combine(desktopPath, $"{fileName}.xlsx");

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Add();
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            sheet.Name = fileName;

            // RTL
            sheet.DisplayRightToLeft = true;

            var visibleCols = advancedDataGridView1.Columns
                .Cast<DataGridViewColumn>()
                .Where(c => c.Visible && c.Name.ToLower() != "select")
                .ToList();

            int colCount = visibleCols.Count;
            int rowCount = advancedDataGridView1.Rows.Cast<DataGridViewRow>()
                .Count(r => !r.IsNewRow);

            // =======================
            // Title
            // =======================
            var titleRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, colCount]];
            titleRange.Merge();
            titleRange.Value2 = fileName;
            titleRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            titleRange.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            titleRange.Font.Size = 28;
            titleRange.Font.Bold = true;
            titleRange.RowHeight = 80;
            titleRange.Interior.Color =
                ColorTranslator.ToOle(Color.LightGray);

            // =======================
            // Header
            // =======================
            object[,] headers = new object[1, colCount];
            for (int c = 0; c < colCount; c++)
                headers[0, c] = visibleCols[c].HeaderText;

            var headerRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, colCount]];
            headerRange.Value2 = headers;
            headerRange.Font.Bold = true;
            headerRange.HorizontalAlignment =
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            headerRange.Interior.Color =
                ColorTranslator.ToOle(Color.Gainsboro);

            // =======================
            // Write Data
            // =======================
            object[,] data = new object[rowCount, colCount];
            int rIndex = 0;

            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                for (int c = 0; c < colCount; c++)
                {
                    var val = row.Cells[visibleCols[c].Name].Value;
                    if (val is DateTime dt)
                        data[rIndex, c] = dt.ToString("dd/MM/yyyy");
                    else
                        data[rIndex, c] = val?.ToString() ?? "";
                }
                rIndex++;
            }

            var start = sheet.Cells[3, 1];
            var end = sheet.Cells[rowCount + 2, colCount];
            var writeRange = sheet.Range[start, end];
            writeRange.Value2 = data;
            writeRange.HorizontalAlignment =
                Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            // =======================
            // Apply ROW COLORS (Aqua)
            // =======================
            int excelRow = 3;
            foreach (DataGridViewRow gridRow in advancedDataGridView1.Rows)
            {
                if (gridRow.IsNewRow) continue;

                Color rowColor = gridRow.DefaultCellStyle.BackColor;

                if (rowColor != Color.Empty && rowColor != Color.White)
                {
                    var excelRowRange = sheet.Range[
                        sheet.Cells[excelRow, 1],
                        sheet.Cells[excelRow, colCount]
                    ];

                    excelRowRange.Interior.Color =
                        ColorTranslator.ToOle(rowColor);
                }

                excelRow++;
            }

            // =======================
            // Apply EXPIRY DATE CELL COLORS (RED / YELLOW)
            // =======================
            int expiryColIndex = visibleCols
                .FindIndex(c => c.Name == "Contract_expiry_date") + 1;

            excelRow = 3;

            foreach (DataGridViewRow gridRow in advancedDataGridView1.Rows)
            {
                if (gridRow.IsNewRow) continue;

                if (expiryColIndex > 0)
                {
                    var gridCell = gridRow.Cells["Contract_expiry_date"];
                    Color cellColor = gridCell.Style.BackColor;

                    if (cellColor != Color.Empty && cellColor != Color.White)
                    {
                        var excelCell = (Microsoft.Office.Interop.Excel.Range)
                            sheet.Cells[excelRow, expiryColIndex];

                        excelCell.Interior.Color =
                            ColorTranslator.ToOle(cellColor);

                        excelCell.Font.Color =
                            ColorTranslator.ToOle(gridCell.Style.ForeColor);
                    }
                }

                excelRow++;
            }

            // =======================
            // Formatting
            // =======================
            sheet.Cells.Font.Size = 14;
            sheet.Cells.VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            sheet.Columns.AutoFit();
            sheet.Rows.AutoFit();
            sheet.Rows.RowHeight = 25;

            int totalRows = rowCount + 2;
            var fullRange = sheet.Range[
                sheet.Cells[1, 1],
                sheet.Cells[totalRows, colCount]
            ];

            fullRange.Borders.LineStyle =
                Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            fullRange.Borders.Weight =
                Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            try
            {
                workbook.SaveAs(fullPath);
                excelApp.Visible = true;
                ShowAlert("تم تصدير الملف بنجاح", AlertForm.AlertType.Success);
                this.FindForm().WindowState = FormWindowState.Minimized;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"حدث خطأ أثناء التصدير:\n{ex.Message}",
                    "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void guna2Button3_Click(object sender, EventArgs e)
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
                if (SessionData.Investment_Type == "استثمارات على الطرق")
                {
                    LoadInvestmentDataForRoads();
                }
                if (SessionData.Investment_Type == "محلات شل اوت داخل")
                {
                    LoadInvestmentDataMarketsIn();
                }
                if (SessionData.Investment_Type != "محلات شل اوت داخل"
               && SessionData.Investment_Type != "استثمارات على الطرق")
                {
                    LoadInvestmentData();
                }
                ShowAlert($"{deleted} صف تم حذفه بنجاح", AlertForm.AlertType.Success);
            }
            else
            {
                ShowAlert("لم يتم تحديد أي صفوف للحذف", AlertForm.AlertType.Warning);
            }
            UpdateRowCount();
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
            if (string.IsNullOrWhiteSpace(Activity_Type_TB1.Text))
            {
                Error_Type.Visible = true;
                Activity_Type_TB1.Focus();
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
            if (InLand_RA.Checked) {
                var row = this.investmentsTableAdapter.GetData()
               .OrderByDescending(r => Convert.ToInt32(r.investments_id))
               .FirstOrDefault();

                int lastId = row != null ? Convert.ToInt32(row.investments_id) : 0;
                int newId = lastId + 1;
                if (!activityLookup.ContainsValue(Activity_Type_TB1.Text))
                {
                    MessageBox.Show("الرجاء اختيار نوع النشاط من القائمة فقط.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Error_Type.Visible = true;
                    return;
                }
                this.investmentsTableAdapter.Insert(
                    (newId).ToString(),
                    Location_TB.Text,
                    InvestName_TB.Text,
                    igover.First().governorate_id,
                    Land_ID_COB.Text,
                    Dependent_neighborhood_TB.Text,
                    Activity_Type_TB1.Text,
                    Activity_Name_TB.Text,
                    Place_number_Tb.Text,
                    Offer_memorandum_Number_TB.Text,
                    Contract_number_TB.Text,
                    Start.Value,
                    Expiray.Value,
                    decimal.Parse(Rental_value_TB.Text),
                    OfferfilePath,
                    ContractFilePath,
                    guna2DateTimePicker1.Value,
                    guna2ComboBox1.Text,
                    "",
                    1,
                    true,"","",""
                    );
                if (SessionData.Investment_Type == "استثمارات على الطرق")
                {
                    LoadInvestmentDataForRoads();
                }
                if (SessionData.Investment_Type == "محلات شل اوت داخل")
                {
                    LoadInvestmentDataMarketsIn();
                }
                if (SessionData.Investment_Type != "محلات شل اوت داخل"
               && SessionData.Investment_Type != "استثمارات على الطرق")
                {
                    LoadInvestmentData();
                }
                ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
                Error_Type.Visible = true;
                Governorate_COB.SelectedItem = -1;
                Land_ID_COB.Text = "";
                Dependent_neighborhood_TB.Text="";
                Activity_Type_TB1.Text="";
                Activity_Name_TB.Text="";
                Place_number_Tb.Text="";
                Offer_memorandum_Number_TB.Text="";
                Contract_number_TB.Text = "";
                Rental_value_TB.Text="";
                OfferfilePath=null;
                ContractFilePath=null;
                guna2ComboBox1.SelectedItem=-1;
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel7.Controls.Clear();
                UpdateRowCount();
            }
            else if (OutOfLand_RA.Checked)
            {
                string prefix = "";
                if (guna2ComboBox1.Text == "طرق")
                    prefix = "&";
                else if (guna2ComboBox1.Text == "كباري")
                    prefix = "#";
                else if (guna2ComboBox1.Text == "مواقف")
                    prefix = "*";

                int lastId = this.investmentsTableAdapter.GetData()
                 .Select(r =>
                     int.TryParse(r.investments_id?.TrimStart('@', '#', '*'), out int id)
                         ? id
                         : 0
                 )
                 .DefaultIfEmpty(0)  // ✅ يمنع الخطأ لو مفيش أي صفوف
                 .Max();

                int newId = lastId + 1;
                // land وهمي
                string landId = prefix + newId.ToString();

                // أولاً: إضافة land وهمي
                this.landsTableAdapter.Insert(
                    landId,
                    " ", // land_number
                    "لا يوجد ارض" + guna2ComboBox1.Text, // land_name (تمييز فقط)
                    " ", // total_area
                    " ", 
                    " ", 
                    " ",
                    " ", 
                    " ", 
                    igover.First().governorate_id, " ",
                    " ", // governorate_fk
                    " ",
                    " ", 
                    " ",
                    0, // total_Land_Price
                    " ", // Ownership_Authority
                    " ", // Address
                    0,"","","" ,""// price_per_meter
                );
                if (!activityLookup.ContainsValue(Activity_Type_TB1.Text))
                {
                    MessageBox.Show("الرجاء اختيار نوع النشاط من القائمة فقط.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Error_Type.Visible = true;
                    return;
                }
                this.investmentsTableAdapter.Insert(
                (newId).ToString(),
                Location_TB.Text,
                InvestName_TB.Text,
                igover.First().governorate_id,
                landId,  // الربط بالـ land_id الوهمي
                Dependent_neighborhood_TB.Text,
                Activity_Type_TB1.Text,
                Activity_Name_TB.Text,
                Place_number_Tb.Text,
                Offer_memorandum_Number_TB.Text,
                Contract_number_TB.Text,
                Start.Value,
                Expiray.Value,
                decimal.Parse(Rental_value_TB.Text),
                OfferfilePath,
                ContractFilePath,
                guna2DateTimePicker1.Value,
                guna2ComboBox1.Text,
                "",
                1,
                true,"","",""
                 );
                if (SessionData.Investment_Type == "استثمارات على الطرق")
                {
                    LoadInvestmentDataForRoads();
                }
                if (SessionData.Investment_Type == "محلات شل اوت داخل")
                {
                    LoadInvestmentDataMarketsIn();
                }
                if (SessionData.Investment_Type != "محلات شل اوت داخل"
               && SessionData.Investment_Type != "استثمارات على الطرق")
                {
                    LoadInvestmentData();
                }
                ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
                Error_Type.Visible = true;
                UpdateRowCount();
                Governorate_COB.SelectedItem = -1;
                Land_ID_COB.Text = "";
                Dependent_neighborhood_TB.Text = "";
                Activity_Type_TB1.Text = "";
                Activity_Name_TB.Text = "";
                Place_number_Tb.Text = "";
                Offer_memorandum_Number_TB.Text = "";
                Contract_number_TB.Text = "";
                Rental_value_TB.Text = "";
                OfferfilePath = null;
                ContractFilePath = null;
                guna2ComboBox1.SelectedItem = -1;
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel7.Controls.Clear();
            }
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
            var x = this.investmentsTableAdapter.GetDataByInvestId(serial_number_TB.Text);
            if (x == null || x.Count == 0)
            {
                ShowAlert("لا يوجد هذا البيان للتعديل", AlertForm.AlertType.Error);
                return;
            }
            if (!activityLookup.ContainsValue(Activity_Type_TB1.Text))
            {
                MessageBox.Show("الرجاء اختيار نوع النشاط من القائمة فقط.", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Error_Type.Visible = true;
                return;
            }
            this.investmentsTableAdapter.UpdateQuery(
                Location_TB.Text,
                InvestName_TB.Text,
                igover.First().governorate_id,
                Land_ID_COB.Text,
                Dependent_neighborhood_TB.Text,
                Activity_Type_TB1.Text,
                Activity_Name_TB.Text,
                Place_number_Tb.Text,
                Offer_memorandum_Number_TB.Text,
                Contract_number_TB.Text,
                Start.Value.ToString(),
                Expiray.Value.ToString(),
                decimal.Parse(Rental_value_TB.Text),
                OfferfilePath,
                ContractFilePath,
                guna2DateTimePicker1.Value.ToString(),
                guna2ComboBox1.Text,
                serial_number_TB.Text
                );
            if (SessionData.Investment_Type == "استثمارات على الطرق")
            {
                LoadInvestmentDataForRoads();
            }
            if (SessionData.Investment_Type == "محلات شل اوت داخل")
            {
                LoadInvestmentDataMarketsIn();
            }
            if (SessionData.Investment_Type != "محلات شل اوت داخل"
               && SessionData.Investment_Type != "استثمارات على الطرق")
            {
                LoadInvestmentData();
            }
            ShowAlert("تم التعديل بنجاح", AlertForm.AlertType.Success);
            Error_Type.Visible = false;
            Governorate_COB.SelectedItem = -1;
            Land_ID_COB.Text = "";
            Dependent_neighborhood_TB.Text = "";
            Activity_Type_TB1.Text = "";
            Activity_Name_TB.Text = "";
            Place_number_Tb.Text = "";
            Offer_memorandum_Number_TB.Text = "";
            Contract_number_TB.Text = "";
            Rental_value_TB.Text = "";
            OfferfilePath = null;
            ContractFilePath = null;
            guna2ComboBox1.SelectedItem = -1;
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
            serial_number_TB.Text ="";
        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
          
            advancedDataGridView1.CleanFilter();
            advancedDataGridView1.CleanSort();
            guna2Button1.Visible = Add_Radio.Checked;
            guna2Button1.Enabled = Add_Radio.Checked;
            guna2Button2.Visible = Update_Radio.Checked;
            guna2Button2.Enabled = Update_Radio.Checked;
            guna2Button3.Parent = guna2TabControl1.Parent; // Not inside the tab page
            guna2Button3.BringToFront();
            guna2Button3.Size = new Size(186, guna2TabControl1.ItemSize.Height - 1);
            PositionHeaderButton();
            loadcomboxes();
            if (SessionData.Investment_Type == "كل الاستثمارات")
            {
                LoadInvestmentDataAll();
                //checkedListBox2.Visible = true;
                //LoadRoadsToCheckedListBox();
            }
            if (SessionData.Investment_Type == "استثمارات على الطرق")
            {
                LoadInvestmentDataForRoads();
                checkedListBox2.Visible = true;
                LoadRoadsToCheckedListBox();
            }
            if (SessionData.Investment_Type == "محلات شل اوت داخل")
            {
                LoadInvestmentDataMarketsIn();
                checkedListBox2.Visible = false;
                dreamButton4.Visible = true;
            }
            if (SessionData.Investment_Type == "مولات")
            {
                LoadInvestmentDataMALLS();
                checkedListBox2.Visible = false;
                dreamButton4.Visible = true;
            }
            if (SessionData.Investment_Type != "محلات شل اوت داخل" 
                && SessionData.Investment_Type != "مولات"
                 && SessionData.Investment_Type != "استثمارات على الطرق"
                 && SessionData.Investment_Type != "كل الاستثمارات")
            {
                checkedListBox2.Visible = false;
                LoadInvestmentData();
                
            }
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
            this.advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            Governorate_COB.SelectedItem = -1;
            Land_ID_COB.Text = "";
            Dependent_neighborhood_TB.Text = "";
            Activity_Type_TB1.Text = "";
            Activity_Name_TB.Text = "";
            Place_number_Tb.Text = "";
            Offer_memorandum_Number_TB.Text = "";
            Contract_number_TB.Text = "";
            Rental_value_TB.Text = "";
            OfferfilePath = null;
            ContractFilePath = null;
            guna2ComboBox1.SelectedItem = -1;
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();

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
        string ContractFilePath = " "; 
        string OfferfilePath = " "; 
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
                        if (column.Name == "OfferfilePaths" 
                        || column.Name == "ContractFilePaths" 
                        || column.Name == "Project_Id"
                        || column.Name == "Visible"
                        || column.Name =="investment_type"
                        || column.Name == "Dependent_neighborhood"
                        || column.Name == "Model8File"
                        || column.Name == "CivilDefenseFile")

                            continue;
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
        }
        private void LoadInvestmentDataForUpdatePart(string serialNumber)
        {
            bool hasText = !string.IsNullOrWhiteSpace(serialNumber);
            if (!hasText) return;

            var x = this.investmentsTableAdapter.GetDataByInvestId(serialNumber);

            if (x == null || x.Count == 0)
            {
                // No data found — clear all fields
                Location_TB.Text = "";
                Dependent_neighborhood_TB.Text = "";
                Activity_Type_TB1.Text = "";
                Activity_Name_TB.Text = "";
                Governorate_COB.SelectedIndex = -1;
                Place_number_Tb.Text = "";
                Offer_memorandum_Number_TB.Text = "";
                Contract_number_TB.Text = "";
                Land_ID_COB.SelectedIndex = -1;
                Land_Name_COB.SelectedIndex = -1;
                plate_number_COB.SelectedIndex = -1;
                Rental_value_TB.Text = "";
                guna2ComboBox1.Text = "";
                flowLayoutPanel1.Controls.Clear();
                flowLayoutPanel7.Controls.Clear();
                return;
            }

            var first = x.First();
            var l = this.landsTableAdapter.GetDataBySerial(first.land_fk);

            // نصوص
            Location_TB.Text = first.Location;
            Dependent_neighborhood_TB.Text = first.Dependent_neighborhood;
            Activity_Type_TB1.Text = first.Activity_Type;
            Activity_Name_TB.Text = first.Activity_Name;
            Place_number_Tb.Text = first.Place_number;
            Offer_memorandum_Number_TB.Text = first.Offer_memorandum_number;
            Contract_number_TB.Text = first.Contract_number;
            Rental_value_TB.Text = first.Rental_value.ToString();
            guna2ComboBox1.SelectedItem = first.investment_type;
            // Governorate
            Governorate_COB.SelectedItem = governorateLookup.TryGetValue(first.governorate_fk.ToString(), out var govName)
                ? govName
                : "غير معروف";

            // Land
            if (l.Any() && first.land_fk == l.First().land_id)
            {
                Land_Name_COB.SelectedItem = l.First().land_name;
                plate_number_COB.SelectedItem = l.First().plate_number;
            }
            Land_ID_COB.SelectedItem = first.land_fk;

            // ملفات
            ContractFilePath = first.Contract_number_File;
            OfferfilePath = first.Offer_memorandum_number_File;

            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();

            if (!string.IsNullOrWhiteSpace(ContractFilePath))
            {
                AddFileIconToPanel(ContractFilePath, Path.GetFileName(ContractFilePath));
            }
            if (!string.IsNullOrWhiteSpace(OfferfilePath))
            {
                lastAddedFilePaths.Add(OfferfilePath);
                AddFileIconToPanel1(OfferfilePath, Path.GetFileName(OfferfilePath));
            }
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
            try
            {
                if (originalData == null)
                    return;

                // Create a DataView based on the original data
                DataView view = new DataView(originalData);
                view.RowFilter = advancedDataGridView1.FilterString;

                // Apply to grid
                advancedDataGridView1.DataSource = view;
                originalData = view.ToTable();
                UpdateRowCount();
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ أثناء تطبيق الفلتر: " + ex.Message, AlertForm.AlertType.Error);
            }
        }
        private Dictionary<string, string> governorateLookup = new Dictionary<string, string>();
        private Dictionary<string, string> activityLookup = new Dictionary<string, string>();
        private DataTable LandData;
        private DataTable GoverData;
        private DataTable Activity;
        private void serial_number_TB_TextChanged(object sender, EventArgs e)
        {

            
            //var x = this.investmentsTableAdapter.GetDataByInvestId(serial_number_TB.Text);
           
            //if (x == null || x.Count == 0)
            //{
            //    // No data found — clear all fields
            //    Location_TB.Text = "";
            //    Dependent_neighborhood_TB.Text = "";
            //    Activity_Type_TB.Text = "";
            //    Activity_Name_TB.Text = "";
            //    Governorate_COB.Text = "";
            //    Place_number_Tb.Text = "";
            //    Offer_memorandum_Number_TB.Text = "";
            //    Contract_number_TB.Text = "";
            //    Land_ID_COB.Text = "";
            //    Land_Name_COB.Text = "";
            //    plate_number_COB.Text = "";
            //    flowLayoutPanel1.Controls.Clear();
            //    flowLayoutPanel7.Controls.Clear();
            //    return;
            //}
            //var l = this.landsTableAdapter.GetDataBySerial(x.First().land_fk);
            //flowLayoutPanel1.Controls.Clear();
            //flowLayoutPanel7.Controls.Clear();
            //Location_TB.Text = x.First().Location;
            //Dependent_neighborhood_TB.Text = x.First().Dependent_neighborhood;
            //Activity_Type_TB.Text = x.First().Activity_Type;
            //Activity_Name_TB.Text = x.First().Activity_Name;
            //Governorate_COB.SelectedItem = governorateLookup.TryGetValue
            //(x.First().governorate_fk.ToString(), out var govName)
            //? govName
            //: "غير معروف";
            //Place_number_Tb.Text = x.First().Place_number;
            //Offer_memorandum_Number_TB.Text = x.First().Offer_memorandum_number;
            //Contract_number_TB.Text = x.First().Contract_number;
            //if (x.First().land_fk == l.First().land_id)
            //{
            //    Land_Name_COB.SelectedItem = l.First().land_name;
            //    plate_number_COB.SelectedItem = l.First().plate_number;
            //}
            //Land_ID_COB.SelectedItem = x.First().land_fk;
            //Rental_value_TB.Text = x.First().Rental_value.ToString();
            //ContractFilePath = x.First().Contract_number_File;
            
            //OfferfilePath = x.First().Offer_memorandum_number_File;
            
            //if (!string.IsNullOrWhiteSpace(ContractFilePath))
            //{
            //    //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
            //    AddFileIconToPanel(ContractFilePath, Path.GetFileName(ContractFilePath));
            //}
            //if (!string.IsNullOrWhiteSpace(OfferfilePath))
            //{
            //    lastAddedFilePaths.Add(OfferfilePath);
            //    //flowLayoutPanel1.Controls.Clear(); // Clear previous icons if needed
            //    AddFileIconToPanel1(OfferfilePath, Path.GetFileName(OfferfilePath));
            //}

        }

        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {
            //if (SessionData.Investment_para1 != null && SessionData.Investment_Type == "استثمارات على الطرق")
            //{
            //    string name = SessionData.UserName;
            //    long userId = SessionData.UserId;
            //    ENGReportForm menu = new ENGReportForm(name, userId);
            //    menu.Show();
            //    menu.ShowMenuView2();
            //    Form parentForm = this.FindForm();
            //    if (parentForm != null)
            //    {
            //        parentForm.Close(); // or parentForm.Hide(); if you just want to hide it
            //    }
            //}
            if (SessionData.Investment_Type == "محلات شل اوت خارج")
            {
                string name = SessionData.UserName;
                long userId = SessionData.UserId;
                ENGReportForm menu = new ENGReportForm(name, userId);
                menu.Show();
                menu.ShowMenuView3();
                Form parentForm = this.FindForm();
                if (parentForm != null)
                {
                    parentForm.Close(); // or parentForm.Hide(); if you just want to hide it
                }
            }
            else
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
            Land_ID_COB.Text = "";
            Dependent_neighborhood_TB.Text = "";
            Activity_Type_TB1.Text = "";
            Activity_Name_TB.Text = "";
            Place_number_Tb.Text = "";
            Offer_memorandum_Number_TB.Text = "";
            Contract_number_TB.Text = "";
            Rental_value_TB.Text = "";
            OfferfilePath = null;
            ContractFilePath = null;
            guna2ComboBox1.SelectedItem = -1;
            flowLayoutPanel1.Controls.Clear();
            flowLayoutPanel7.Controls.Clear();
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

        private void skyButton4_Click(object sender, EventArgs e)
        {
            //reportViewer1.Visible = true;
            //reportViewer1.LocalReport.DataSources.Clear();
            //DataTable original = ((DataView)((BindingSource)advancedDataGridView1.DataSource).List).ToTable();
            //DataTable filtered = new DataTable();

            //foreach (string headerText in checkedListBox1.CheckedItems)
            //{
            //    if (columnMap.ContainsKey(headerText)) // map header → real column
            //    {
            //        string colName = columnMap[headerText];
            //        filtered.Columns.Add(colName, original.Columns[colName].DataType);
            //    }
            //}

            //foreach (DataRow row in original.Rows)
            //{
            //    var newRow = filtered.NewRow();
            //    foreach (string headerText in checkedListBox1.CheckedItems)
            //    {
            //        if (columnMap.ContainsKey(headerText))
            //        {
            //            string colName = columnMap[headerText];
            //            newRow[colName] = row[colName];
            //        }
            //    }
            //    filtered.Rows.Add(newRow);
            //}

            //string rdlc = GenerateDynamicRDLC(filtered);

            //using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(rdlc)))
            //{
            //    reportViewer1.LocalReport.LoadReportDefinition(stream);
            //}

            //reportViewer1.LocalReport.DataSources.Clear();
            //reportViewer1.LocalReport.DataSources.Add(
            //    new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", filtered));//change this data set to new one
            //reportViewer1.RefreshReport();
            try
            {
                TrueFunction();

                if (originalData == null || originalData.Rows.Count == 0)
                {
                    ShowAlert("لا يوجد بيانات", AlertForm.AlertType.Error);
                    return;
                }

                // ------------------------------------------
                // Prepare Excel
                // ------------------------------------------
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fileName = $"بيان_موقف_الاستثمارات_{DateTime.Now:yyyy_MM_dd}.xlsx";
                string path = Path.Combine(desktop, fileName);

                var excel = new Microsoft.Office.Interop.Excel.Application();
                var book = excel.Workbooks.Add();
                var sheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;

                sheet.DisplayRightToLeft = true;
                sheet.Name = "التقرير";

                // ------------------------------------------
                // Helper function: analyze counts + shop sums
                // ------------------------------------------
                Func<string, (
                    int مؤجر, int غير_مؤجر, int توقيع, int عقود_مالية, int عقود_منتظرة, int اجمالي,
                    int محلات_مؤجرة, int محلات_غير_مؤجرة, int محلات_توقيع, int محلات_مالية, int محلات_منتظرة
                )> Analyze = (type) =>
                {
                    var rows = originalData.AsEnumerable()
                        .Where(row => row.Field<string>("investment_type") == type && !row.Field<string>("Land_Id").Contains("$"))
                        .ToList();

        // -------- Counters --------
                    var rentedRows = rows.Where(rr => rr.Field<string>("Rental_Status") == "مؤجر" ||
                    rr.Field<string>("Rental_Status") == "مؤجر (جاري التوقيع علي العقد)"||
                    rr.Field<string>("Rental_Status") == "مؤجر (بدون عقد)");

                    var not_rentedShops = rows.Where(rr => rr.Field<string>("Rental_Status") != "مؤجر" &&
                    rr.Field<string>("Rental_Status") != "مؤجر (جاري التوقيع علي العقد)" &&
                    rr.Field<string>("Rental_Status") != "منتظر العقد" &&
                    rr.Field<string>("Rental_Status") != "مؤجر (بدون عقد)");

                    var shops_waited = rows.Where(rr => rr.Field<string>("Rental_Status") == "منتظر العقد");

                    var paidRows = rows.Where(rr => rr.Field<string>("Rental_Status") == "منتظر العقد");

                    var waitingRows = rows.Where(rr =>
                        rr.Field<string>("Activity_Name") == "لا يوجد" &&
                        rr.Field<string>("Rental_Status") != "مؤجر" &&
                        rr.Field<string>("Rental_Status") != "مؤجر (جاري التوقيع علي العقد)"
                    );

        // -------- Shop Sums --------
                    int sumRented = rentedRows.Sum(rr => rr.Field<int?>("Shops_Count")  ?? 0);
                    int sumEmpty = not_rentedShops.Sum(rr => rr.Field<int?>("Shops_Count") ?? 0);
                    int sumFixing = shops_waited.Sum(rr => rr.Field<int?>("Shops_Count") ?? 0);
                    int sumPaid = paidRows.Sum(rr => rr.Field<int?>("Shops_Count") ?? 0);
                    int sumWaiting = waitingRows.Sum(rr => rr.Field<int?>("Shops_Count") ?? 0);

                    return (
                        rentedRows.Count(),
                        not_rentedShops.Count(),
                        shops_waited.Count(),
                        paidRows.Count(),
                        waitingRows.Count(),
                        rows.Count,
                        sumRented,
                        sumEmpty,
                        sumFixing,
                        sumPaid,
                        sumWaiting
                    );
                };

                // ------------------------------------------
                // Helper: draw each block
                // ------------------------------------------
                Action<int, string, string> DrawBlock = (rowStart, title, type) =>
                {
                    var result = Analyze(type);
                    int row = rowStart;

                    // ====== TITLE ======
                    var titleRange = (Microsoft.Office.Interop.Excel.Range)
                        sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 7]];
                    titleRange.Merge();
                    titleRange.Value2 = title;
                    titleRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(198, 224, 180));
                    titleRange.Font.Bold = true;
                    titleRange.Font.Size = 16;
                    titleRange.HorizontalAlignment =
                        Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    row++;

                    // ====== HEADERS ======
                    string[] headers =
                    {
            "محلات مؤجرة",
            "محلات غير مؤجرة",
            "محلات جاري توقيعها",
            "عقود منتظرة من المالية",
            "عقود مؤجرة",
            "الإجمالي",
            "ملاحظات"
        };

                    for (int c = 1; c <= headers.Length; c++)
                    {
                        var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, c];
                        cell.Value2 = headers[c - 1];
                        cell.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                        cell.Font.Bold = true;
                        cell.HorizontalAlignment =
                            Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    }
                    row++;

                    // ====== DATA (Counts) ======
                    //((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 1]).Value2 = result.اجمالي;
                    //((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 2]).Value2 = result.غير_مؤجر;
                    //((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 3]).Value2 = result.توقيع;
                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 4]).Value2 = result.عقود_مالية;
                    //result.عقود_منتظرة
                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 5]).Value2 = result.مؤجر;
                ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 6]).Value2 = result.عقود_مالية + result.مؤجر;
                ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 7]).Value2 = "إجمالي العقود (العقود الموقعة تتضمن عدد محلات فى عقد واحد)";

                    var dataRange = (Microsoft.Office.Interop.Excel.Range)
                        sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 6]];
                    dataRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 242, 204));

                    row++;

                    // ====== DATA (Shop Sums) ======
                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 1]).Value2 = result.محلات_مؤجرة;

                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 2]).Value2 = result.محلات_غير_مؤجرة;
                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 3]).Value2 = result.محلات_توقيع;
                    //((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 4]).Value2 = result.محلات_مالية;
                    //((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 5]).Value2 = result.محلات_منتظرة;
                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 6]).Value2 = result.محلات_توقيع + result.محلات_غير_مؤجرة+ result.محلات_مؤجرة;
                    ((Microsoft.Office.Interop.Excel.Range)sheet.Cells[row, 7]).Value2 = "إجمالي المحلات";

                    var dataRangeShops = (Microsoft.Office.Interop.Excel.Range)
                        sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 6]];
                    dataRangeShops.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(255, 230, 153));

                    // Borders for block
                    var blockRange = (Microsoft.Office.Interop.Excel.Range)
                        sheet.Range[sheet.Cells[rowStart, 1], sheet.Cells[row, 7]];
                    blockRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;

                  ((Microsoft.Office.Interop.Excel.Range)sheet.Rows[rowStart]).RowHeight = 35;
                };

                // ------------------------------------------
                // Report Title
                // ------------------------------------------
                int r = 1;

                var mainTitle = (Microsoft.Office.Interop.Excel.Range)
                    sheet.Range[sheet.Cells[r, 1], sheet.Cells[r, 7]];
                mainTitle.Merge();
                mainTitle.Value2 = $"بيان موقف الاستثمارات عن يوم {DateTime.Now:yyyy/MM/dd}";
                mainTitle.Font.Size = 22;
                mainTitle.Font.Bold = true;
                mainTitle.HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                r += 2;

                // ------------------------------------------
                // Blocks
                // ------------------------------------------
                DrawBlock(r, "محلات شل أوت خارج", "محلات شل اوت خارج");
                r += 5;

                DrawBlock(r, "محلات شل أوت داخل", "محلات شل اوت داخل");
                r += 5;

                DrawBlock(r, "اسفل الكباري", "اسفل كباري");
                r += 5;

                DrawBlock(r, "استثمارات على الطرق", "استثمارات على الطرق");
                r += 5;

                sheet.Columns.AutoFit();

                book.SaveAs(path);
                excel.Visible = true;

                ShowAlert("تم إنشاء التقرير على سطح المكتب", AlertForm.AlertType.Success);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                FalseFunction();
            }


        }
        private string GenerateDynamicRDLC(DataTable dt)
        {
            using (var ms = ReportHelperEnhanced.GenerateDynamicRDLC(dt, columnMap, "DataSet1",Report_TB.Text))
            {
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        private void OutOfLand_RA_CheckedChanged(object sender, EventArgs e)
        {
            Land_ID_COB.Enabled = !OutOfLand_RA.Checked;
            plate_number_COB.Enabled = !OutOfLand_RA.Checked;
            Land_Name_COB.Enabled = !OutOfLand_RA.Checked;
            if (OutOfLand_RA.Checked) 
            {
                LoadGovernorates();
            }

        }

        private void InLand_RA_CheckedChanged(object sender, EventArgs e)
        {
            Land_ID_COB.Enabled = InLand_RA.Checked;
            plate_number_COB.Enabled = InLand_RA.Checked;
            Land_Name_COB.Enabled = InLand_RA.Checked;
            if (InLand_RA.Checked)
            {
                loadcomboxes();
            }
        }

        private void Land_Name_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Land_Name_COB.SelectedItem != null)
                SyncSelection(selectedName: Land_Name_COB.SelectedItem.ToString());
        }

        private void plate_number_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (plate_number_COB.SelectedItem != null)
                SyncSelection(selectedPlate: plate_number_COB.SelectedItem.ToString());
        }

        private void Land_ID_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Land_ID_COB.SelectedItem != null)
                SyncSelection(selectedLandId: Land_ID_COB.SelectedItem.ToString());
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
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var clickedColumn = advancedDataGridView1.Columns[e.ColumnIndex];
                // make sure user clicked a valid row and on the Activity_Name column
                if (e.RowIndex >= 0 && advancedDataGridView1.Columns[e.ColumnIndex].Name == "Activity_Name")
                {
                    string activityName = advancedDataGridView1.Rows[e.RowIndex].Cells["Activity_Name"].Value?.ToString()+".PDF";
                    string projectId = advancedDataGridView1.Rows[e.RowIndex].Cells["Project_Id"].Value?.ToString();

                    if (string.IsNullOrWhiteSpace(activityName) || string.IsNullOrWhiteSpace(projectId))
                        return;

                    // Get all docs for that project
                    var docs = this.documentsTableAdapter.GetDataByProjectOnly(projectId);
                                

                    var match = docs.FirstOrDefault(d =>
                        !d.IspathsNull() &&
                        Path.GetFileName(d.paths).Equals(activityName, StringComparison.OrdinalIgnoreCase)
                    );

                    if (match != null)
                    {
                        string path = match.paths;
                        if (File.Exists(path))
                        {
                            System.Diagnostics.Process.Start(path);
                            Form parentForm = this.FindForm();
                            parentForm.WindowState = FormWindowState.Minimized;
                        }
                        else
                        {
                            if (path.StartsWith(@"Z:\", StringComparison.OrdinalIgnoreCase))
                            {
                                // اسم الفولدر اللي عايز تضيفه
                                string folderName = "sho8l";

                                // باقي المسار بدون Z:\
                                string relativePath = path.Substring(3);

                                // المسار الجديد
                                string newPath = Path.Combine(@"D:\", folderName, relativePath);

                                if (File.Exists(newPath))
                                {
                                    System.Diagnostics.Process.Start(newPath);
                                    Form parentForm = this.FindForm();
                                    parentForm.WindowState = FormWindowState.Minimized;
                                }
                                else
                                {
                                    MessageBox.Show("الملف غير موجود في المسار");
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("لا يوجد ملف مطابق لهذا النشاط والمشروع", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                if (e.RowIndex >= 0 && advancedDataGridView1.Columns[e.ColumnIndex].Name == "Description_Drawing_Place")
                {
                    string activityName = advancedDataGridView1.Rows[e.RowIndex].Cells["Description_Drawing_Place"].Value?.ToString();
                    string projectId = advancedDataGridView1.Rows[e.RowIndex].Cells["Project_Id"].Value?.ToString();

                    if (string.IsNullOrWhiteSpace(activityName) || string.IsNullOrWhiteSpace(projectId))
                        return;

                    // Get all docs for that project
                    var docs = this.documentsTableAdapter.GetDataByProjectOnly(projectId);


                    var match = docs.FirstOrDefault(d =>
                        !d.IspathsNull() &&
                        Path.GetFileName(d.paths).Equals(activityName, StringComparison.OrdinalIgnoreCase)
                    );

                    if (match != null)
                    {
                        string path = match.paths;
                        if (File.Exists(path))
                        {
                            System.Diagnostics.Process.Start(path);
                            Form parentForm = this.FindForm();
                            parentForm.WindowState = FormWindowState.Minimized;
                        }
                        else
                        {
                            MessageBox.Show("الملف غير موجود في المسار المحدد", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    else
                    {
                        MessageBox.Show("لا يوجد ملف مطابق لهذا النشاط والمشروع", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
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
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ: " + ex.Message, AlertForm.AlertType.Error);
            }
            if (Update_Radio.Checked)
            {
                if (e.RowIndex >= 0)
                {
                    var serialNumber = advancedDataGridView1.Rows[e.RowIndex]
                                         .Cells["investments_id"].Value?.ToString();

                    if (!string.IsNullOrEmpty(serialNumber))
                    {
                        serial_number_TB.Text = serialNumber;
                        LoadInvestmentDataForUpdatePart(serialNumber);
                    }
                }
            }

        }
        public class DocInfo
        {
            public string FileName { get; set; }
            public string FullPath { get; set; }
            public int DocumentID { get; set; }
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
        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void advancedDataGridView1_Scroll(object sender, ScrollEventArgs e)
        {
            advancedDataGridView1.Invalidate();
        }

        private void Activity_Type_TB_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }

        private void Activity_Type_TB_TextChanged(object sender, EventArgs e)
        {
            string typedText = Activity_Type_TB.Text.ToLower();
            if (string.IsNullOrWhiteSpace(typedText))
            {
                return;
            }

            // Find first match
            var item = Activity_Type_TB.Items
                         .Cast<object>()
                         .FirstOrDefault(x => x.ToString().ToLower().StartsWith(typedText));

            if (item != null)
            {
                int index = Activity_Type_TB.Items.IndexOf(item);
                Activity_Type_TB.SelectedIndex = index;

                // Select remaining part (auto-complete style)
                Activity_Type_TB.SelectionStart = typedText.Length;
                Activity_Type_TB.SelectionLength = Activity_Type_TB.Text.Length - typedText.Length;
            }
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            var lines = guna2TextBox1.Lines
           .Where(line => !string.IsNullOrWhiteSpace(line)) // skip empty lines
           .Select(line => line.Trim())
           .ToList();
            if (lines.Count == 0)
            {
                MessageBox.Show("ادخل علي الاقل نشاط واحد");
                return;
            }
            // Optional: Get last ID from DB and increment from there
            var data = this.activity_LookupTableAdapter.GetData(); // or start from 1
            int nextId = data.Any() ? data.Last().Id + 1 : 1; 
            foreach (var name in lines)
            {
                int id = nextId;
                var ExistedId = this.activity_LookupTableAdapter.GetDataByActType(name);
                if (ExistedId.Any())
                {
                    ShowAlert("غير مسموح بي تكرار نفس النشاط", AlertForm.AlertType.Error);
                    return;
                }
                this.activity_LookupTableAdapter.Insert(id, name); // adjust to your adapter
                nextId++;
            }
            MessageBox.Show("تمت اضافة النشاط بنجاح");
            this.advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            guna2TextBox1.Clear();
        }

        private void advancedDataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {

        }

        private void guna2TabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage4 && Add_Radio.Checked == false && Update_Radio.Checked == false)
                e.Cancel = true;
            else if (e.TabPage == tabPage4 && (Add_Radio.Checked == true || Update_Radio.Checked == true))
            {
                e.Cancel = false;
            }
        }

        private void dreamButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void guna2CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
         
        }

        private void guna2CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void guna2CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            
        }
        bool isCheckedAll = true;
        bool isCheckedRented = false;
        bool isCheckedNotRented = false;
        bool IsWaitedDocs = false;
        private void dreamButton1_Click_1(object sender, EventArgs e)
        {
            isCheckedAll = !isCheckedAll;
            if (isCheckedAll)
            {
                dreamButton1.BackColor = Color.LightGreen;
                dreamButton1.Text = "الكل ✅";
                dreamButton2.BackColor = Color.LightGray;
                dreamButton2.Text = "مؤجر ☐";
                dreamButton3.BackColor = Color.LightGray;
                dreamButton3.Text = "غير مؤجر ☐";
                dreamButton6.BackColor = Color.LightGray;
                dreamButton6.Text = "منتظر العقد ☐";
                isCheckedRented = false;
                isCheckedNotRented = false;
                IsWaitedDocs = false;

                if (originalData != null)
                {
                    DataView view = new DataView(originalData);
                    view.RowFilter = "[Rental_Status] = 'مؤجر' " +
                        "OR[Rental_Status] = 'مؤجر (بدون عقد)' " +
                        "OR [Rental_Status] = 'منتظر العقد' " +
                        "OR [Rental_Status] = 'غير مؤجر (فارغ)' " +
                        "OR [Rental_Status] = 'غير مؤجر (تم الفسخ)' " +
                        "OR [Rental_Status] = 'منتظر التوقيع' " +
                        "OR [Rental_Status] = 'تابع ادارة شل اوت' " +
                        "OR [Rental_Status] = 'غير مؤجر (*)' " +
                        "OR [Rental_Status] ='جاري التجديد' " +
                        "OR [Rental_Status] ='مؤجر (جاري التوقيع علي العقد)' " + 
                        "OR [Rental_Status] ='مؤجر (*)' " +
                        "AND [Visible] = true";
                    advancedDataGridView1.DataSource = view;
                }
                loadcomboxes();
                UpdateRowCount();
                GenerativePanalFlow();
                ApplyGuna2StyleToGrid(advancedDataGridView1);
                ApplyGuna2StyleToGrid(advancedDataGridView2);
                ArabicColumnGrid();
                LoadColumnsIntoCheckedListBox();
            }
            else
            {
                dreamButton1.BackColor = Color.LightGray;
                dreamButton1.Text = "الكل ☐";
            }

        }

        private void dreamButton2_Click(object sender, EventArgs e)
        {
            isCheckedRented = !isCheckedRented;

            if (isCheckedRented)
            {
                try
                {
                    isCheckedAll = false;
                    isCheckedNotRented = false;
                    IsWaitedDocs = false;
                    dreamButton2.BackColor = Color.LightGreen;
                    dreamButton2.Text = "مؤجر ✅";
                    dreamButton1.BackColor = Color.LightGray;
                    dreamButton1.Text = "الكل ☐";
                    dreamButton3.BackColor = Color.LightGray;
                    dreamButton3.Text = "غير مؤجر ☐";
                    dreamButton6.BackColor = Color.LightGray;
                    dreamButton6.Text = "منتظر العقد ☐";

                    // ✅ Apply in-memory filter instead of reloading data
                    if (originalData != null)
                    {
                        DataView view = new DataView(originalData);
                        view.RowFilter = "[Rental_Status] = 'مؤجر' " +
                            "OR [Rental_Status] = 'منتظر التوقيع' " +
                            "OR [Rental_Status] = 'مؤجر (بدون عقد)' " +
                            "OR [Rental_Status] ='مؤجر (*)' " +
                            "OR [Rental_Status] ='مؤجر (جاري التوقيع علي العقد)' " + 
                            "AND [Visible] = true";
                        advancedDataGridView1.DataSource = view;
                    }
                }
                catch (Exception ex)
                {
                    ShowAlert("حدث خطأ غير متوقع: " + ex.Message, AlertForm.AlertType.Error);
                }
            }
            else
            {
                // if unchecked, reset back to all data
                advancedDataGridView1.DataSource = originalData;
            }

            // keep the rest of your functions
            advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
            FalseFunction();
            loadcomboxes();
            UpdateRowCount();
            GenerativePanalFlow();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ArabicColumnGrid();
            LoadColumnsIntoCheckedListBox();
        }

        private void dreamButton3_Click(object sender, EventArgs e)
        {
            isCheckedNotRented = !isCheckedNotRented;
            if (isCheckedNotRented)
            {
                isCheckedAll = false;
                isCheckedRented = false;
                IsWaitedDocs = false;
                dreamButton2.BackColor = Color.LightGreen;
                dreamButton2.Text = " مؤجر ☐";
                dreamButton1.BackColor = Color.LightGray;
                dreamButton1.Text = " الكل ☐";
                dreamButton3.BackColor = Color.LightGray;
                dreamButton3.Text = "غير مؤجر ✅";
                dreamButton6.BackColor = Color.LightGray;
                dreamButton6.Text = "منتظر العقد ☐";
                // ✅ Apply in-memory filter instead of reloading data
                if (originalData != null)
                {
                    DataView view = new DataView(originalData);
                    view.RowFilter = "[Rental_Status] = 'غير مؤجر (فارغ)' " +
                        "OR [Rental_Status] = 'غير مؤجر (تم الفسخ)' " +
                        "OR [Rental_Status] = 'غير مؤجر (*)' " +
                        "AND [Visible] = true";
                    advancedDataGridView1.DataSource = view;
                }
                else
                {
                    // if unchecked, reset back to all data
                    advancedDataGridView1.DataSource = originalData;
                }
                advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
                FalseFunction();
                loadcomboxes();
                UpdateRowCount();
                GenerativePanalFlow();
                ApplyGuna2StyleToGrid(advancedDataGridView1);
                ApplyGuna2StyleToGrid(advancedDataGridView2);
                ArabicColumnGrid();
                LoadColumnsIntoCheckedListBox();
            }
        }

        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {
            using (Graphics g = guna2TextBox2.CreateGraphics())
            {
                SizeF textSize = g.MeasureString(guna2TextBox2.Text, guna2TextBox2.Font);
                int padding = 65;
                guna2TextBox2.Width = (int)textSize.Width + padding;
                guna2TextBox2.TextAlign = HorizontalAlignment.Left;
            }
        }

        private void advancedDataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }
        private TextBox editingTextBox = null;
        private void advancedDataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // 1️⃣ Ignore invalid rows / columns
                if (e.RowIndex < 0 || e.ColumnIndex < 0)
                    return;

                var grid = advancedDataGridView1;
                string columnName = grid.Columns[e.ColumnIndex].Name;

                // Ignore checkbox / select column
                if (columnName.Equals("select", StringComparison.OrdinalIgnoreCase))
                    return;

                // 2️⃣ Get latest edited value
                string newValue = editingTextBox?.Text ??
                                  grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();

                editingTextBox = null;

                // 3️⃣ Get investment ID
                string investmentsID = grid.Rows[e.RowIndex]
                                           .Cells["investments_id"]
                                           .Value?.ToString();

                if (string.IsNullOrWhiteSpace(investmentsID))
                    return;

                // 4️⃣ Fetch row from DB
                var table = investmentsTableAdapter.GetDataByInvestId(investmentsID);
                if (table.Rows.Count == 0)
                {
                    MessageBox.Show("لم يتم العثور على السجل");
                    return;
                }

                var row = table[0];

                // 5️⃣ Handle empty value
                if (string.IsNullOrWhiteSpace(newValue))
                {
                    row[columnName] = DBNull.Value;
                }
                else
                {
                    // Clean invisible characters
                    newValue = newValue.Replace("\u200E", "").Trim();

                    // 6️⃣ Detect numeric columns automatically

                    Type columnType = row.Table.Columns[columnName].DataType;

                    if (columnType == typeof(decimal) ||
                        columnType == typeof(double) ||
                        columnType == typeof(float))
                    {
                        if (!decimal.TryParse(
                                newValue,
                                NumberStyles.Any,
                                CultureInfo.InvariantCulture,
                                out decimal numericValue))
                        {
                            MessageBox.Show("قيمة رقمية غير صالحة");
                            return;
                        }

                        row[columnName] = numericValue;
                    }
                    else if (columnType == typeof(int) || columnType == typeof(long))
                    {
                        if (!long.TryParse(newValue, out long intValue))
                        {
                            MessageBox.Show("قيمة رقمية صحيحة مطلوبة");
                            return;
                        }

                        row[columnName] = intValue;
                    }
                    else
                    {
                        // Text / other types
                        row[columnName] = newValue;
                    }
                }

                // 7️⃣ Save
                investmentsTableAdapter.Update(row);
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
        private void checkedListBox2_ItemCheck(object sender, ItemCheckEventArgs e)
        {
           
                this.BeginInvoke((MethodInvoker)(() =>
                {
                    if (SessionData.Investment_Type != "مناطق تنموية")
                    {
                        LoadInvestmentDataForSelectedRoads();
                    }
                    else if (SessionData.Investment_Type == "مناطق تنموية")
                    {
                        //this.BeginInvoke((MethodInvoker)(() =>
                        //{
                           // if (checkedListBox2.Items[e.Index].ToString() == "الكل")
                            //{
                              //  for (int i = 0; i < checkedListBox2.Items.Count; i++)
                                //    checkedListBox2.SetItemChecked(i, i == e.Index);
                            //}
                            //else
                            //{
                            //    int allIndex = checkedListBox2.Items.IndexOf("الكل");
                            //    if (allIndex >= 0)
                            //        checkedListBox2.SetItemChecked(allIndex, false);
                            //}

                            LoadInvestmentDataForSelectedRoads2();

                        //}));
                    }
                }));
        }
        private void checkedListBox2_MouseDown(object sender, MouseEventArgs e)
        {
           
        }
        private void checkedListBox2_MouseUp(object sender, MouseEventArgs e)
        {
            
        }

        private void dreamButton4_Click(object sender, EventArgs e)
        {
            try
            {
                // Load data from your TableAdapters
                var invList = this.invesments_ListTableAdapter.GetData();   // Invesments_List
                var investments = this.investmentsTableAdapter.GetData();   // investments
                var projects = this.projectsTableAdapter.GetData();         // projects

                // --- Step 1: Get all Name_Projects linked to investments starting with "!" ---
                var projectNames = (from i in investments
                                    join p in projects on i.land_fk equals p.land_fk into pj
                                    from p in pj.DefaultIfEmpty()
                                    where i.investments_id.StartsWith("!")
                                          && !p.IsName_ProjectsNull()
                                    select p.Name_Projects)
                                    .Distinct()
                                    .ToList();

                // --- Step 2: Get names in Invesments_List that are NOT in that list ---
                var missing = (from il in invList
                               where !il.IsName_InvestmentsNull()
                                     && !projectNames.Contains(il.Name_Investments)
                               select il.Name_Investments)
                               .ToList();

                // --- Step 3: Show results in MessageBox ---
                if (missing.Any())
                {
                    string msg = $"Found {missing.Count} missing item(s):\n\n" + string.Join("\n", missing);
                    MessageBox.Show(msg, "Missing in Projects", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("✅ All Name_Investments are matched with Name_Projects.",
                                    "Check Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dreamButton5_Click(object sender, EventArgs e)
        {
            try
            {
                // Step 1: Get all governorates from DB
                var allGovernorates = governorateTableAdapter.GetData()
                    .Where(g => g.governorate_id != "0")
                    .Select(g => new
                    {
                        g.governorate_id,
                        g.governorate
                    })
                    .ToList();

                // Step 2: Get governorates already present in current DataGridView
                var displayedGovernorates = new HashSet<string>(
                    from DataGridViewRow row in advancedDataGridView1.Rows
                    where row.Cells["GovernorateName"] != null &&
                          row.Cells["GovernorateName"].Value != null &&
                          !string.IsNullOrWhiteSpace(row.Cells["GovernorateName"].Value.ToString())
                    select row.Cells["GovernorateName"].Value.ToString()
                );

                // Step 3: Filter out the ones that exist in the DataGridView
                var missingGovernorates = allGovernorates
                    .Where(g => !displayedGovernorates.Contains(g.governorate))
                    .OrderBy(g => g.governorate)
                    .ToList();

                // Step 4: Show result (you can change this UI part as needed)
                if (missingGovernorates.Count == 0)
                {
                    ShowAlert("كل المحافظات موجودة بالفعل في البيانات المعروضة.", AlertForm.AlertType.Info);
                    return;
                }

                // Example: Show them in a MessageBox
                string result = string.Join(Environment.NewLine, missingGovernorates.Select(g => g.governorate));
                MessageBox.Show(result, "المحافظات غير الموجودة في البيانات", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Or: If you have a ListBox to show them, use this:
                // listBoxGovernorates.Items.Clear();
                // foreach (var g in missingGovernorates)
                //     listBoxGovernorates.Items.Add(g.governorate);
            }
            catch (Exception ex)
            {
                ShowAlert("حدث خطأ أثناء تحميل المحافظات: " + ex.Message, AlertForm.AlertType.Error);
            }
        }

        private void advancedDataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            HighlightExpiryCells();
            if (MultyUnitsCHB.Checked && originalData != null)
            {
                // Recalculate Name_Projects repetitions from ORIGINAL data
                var nameGroups = originalData.AsEnumerable()
                    .Where(r => !string.IsNullOrWhiteSpace(r.Field<string>("Name_Projects")))
                    .GroupBy(r => r.Field<string>("Name_Projects"))
                    .ToDictionary(g => g.Key, g => g.Count());

                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.IsNewRow) continue;

                    string projectName = row.Cells["Name_Projects"].Value?.ToString();

                    if (!string.IsNullOrWhiteSpace(projectName) &&
                        nameGroups.ContainsKey(projectName) &&
                        nameGroups[projectName] == 1)
                    {
                        row.DefaultCellStyle.BackColor = Color.Aqua;
                    }
                    else
                    {
                        row.DefaultCellStyle.BackColor = Color.White;
                    }
                }
            }
            else
            {
                // Reset colors when checkbox is unchecked
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                        row.DefaultCellStyle.BackColor = Color.White;
                }
            }
        }

        private void labelBridges_Click(object sender, EventArgs e)
        {
            if (mixedStationsCache == null || mixedStationsCache.Count == 0)
            {
                MessageBox.Show("لا توجد محطات تحتوي على حالات مؤجرة وغير مؤجرة في نفس الوقت",
                                "معلومات", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Create scrollable popup
            Form popup = new Form();
            popup.StartPosition = FormStartPosition.CenterParent;
            popup.Text = "المحطات التي تحتوي على حالات متداخلة";
            popup.Size = new Size(500, 600);
            popup.MinimizeBox = false;
            popup.MaximizeBox = false;
            popup.ShowIcon = false;
            popup.FormBorderStyle = FormBorderStyle.FixedDialog;
            popup.RightToLeft = RightToLeft.Yes;
            popup.RightToLeftLayout = true;

            Label titleLabel = new Label()
            {
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 60,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 11, FontStyle.Bold),
                Text = $"عدد المحطات التي تحتوي على حالات متداخلة: {mixedStationsCache.Count}"
            };

            ListBox listBox = new ListBox()
            {
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Segoe UI", 10),
                HorizontalScrollbar = true
            };

            int index = 1;
            foreach (var name in mixedStationsCache)
            {
                listBox.Items.Add($"{index}. {name}");
                index++;
            }

            Button closeButton = new Button()
            {
                Text = "إغلاق",
                Dock = DockStyle.Bottom,
                Height = 40,
                Font = new System.Drawing.Font("Segoe UI", 10, FontStyle.Bold)
            };
            closeButton.Click += (s, ev) => popup.Close();

            // 🔹 Double-click on station name to filter DataGridView manually
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
                    if (originalData == null)
                    {
                        MessageBox.Show("البيانات الأصلية غير متوفرة لتطبيق الفلتر.",
                                        "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Create a new DataView and filter manually
                    DataView view = new DataView(originalData);
                    view.RowFilter = $"Name_Projects LIKE '%{selectedName.Replace("'", "''")}%'";

                    advancedDataGridView1.DataSource = view;
                    originalData = view.ToTable(); // Keep it consistent with your FilterStringChanged code
                    UpdateRowCount();

                    popup.Close();

                    // Optional feedback message
                    ShowAlert($"تم عرض بيانات المحطة: {selectedName}", AlertForm.AlertType.Info);
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

        private void dreamButton6_Click(object sender, EventArgs e)
        {
            IsWaitedDocs = !IsWaitedDocs;
            if (IsWaitedDocs)
            {
                isCheckedAll = false;
                isCheckedRented = false;
                isCheckedNotRented = false;
                dreamButton2.BackColor = Color.LightGreen;
                dreamButton2.Text = " مؤجر ☐";
                dreamButton1.BackColor = Color.LightGray;
                dreamButton1.Text = " الكل ☐";
                dreamButton3.BackColor = Color.LightGray;
                dreamButton3.Text = "غير مؤجر ☐";
                dreamButton6.BackColor = Color.LightGray;
                dreamButton6.Text = "منتظر العقد ✅";
                // ✅ Apply in-memory filter instead of reloading data
                if (originalData != null)
                {
                    DataView view = new DataView(originalData);
                    view.RowFilter = "[Rental_Status] = 'منتظر العقد' " +
                        "OR [Rental_Status] = 'منتظر التوقيع' " +
                        "OR [Rental_Status] = 'تابع ادارة شل اوت' " +
                        "OR [Rental_Status] = 'جاري التجديد' " +
                        //"OR [Rental_Status] = 'مؤجر (بدون عقد)' " +
                        //"OR [Rental_Status] = 'جاري التوقيع علي العقد' " +
                        "AND [Visible] = true";
                    advancedDataGridView1.DataSource = view;
                }
                else
                {
                    // if unchecked, reset back to all data
                    advancedDataGridView1.DataSource = originalData;
                }
                advancedDataGridView2.DataSource = this.activity_LookupTableAdapter.GetData();
                FalseFunction();
                loadcomboxes();
                UpdateRowCount();
                GenerativePanalFlow();
                ApplyGuna2StyleToGrid(advancedDataGridView1);
                ApplyGuna2StyleToGrid(advancedDataGridView2);
                ArabicColumnGrid();
                LoadColumnsIntoCheckedListBox();
            }
        }
        
        private void dreamButton7_Click(object sender, EventArgs e)
        {
            string target = "خدمات سيارات";

            var names =
                originalData.AsEnumerable()
                .Where(r =>
                {
                    string type = r.Field<string>("Activity_Type")?.Trim();
                    if (string.IsNullOrWhiteSpace(type)) return true; // empty = not خدمات سيارات
                    return !type.Contains(target);  // <-- the FIX
                })
                .Select(r => r.Field<string>("Name_Projects")?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct()
                .ToList();

            MessageBox.Show(string.Join("\n", names));

        }

        private void dreamButton8_Click(object sender, EventArgs e)
        {

        }

        private void skyButton5_Click(object sender, EventArgs e)
        {
            try
            {
                // ------------------------------------------------
                // 1) Validate Data
                // ------------------------------------------------
                if (advancedDataGridView1.Rows.Count == 0)
                {
                    ShowAlert("لا يوجد بيانات", AlertForm.AlertType.Error);
                    return;
                }

                string reportName = "تقرير بحصر محطات شل أوت";

                // ------------------------------------------------
                // 2) Prepare Excel
                // ------------------------------------------------
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string path = Path.Combine(desktop, $"{reportName}.xlsx");

                var excel = new Microsoft.Office.Interop.Excel.Application();
                var wb = excel.Workbooks.Add();
                var ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                ws.Name = "ShellOut Report";

                ws.DisplayRightToLeft = true;


                // ------------------------------------------------
                // 3) Extract and Group Data
                // ------------------------------------------------
                var stationGroups = advancedDataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow)
                    .Select(r => new
                    {
                        Land_Id = r.Cells["Land_Id"].Value?.ToString() ?? "",
                        Name = r.Cells["Name_Projects"].Value?.ToString() ?? "",
                        Type = r.Cells["investment_type"].Value?.ToString() ?? "",
                        Shops = Convert.ToInt32(r.Cells["Shops_Count"].Value ?? 0)
                    })
                    .Where(x => (x.Type == "محلات شل اوت خارج" || x.Type == "محلات شل اوت داخل" )&& !x.Land_Id.Contains("$"))
                    .GroupBy(x => x.Name)
                    .Select(g => new
                    {
                        Station = g.Key,
                        ShopsIn = g.Where(x => x.Type == "محلات شل اوت داخل").Sum(x => x.Shops),
                        ShopsOut = g.Where(x => x.Type == "محلات شل اوت خارج").Sum(x => x.Shops),
                    })
                    .OrderBy(x => x.Station)
                    .ToList();


                // ------------------------------------------------
                // 4) Write Title
                // ------------------------------------------------
                ws.Range["A1", "E1"].Merge();
                ws.Range["A1"].Value = reportName;
                ws.Range["A1"].Font.Bold = true;
                ws.Range["A1"].Font.Size = 28;
                ws.Range["A1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Range["A1"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ws.Range["A1"].RowHeight = 60;


                // ------------------------------------------------
                // 5) Write Headers
                // ------------------------------------------------
                string[] headers = { "مسلسل", "اسم المحطة", "داخل", "خارج", "الإجمالي" };

                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cells[2, i + 1] = headers[i];
                }

                var headerRange = ws.Range["A2", "E2"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                // ------------------------------------------------
                // 6) Write Data
                // ------------------------------------------------
                int row = 3;
                int serial = 1;

                foreach (var s in stationGroups)
                {
                    ws.Cells[row, 1] = serial++;
                    ws.Cells[row, 2] = s.Station;
                    ws.Cells[row, 3] = s.ShopsIn;
                    ws.Cells[row, 4] = s.ShopsOut;
                    ws.Cells[row, 5] = s.ShopsIn + s.ShopsOut;

                    row++;
                }


                // ------------------------------------------------
                // 7) Format Table
                // ------------------------------------------------
                var fullRange = ws.Range["A1", $"E{row - 1}"];
                fullRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                fullRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws.Columns.AutoFit();
                ws.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                wb.SaveAs(path);
                excel.Visible = true;

                ShowAlert("تم إنشاء التقرير بنجاح", AlertForm.AlertType.Success);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطأ أثناء إنشاء التقرير:\n{ex.Message}");
            }
            finally
            {
                FalseFunction();
            }
        }

        private void skyButton6_Click(object sender, EventArgs e)
        {
            try
            {
                // ------------------------------------------------
                // 1) Validate Data
                // ------------------------------------------------
                if (advancedDataGridView1.Rows.Count == 0)
                {
                    ShowAlert("لا يوجد بيانات", AlertForm.AlertType.Error);
                    return;
                }

                string reportName = "تقرير بحصر كباري مسجله";

                // ------------------------------------------------
                // 2) Prepare Excel
                // ------------------------------------------------
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string path = Path.Combine(desktop, $"{reportName}.xlsx");

                var excel = new Microsoft.Office.Interop.Excel.Application();
                var wb = excel.Workbooks.Add();
                var ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.ActiveSheet;
                ws.Name = "bridges Report";

                ws.DisplayRightToLeft = true;


                // ------------------------------------------------
                // 3) Extract and Group Data
                // ------------------------------------------------
                var stationGroups = advancedDataGridView1.Rows
                    .Cast<DataGridViewRow>()
                    .Where(r => !r.IsNewRow)
                    .Select(r => new
                    {
                        Land_Id = r.Cells["Land_Id"].Value?.ToString() ?? "",
                        Name = r.Cells["Name_Projects"].Value?.ToString() ?? "",
                        Type = r.Cells["investment_type"].Value?.ToString() ?? "",
                        Rental = r.Cells["Rental_Status"].Value?.ToString() ?? "",
                        Activity_Name = r.Cells["Activity_Name"].Value?.ToString() ?? "",
                        investment_name = r.Cells["investment_name"].Value?.ToString() ?? "",
                        Shops = Convert.ToInt32(r.Cells["Shops_Count"].Value ?? 0)
                    })
                    .Where(x => (x.Type == "اسفل كباري"))
                    .GroupBy(x => x.Name)
                    .Select(g => new
                    {
                        Station = g.Key,
                        Rented = g.Where(x => x.Activity_Name !="لا يوجد").Sum(x => x.Shops),
                        NotRented = g.Where(x => (x.Rental.Contains("غير")) ||
                        (x.Activity_Name == "لا يوجد" && string.IsNullOrEmpty(x.investment_name)))
                        .Sum(x => x.Shops),
                        Waited = g.Where(x => (x.Rental == null || x.Rental.Contains("منتظر العقد")) &&
                        x.Activity_Name == "لا يوجد" &&
                        !string.IsNullOrEmpty(x.investment_name))
                        .Sum(x => x.Shops),
                    })  
                    .OrderBy(x => x.Station)
                    .ToList();


                // ------------------------------------------------
                // 4) Write Title
                // ------------------------------------------------
                ws.Range["A1", "E1"].Merge();
                ws.Range["A1"].Value = reportName;
                ws.Range["A1"].Font.Bold = true;
                ws.Range["A1"].Font.Size = 28;
                ws.Range["A1"].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Range["A1"].VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                ws.Range["A1"].RowHeight = 60;


                // ------------------------------------------------
                // 5) Write Headers
                // ------------------------------------------------
                string[] headers = { "مسلسل", "اسم المحطة", "مؤجر", "غير مؤجر","منتظر العقد", "الإجمالي" };

                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cells[2, i + 1] = headers[i];
                }

                var headerRange = ws.Range["A2", "E2"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;


                // ------------------------------------------------
                // 6) Write Data
                // ------------------------------------------------
                int row = 3;
                int serial = 1;

                foreach (var s in stationGroups)
                {
                    ws.Cells[row, 1] = serial++;
                    ws.Cells[row, 2] = s.Station;
                    ws.Cells[row, 3] = s.Rented;
                    ws.Cells[row, 4] = s.NotRented;
                    ws.Cells[row, 5] = s.Waited;
                    ws.Cells[row, 6] = s.Rented + s.NotRented + s.Waited;

                    row++;
                }


                // ------------------------------------------------
                // 7) Format Table
                // ------------------------------------------------
                var fullRange = ws.Range["A1", $"E{row - 1}"];
                fullRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                fullRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                ws.Columns.AutoFit();
                ws.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                ws.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

                wb.SaveAs(path);
                excel.Visible = true;

                ShowAlert("تم إنشاء التقرير بنجاح", AlertForm.AlertType.Success);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطأ أثناء إنشاء التقرير:\n{ex.Message}");
            }
            finally
            {
                FalseFunction();
            }
        }

        private void MultyUnitsCHB_CheckedChanged(object sender, EventArgs e)
        {
            ApplyUnitFilter();
        }

        private void OneUnitCHB_CheckedChanged(object sender, EventArgs e)
        {
            ApplyUnitFilter();
        }
        private void HandleGridAction(int rowIndex, int columnIndex)
        {
            try
            {
                if (rowIndex < 0 || columnIndex < 0) return;

                string columnName = advancedDataGridView1.Columns[columnIndex].Name;

                // ========= Activity_Name =========
                if (columnName == "Activity_Name")
                {
                    string activityName = advancedDataGridView1.Rows[rowIndex]
                        .Cells["Activity_Name"].Value?.ToString() + ".PDF";

                    string projectId = advancedDataGridView1.Rows[rowIndex]
                        .Cells["Project_Id"].Value?.ToString();

                    if (string.IsNullOrWhiteSpace(activityName) || string.IsNullOrWhiteSpace(projectId))
                        return;

                    var docs = documentsTableAdapter.GetDataByProjectOnly(projectId);

                    var match = docs.FirstOrDefault(d =>
                        !d.IspathsNull() &&
                        Path.GetFileName(d.paths)
                            .Equals(activityName, StringComparison.OrdinalIgnoreCase));

                    if (match == null)
                    {
                        MessageBox.Show("لا يوجد ملف مطابق لهذا النشاط والمشروع");
                        return;
                    }

                    OpenFileWithFallback(match.paths);
                }

                // ========= Description_Drawing_Place =========
                else if (columnName == "Description_Drawing_Place")
                {
                    string fileName = advancedDataGridView1.Rows[rowIndex]
                        .Cells["Description_Drawing_Place"].Value?.ToString();

                    string projectId = advancedDataGridView1.Rows[rowIndex]
                        .Cells["Project_Id"].Value?.ToString();

                    if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(projectId))
                        return;

                    var docs = documentsTableAdapter.GetDataByProjectOnly(projectId);

                    var match = docs.FirstOrDefault(d =>
                        !d.IspathsNull() &&
                        Path.GetFileName(d.paths)
                            .Equals(fileName, StringComparison.OrdinalIgnoreCase));

                    if (match == null)
                    {
                        MessageBox.Show("لا يوجد ملف مطابق لهذا النشاط والمشروع");
                        return;
                    }

                    OpenFileWithFallback(match.paths);
                }

                // ========= Update mode =========
                if (Update_Radio.Checked)
                {
                    var serialNumber = advancedDataGridView1.Rows[rowIndex]
                        .Cells["investments_id"].Value?.ToString();

                    if (!string.IsNullOrEmpty(serialNumber))
                    {
                        serial_number_TB.Text = serialNumber;
                        LoadInvestmentDataForUpdatePart(serialNumber);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void OpenFileWithFallback(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    Process.Start(path);
                    Form parentForm = this.FindForm();
                    if (parentForm != null)
                    {
                        parentForm.WindowState = FormWindowState.Minimized;
                    }
                    return;
                }

                if (path.StartsWith(@"Z:\", StringComparison.OrdinalIgnoreCase))
                {
                    string newPath = Path.Combine(@"D:\sho8l", path.Substring(3));

                    if (File.Exists(newPath))
                    {
                        Process.Start(newPath);
                        Form parentForm = this.FindForm();
                        if (parentForm != null)
                        {
                            parentForm.WindowState = FormWindowState.Minimized;
                        }
                        return;
                    }
                }

                MessageBox.Show("الملف غير موجود في المسار");
            }
            catch (Exception ex)
            {

            }
        }
        private void advancedDataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter && advancedDataGridView1.CurrentCell != null)
                {
                    HandleGridAction(
               advancedDataGridView1.CurrentCell.RowIndex,
               advancedDataGridView1.CurrentCell.ColumnIndex
                    );
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void guna2CircleButton3_Click(object sender, EventArgs e)
        {
            try
            {
                if (advancedDataGridView1.DataSource == null)
                    return;

                var bs = (BindingSource)advancedDataGridView1.DataSource;
                var table = (DataTable)bs.DataSource;
                string newId="";

                if (SessionData.Investment_Type == "كل الاستثمارات")
                {
                    
                }
                if (SessionData.Investment_Type == "استثمارات على الطرق")
                {
                     newId = GetNextInvestmentIdByCategory(table, "&");
                }
                if (SessionData.Investment_Type == "محلات شل اوت داخل")
                {
                     newId = GetNextInvestmentIdByCategory(table, "!");
                }
                if (SessionData.Investment_Type == "مولات")
                {
                     newId = GetNextInvestmentIdByCategory(table, "MA");
                }
                if (SessionData.Investment_Type == "محلات شل اوت خارج")
                {
                     newId = GetNextInvestmentIdByCategory(table, "");
                }
                if (SessionData.Investment_Type == "اسفل كباري")
                {
                     newId = GetNextInvestmentIdByCategory(table, "#");
                }
                // 2) إنشاء صف جديد
                DataRow newRow = table.NewRow();

                foreach (DataColumn col in table.Columns)
                {
                    if (col.ColumnName == "investments_id")
                        newRow[col.ColumnName] = newId.ToString();
                    else if (col.DataType == typeof(bool))
                        newRow[col.ColumnName] = false;
                    else if (col.DataType == typeof(DateTime))
                        newRow[col.ColumnName] = DBNull.Value;
                    else if (col.DataType == typeof(int) || col.DataType == typeof(decimal))
                        newRow[col.ColumnName] = 0;
                    else
                        newRow[col.ColumnName] = ""; // نصوص فاضية
                }

                // 3) إضافة الصف للجدول
                table.Rows.Add(newRow);

                // 4) تحديد الصف الجديد
                bs.Position = bs.Count - 1;

                // 5) حفظ في قاعدة البيانات (اختياري الآن أو عند Save)
                investmentsTableAdapter.Insert(
                    newId.ToString(),   // investments_id
                    null,               // land_fk
                    null,               // governorate_fk
                    "0",               // investment_type
                    "0",               // investment_name
                    null,               // Activity_Type
                    null,             // Location
                    "لا يوجد",               // Activity_Name
                    null,               // Contract_start_date
                    null,              // Contract_expiry_date
                    null,                  // Rental_value
                    null,                  // Shops_Count
                    null,               // Visible
                    null,               // Place_number
                    null,               // Contract_number
                    null,               // Offer_memorandum_number
                    null,               // Notes
                    SessionData.Investment_Type, null,null,true,null,null,null                // Contract_terms
                );

                ShowAlert("تم إضافة صف جديد بنجاح", AlertForm.AlertType.Success);
            }
            catch (Exception ex)
            {
                ShowAlert("خطأ أثناء الإضافة: " + ex.Message, AlertForm.AlertType.Error);
            }
        }
        private string GetNextInvestmentIdByCategory(DataTable table, string prefix)
        {
            int maxNumber = 0;

            foreach (DataRow row in table.Rows)
            {
                string id = row["investments_id"]?.ToString();
                if (string.IsNullOrWhiteSpace(id))
                    continue;

                // لازم يبدأ بنفس الـ prefix
                if (!id.StartsWith(prefix))
                    continue;

                // استخراج الجزء الرقمي
                string numericPart = id.Substring(prefix.Length);

                if (int.TryParse(numericPart, out int num))
                {
                    if (num > maxNumber)
                        maxNumber = num;
                }
            }

            return $"{prefix}{maxNumber + 1}";
        }


    }
}
