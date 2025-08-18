using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Media;

namespace DemoProject
{
    public partial class SettingControl : UserControl
    {
        public long UserId;
        public void SetUserData(string user, long id)
        {
            this.UserId = id;
        }
        public SettingControl()
        {
            InitializeComponent();
        }
        public static DataTable ToDataTable<T>(IEnumerable<T> data)
        {
            var props = typeof(T).GetProperties();
            var table = new DataTable();
            foreach (var prop in props)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }
            foreach (var item in data)
            {
                var row = table.NewRow();
                foreach (var prop in props)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }
        public void LoadGovernorates()
        {
            var governoratesData = this.governorateTableAdapter.GetData();
            advancedDataGridView1.DataSource = governoratesData;
        }
        public void LoadApprovals()
        {
            var approvalsData = this.approvalsTableAdapter.GetData();
            advancedDataGridView2.DataSource = approvalsData;
        }
        public void LoadUsers()
        {
            var usersData = this.usersTableAdapter.GetData();
            advancedDataGridView3.DataSource = usersData;
            foreach (DataRow row in usersData.Rows)
            {
                string name = row["user_name"].ToString(); // use column name exactly as in your table
                User_Name_COB.Items.Add(name);
            }
        }
        public void LoadRoles()
        {
            var usersData = this.usersTableAdapter.GetData();
            var rolesData = this.rolesTableAdapter.GetData();
            // Perform LINQ join to get role data along with the user's real name
            var resultusername = from role in rolesData
                                 join user in usersData on role.user_id equals user.id into userGroup
                                 from user in userGroup.DefaultIfEmpty()
                                 select new
                                 {
                                     role.id,
                                     role.role_name,
                                     user_name = user != null ? user.user_name : "No User"
                                 };
            advancedDataGridView4.DataSource = ToDataTable(resultusername);          
            foreach (DataRow row in rolesData.Rows)
            {
                string name = row["role_name"].ToString(); // use column name exactly as in your table
                Roles_Names_COB.Items.Add(name);
                Role_COB.Items.Add(name);
            }
        }
        public void LoadPages()
        {
            var rolesData = this.rolesTableAdapter.GetData();
            var pagesData = this.pagesTableAdapter.GetData();
            var resultrolenamePage = from page in pagesData
                                     join role in rolesData on page.role_id equals role.id into roleGroup
                                     from role in roleGroup.DefaultIfEmpty()
                                     select new
                                     {
                                         page.id,
                                         PageName = page.Page_name,
                                         RoleName = role != null ? role.role_name : "No Role"
                                     };
            advancedDataGridView5.DataSource = ToDataTable(resultrolenamePage);
        }
        public void LoadFunctions()
        {
            var rolesData = this.rolesTableAdapter.GetData();
            var functionsData = this.functionsTableAdapter.GetData();
            var resultrolenameFunction = from func in functionsData
                                         join role in rolesData on func.role_id equals role.id into roleGroup
                                         from role in roleGroup.DefaultIfEmpty()
                                         select new
                                         {
                                             func.id,
                                             FunctionName = func.Function_name,
                                             RoleName = role != null ? role.role_name : "No Role"
                                         };
            advancedDataGridView6.DataSource = ToDataTable(resultrolenameFunction);
        }
        public void LoadAccess()
        {
            var usersData = this.usersTableAdapter.GetData();
            var functionsData = this.functionsTableAdapter.GetData();
            var pagesData = this.pagesTableAdapter.GetData();
            var rolesData = this.rolesTableAdapter.GetData();
            var accessData = this.accessTableAdapter.GetData();
            var resultAccess = from access in accessData
                               join role in rolesData on access.role_id equals role.id into roleGroup
                               from role in roleGroup.DefaultIfEmpty()
                               join page in pagesData on access.pages_id equals page.id into pageGroup
                               from page in pageGroup.DefaultIfEmpty()
                               join func in functionsData on access.function_id equals func.id into funcGroup
                               from func in funcGroup.DefaultIfEmpty()
                               join user in usersData on access.user_id equals user.id into userGroup
                               from user in userGroup.DefaultIfEmpty()
                               select new
                               {
                                   access_id = access.access_id, 
                                   Role = role != null ? role.role_name : "No Role",
                                   Page = page != null ? page.Page_name : "No Page",
                                   Function = func != null ? func.Function_name : "No Function",
                                   User = user != null ? user.user_name : "No User"     
                               };
            advancedDataGridView7.DataSource = ToDataTable(resultAccess);
            foreach(DataRow row in usersData.Rows)
            {
                string name = row["user_name"].ToString();
                Access_User_COB.Items.Add(name);
            }
            foreach (DataRow row in pagesData.Rows)
            {
                string name = row["Page_name"].ToString();
                Access_Page_COB.Items.Add(name);
            }
            foreach (DataRow row in functionsData.Rows)
            {
                string name = row["Function_name"].ToString();
                Access_Func_COB.Items.Add(name);
            }
        }   
        private void SettingControl_Load(object sender, EventArgs e)
        {
            LoadGovernorates();
            LoadApprovals();
            LoadUsers();
            LoadRoles();
            LoadPages();
            LoadFunctions();
            LoadAccess();
            ApplyGuna2StyleToGrid(advancedDataGridView7);
            ApplyGuna2StyleToGrid(advancedDataGridView6);
            ApplyGuna2StyleToGrid(advancedDataGridView5);
            ApplyGuna2StyleToGrid(advancedDataGridView4);
            ApplyGuna2StyleToGrid(advancedDataGridView3);
            ApplyGuna2StyleToGrid(advancedDataGridView2);
            ApplyGuna2StyleToGrid(advancedDataGridView1);
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
        private void guna2Button2_Click(object sender, EventArgs e)
        {
            var lines = guna2TextBox1.Lines
            .Where(line => !string.IsNullOrWhiteSpace(line)) // skip empty lines
            .Select(line => line.Trim())
            .ToList();
            if (lines.Count == 0)
            {
                MessageBox.Show("Please enter at least one governorate.");
                return;
            }
            // Optional: Get last ID from DB and increment from there
            int nextId = GetNextGovernorateIdFromDB(); // or start from 1
            foreach (var name in lines)
            {
                string id = nextId.ToString();
                var ExistedId =this.governorateTableAdapter.GetDataByGovernorate(name);
                if (ExistedId.First().governorate.Any())
                {
                    ShowAlert("غير مسموح بي تكرار نفس المحافظة",AlertForm.AlertType.Error);
                    return;
                }
                this.governorateTableAdapter.Insert(id, name); // adjust to your adapter
                nextId++;
            }
            MessageBox.Show("Governorates inserted successfully!");
            this.advancedDataGridView1.DataSource = this.governorateTableAdapter.GetData();
            guna2TextBox1.Clear();
        }
        private int GetNextGovernorateIdFromDB()
        {
            var table = governorateTableAdapter.GetData();
            if (table.Rows.Count == 0) return 1;
            var ids = table.AsEnumerable()
                           .Select(row => int.Parse(row["governorate_id"].ToString()))
                           .ToList();
            return ids.Max() + 1;
        }
        private int GetNextApprovalIdFromDB()
        {
            var table = approvalsTableAdapter.GetData();
            if (table.Rows.Count == 0) return 1;
            var ids = table.AsEnumerable()
                           .Select(row => int.Parse(row["approval_id"].ToString()))
                           .ToList();
            return ids.Max() + 1;
        }
        private void guna2Button3_Click(object sender, EventArgs e)
        {
            var lines = guna2TextBox3.Lines
            .Where(line => !string.IsNullOrWhiteSpace(line)) // skip empty lines
            .Select(line => line.Trim())
            .ToList();
            if (lines.Count == 0)
            {
                MessageBox.Show("Please enter at least one approval.");
                return;
            }
            // Optional: Get last ID from DB and increment from there
            int nextId = GetNextApprovalIdFromDB(); // or start from 1
            foreach (var name in lines)
            {
                string id = nextId.ToString();
                this.approvalsTableAdapter.Insert(id, name); // adjust to your adapter
                nextId++;
            }
            ShowAlert("Approvals inserted successfully!",AlertForm.AlertType.Success);
            this.advancedDataGridView2.DataSource = this.approvalsTableAdapter.GetData();
            guna2TextBox3.Clear();
        }
        private void guna2Button4_Click(object sender, EventArgs e)
        {        
        }
        private void guna2Button1_Click(object sender, EventArgs e)
        {    
        }
        private void Land_Name_TB_TextChanged(object sender, EventArgs e)
        {
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
        private void guna2Button4_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(User_Name_TB.Text) || string.IsNullOrWhiteSpace(Password_TB.Text))
            {
                ShowAlert("يجب عليك ادخال بيانات المستخدم",AlertForm.AlertType.Error);
                if (string.IsNullOrWhiteSpace(User_Name_TB.Text)) { }
                if (string.IsNullOrWhiteSpace(Password_TB.Text)) { }
                return;
            }
           int user_Id =int.Parse(this.usersTableAdapter.GetData().Last().id.ToString());
           this.usersTableAdapter.Insert(++user_Id,User_Name_TB.Text,Password_TB.Text,Real_Name_TB.Text);
            User_Name_COB.Items.Clear();
            Roles_Names_COB.Items.Clear();
            Role_COB.Items.Clear();
            Access_User_COB.Items.Clear();
            Access_role_COB.Items.Clear();
            Access_Page_COB.Items.Clear();
            Access_Func_COB.Items.Clear();
            LoadUsers();
            LoadRoles();
            LoadPages();
            LoadFunctions();
            LoadAccess();
            ShowAlert("تمت اضافة المستخدم بنجاح",AlertForm.AlertType.Success);
        }    
        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(User_Name_COB.Text)|| string.IsNullOrWhiteSpace(Role_TB.Text)) { 
                ShowAlert("يجب عليك ادخال بيانات مسئولية المستخدم", AlertForm.AlertType.Error);
                if (string.IsNullOrWhiteSpace(User_Name_COB.Text)) { }
                if (string.IsNullOrWhiteSpace(Role_TB.Text)) { }
                return;
            }
            int userIDValue = int.Parse(this.usersTableAdapter.GetDataByUserName(User_Name_COB.Text).First().id.ToString());
            int role_Id = int.Parse(this.rolesTableAdapter.GetData().Last().id.ToString());
            this.rolesTableAdapter.Insert(++role_Id, Role_TB.Text, userIDValue);
            User_Name_COB.Items.Clear();
            Roles_Names_COB.Items.Clear();
            Role_COB.Items.Clear();
            Access_User_COB.Items.Clear();
            Access_role_COB.Items.Clear();
            Access_Page_COB.Items.Clear();
            Access_Func_COB.Items.Clear();
            LoadUsers();
            LoadRoles();
            LoadPages();
            LoadFunctions();
            LoadAccess();
            LoadcomBox_AccessRole();
            ShowAlert("تمت اضافة المسئولية علي المستخدم بنجاح", AlertForm.AlertType.Success);
        }
        private void guna2Button5_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Page_TB.Text) || string.IsNullOrWhiteSpace(Roles_Names_COB.Text))
            {
                ShowAlert("يجب عليك ادخال بيانات مسئولية الصفحة ", AlertForm.AlertType.Error);
                if (string.IsNullOrWhiteSpace(Page_TB.Text)) { }
                if (string.IsNullOrWhiteSpace(Roles_Names_COB.Text)) { }
                return;
            }
            int RoleIDValue = int.Parse(this.rolesTableAdapter.GetDataByUserRole(Roles_Names_COB.Text).First().id.ToString());
            int Page_Id = int.Parse(this.pagesTableAdapter.GetData().Last().id.ToString());
            this.pagesTableAdapter.Insert(++Page_Id, Page_TB.Text, RoleIDValue);
            User_Name_COB.Items.Clear();
            Roles_Names_COB.Items.Clear();
            Role_COB.Items.Clear();
            Access_User_COB.Items.Clear();
            Access_role_COB.Items.Clear();
            Access_Page_COB.Items.Clear();
            Access_Func_COB.Items.Clear();
            LoadUsers();
            LoadRoles();
            LoadPages();
            LoadFunctions();
            LoadAccess();
            ShowAlert("تمت اضافة مسئولية الصفحة علي المستخدم بنجاح", AlertForm.AlertType.Success);
        }
        private void guna2Button6_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Function_TB.Text) || string.IsNullOrWhiteSpace(Role_COB.Text))
            {
                ShowAlert("يجب عليك ادخال بيانات مسئولية الوظيفة ", AlertForm.AlertType.Error);
                if (string.IsNullOrWhiteSpace(Function_TB.Text)) { }
                if (string.IsNullOrWhiteSpace(Role_COB.Text)) { }
                return;
            }
            int RoleIDValue = int.Parse(this.rolesTableAdapter.GetDataByUserRole(Role_COB.Text).First().id.ToString());
            int Func_Id = int.Parse(this.functionsTableAdapter.GetData().Last().id.ToString());
            this.functionsTableAdapter.Insert(++Func_Id, Function_TB.Text, RoleIDValue);
            User_Name_COB.Items.Clear();
            Roles_Names_COB.Items.Clear();
            Role_COB.Items.Clear();
            Access_User_COB.Items.Clear();
            Access_role_COB.Items.Clear();
            Access_Page_COB.Items.Clear();
            Access_Func_COB.Items.Clear();
            LoadUsers();
            LoadRoles();
            LoadPages();
            LoadFunctions();
            LoadAccess();
            ShowAlert("تمت اضافة مسئولية الوظيفة علي المستخدم بنجاح", AlertForm.AlertType.Success);
        }
        private void guna2Button7_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Access_User_COB.Text) ||
                string.IsNullOrWhiteSpace(Access_role_COB.Text)||
                string.IsNullOrWhiteSpace(Access_Page_COB.Text) ||
                string.IsNullOrWhiteSpace(Access_Func_COB.Text) 
                )
            {
                ShowAlert("يجب عليك ادخال بيانات الصلاحيات الخاصه بالمستخدم في جميع الحقول ", AlertForm.AlertType.Error);
                if (string.IsNullOrWhiteSpace(Access_User_COB.Text)) { }
                if (string.IsNullOrWhiteSpace(Access_role_COB.Text)) { }
                if (string.IsNullOrWhiteSpace(Access_Page_COB.Text)) { }
                if (string.IsNullOrWhiteSpace(Access_Func_COB.Text)) { }
                return;
            }
            int userIDValue = int.Parse(this.usersTableAdapter.GetDataByUserName(Access_User_COB.Text).First().id.ToString());
            int RoleIDValue = int.Parse(this.rolesTableAdapter.GetDataByUserRole(Access_role_COB.Text).First().id.ToString());
            int PageIDValue = int.Parse(this.pagesTableAdapter.GetDataByUserPage(Access_Page_COB.Text).First().id.ToString());
            int FuncIDValue = int.Parse(this.functionsTableAdapter.GetDataByUserFunc(Access_Func_COB.Text).First().id.ToString());
            int Access_Id = int.Parse(this.accessTableAdapter.GetData().Last().access_id.ToString());
            this.accessTableAdapter.Insert(++Access_Id, RoleIDValue, PageIDValue, FuncIDValue, userIDValue);
            User_Name_COB.Items.Clear();
            Roles_Names_COB.Items.Clear();
            Role_COB.Items.Clear();
            Access_User_COB.Items.Clear();
            Access_role_COB.Items.Clear();
            Access_Page_COB.Items.Clear();
            Access_Func_COB.Items.Clear();
            LoadUsers();
            LoadRoles();
            LoadPages();
            LoadFunctions();
            LoadAccess();
            ShowAlert("تمت اضافة الصلاحيه للمستخدم بنجاح", AlertForm.AlertType.Success);
        }
        public void LoadcomBox_AccessRole()
        {
            this.BeginInvoke(new Action(() =>
            {
                var usersData = this.usersTableAdapter.GetData();
                var rolesData = this.rolesTableAdapter.GetData();
                var accessData = this.accessTableAdapter.GetData();
                string selectedUserName = Access_User_COB.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selectedUserName)) return;
                var selectedUser = usersData.AsEnumerable()
                 .FirstOrDefault(r => r["user_name"].ToString() == selectedUserName);
                if (selectedUser == null) return;
                string userId = selectedUser["id"].ToString();
                var userRoles = from role in rolesData.AsEnumerable()
                                where role["user_id"].ToString() == userId
                                select role["role_name"].ToString();
                var distinctRoles = userRoles.Distinct().ToList();
                // Force ComboBox to treat list as simple strings
                Access_role_COB.DataSource = null;
                Access_role_COB.Items.Clear();
                foreach (var roleName in distinctRoles)
                {
                    Access_role_COB.Items.Add(roleName);
                }
                Access_role_COB.SelectedIndex = -1;
            }));
        }
        private void Access_User_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.BeginInvoke(new Action(() =>
            {
                var usersData = this.usersTableAdapter.GetData();
                var rolesData = this.rolesTableAdapter.GetData();
                var accessData = this.accessTableAdapter.GetData();
                string selectedUserName = Access_User_COB.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(selectedUserName)) return;
                var selectedUser = usersData.AsEnumerable()
                .FirstOrDefault(r => r["user_name"].ToString() == selectedUserName);
                if (selectedUser == null) return;
                string userId = selectedUser["id"].ToString();
                var userRoles = from role in rolesData.AsEnumerable()
                                where role["user_id"].ToString() == userId
                                select role["role_name"].ToString();
                var distinctRoles = userRoles.Distinct().ToList();
                // Force ComboBox to treat list as simple strings
                Access_role_COB.DataSource = null;
                Access_role_COB.Items.Clear();
                foreach (var roleName in distinctRoles)
                {
                    Access_role_COB.Items.Add(roleName);
                }
                Access_role_COB.SelectedIndex = -1;
            }));
        }
        private void Roles_Names_COB_SelectedIndexChanged(object sender, EventArgs e)
        {
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
                parentForm.Close();
            }
        }
    }
}
