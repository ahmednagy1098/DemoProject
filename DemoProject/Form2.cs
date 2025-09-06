using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Timer = System.Windows.Forms.Timer;

namespace DemoProject
{
    public partial class Form2 : Form
    {
        private string _username;
        private long _user_id;

        public void ShowMyUserControl(string user, long id)
        {
            ShowLoading(true);
            engineeringManagementControl1.SetUserData(user, id); // Pass data to the control
            projectsControl1.SetUserData(user, id);
            engineeringManagementControl1.BringToFront();
            settingControl1.SendToBack();
            projectsControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            roadsControl1.SendToBack();
            


            ShowLoading(false);
            engineeringManagementControl1.Visible = true;
            settingControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;

            focusENG.Visible = true;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;

        }
        public void ShowMyUserControl2(string user, long id)
        {
            settingControl1.SetUserData(user, id);
            settingControl1.BringToFront();
            projectsControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            roadsControl1.SendToBack();

            settingControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;

            focusSetting.Visible = true;
            foucsRoad.Visible = false;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;
        }
        public void ShowMyUserControl5(string user, long id)//admin
        {
            ShowLoading(true);
            systemAdministratorControl1.SetUserData(user, id);
            
            
          
            systemAdministratorControl1.Visible = true;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;

            systemAdministratorControl1.BringToFront();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            projectsControl1.SendToBack();

            focusAdmin.Visible = true;
            focusENG.Visible = false;
            focusInvest.Visible = false;           
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            ShowLoading(false);
        }
        public void ShowMyUserControl3(string user, long id)
        {
            investmentsControl1.SetUserData(user, id);
            projectsControl1.Visible = false;
            investmentsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;

            projectsControl1.SendToBack();
            investmentsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();

            focusInvest.Visible = true;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }
        public void ShowMyUserControl4(string user, long id)
        {
            roadsControl1.SetUserData(user, id);
            roadsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;    
            projectsControl1.Visible = false;

            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.BringToFront();
            projectsControl1.SendToBack();

            foucsRoad.Visible = true;
            focusInvest.Visible = false;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;   
        }
        public void ShowMyUserControl6(string user, long id)
        {
            projectsControl1.SetUserData(user, id);
            projectsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;

            projectsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();

            focusInvest.Visible = false;
            focusENG.Visible = true;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }
        public Form2(string username, long user_id)
        {
            InitializeComponent();
            _username = username;
            _user_id = user_id;
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
        Dictionary<string, bool> pageAccess = new Dictionary<string, bool>();
        Dictionary<string, bool> functionAccess = new Dictionary<string, bool>();
        private void Form2_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dATABASE2DataSet.pages' table. You can move, or remove it, as needed.
            this.pagesTableAdapter.Fill(this.dATABASE2DataSet.pages);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.functions' table. You can move, or remove it, as needed.
            this.functionsTableAdapter.Fill(this.dATABASE2DataSet.functions);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.dATABASE2DataSet.roles);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.dATABASE2DataSet.users);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);



            //Investment
            //Roads
            //Settings
            //Admin
            //Eng
            
            button1.Tag = "Page:Admin";
            button2.Tag = "Page:Roads";
            Menu.Tag = "Page:Eng";
            button4.Tag = "Page:Investment";
            button5.Tag = "Page:Settings";
            var userRow = dATABASE2DataSet.users.FirstOrDefault(u => u.id == _user_id);
            var role = dATABASE2DataSet.roles.FirstOrDefault(r => r.user_id == _user_id);
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
        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void menutranstion_Tick(object sender, EventArgs e)
        {
            if (MenuExpand == false)
            {
                MenuContainer.Height += 10;
                if (MenuContainer.Height >= 108) {
                    menutranstion.Stop();
                    MenuExpand = true;
                }
            }
            else
            {
                MenuContainer.Height -= 10;
                if (MenuContainer.Height <= 35)
                {
                    menutranstion.Stop();
                    MenuExpand = false;
                }
            }
        }

        private void Menu_Click(object sender, EventArgs e)
        {
            TrueFunction();
            ShowMyUserControl(_username, _user_id);
            menutranstion.Start();
            FalseFunction();
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void advancedDataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            
        }

        private void Survying_TB_Load(object sender, EventArgs e)
        {
          
        }
        private void UpdateColumnCounters()
        {
            
        }

        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            //
        }

        private void advancedDataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
         
        }

        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
           
            
        }

        private void advancedDataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TrueFunction();
            ShowMyUserControl4(_username, _user_id);
            /*Form3 report = new Form3();
            report.Show();*/
            FalseFunction();
        }

        private void Survying_TB_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2CheckBox4_CheckStateChanged(object sender, EventArgs e)
        {
          
        }

        private void PdfStudy_CB_CheckStateChanged(object sender, EventArgs e)
        {
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 Login = new Form1();
            Login.Show();
            this.Close();
        }

        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {

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
        private void button1_Click(object sender, EventArgs e)
        {
            ShowMyUserControl5(_username, _user_id);  
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ShowMyUserControl3(_username, _user_id);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            TrueFunction();
            ShowMyUserControl(_username, _user_id);
            menutranstion.Start();
            FalseFunction();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            TrueFunction();
            ShowMyUserControl2(_username, _user_id);
            FalseFunction();
        }

        private void button7_Click(object sender, EventArgs e)
        {

            ShowMyUserControl6(_username, _user_id);
           
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (guna2CircleProgressBar1.Value < 100)
            {
                guna2CircleProgressBar1.Value += 2; // speed of loading
            }
            else
            {
                guna2CircleProgressBar1.Value = 0; // reset for loop effect
            }
        }

        public void ShowLoading(bool show)
        {
            guna2CircleProgressBar1.Visible = show;
            guna2CircleProgressBar1.Value = 0;
            guna2CircleProgressBar1.BringToFront();

            if (show)

                timer1.Start();
            else
                timer1.Stop();
        }

        private void accessBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.accessBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dATABASE2DataSet);

        }
    }
}

