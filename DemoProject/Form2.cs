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
using System.Runtime.InteropServices;
namespace DemoProject
{
    public partial class Form2 : Form
    {
        private string _username;
        private long _user_id;
        private string _investment_type;

        private Size originalFormSize;
        private Dictionary<Control, Rectangle> originalControls = new Dictionary<Control, Rectangle>();

        //private void SaveOriginalSizes(Control parent)
        //{
        //    foreach (Control ctrl in parent.Controls)
        //    {
        //        originalControls[ctrl] = new Rectangle(ctrl.Location, ctrl.Size);

        //        if (ctrl.Controls.Count > 0)
        //            SaveOriginalSizes(ctrl);
        //    }
        //}
        //private void ResizeControls(Control parent, float xRatio, float yRatio)
        //{
        //    foreach (Control ctrl in parent.Controls)
        //    {
        //        if (!originalControls.ContainsKey(ctrl)) continue;

        //        Rectangle original = originalControls[ctrl];

        //        ctrl.Location = new Point(
        //            (int)(original.X * xRatio),
        //            (int)(original.Y * yRatio));

        //        ctrl.Size = new Size(
        //            (int)(original.Width * xRatio),
        //            (int)(original.Height * yRatio));

        //        ctrl.Font = new Font(ctrl.Font.FontFamily,
        //            ctrl.Font.Size * Math.Min(xRatio, yRatio));

        //        if (ctrl.Controls.Count > 0)
        //            ResizeControls(ctrl, xRatio, yRatio);
        //    }
        //}
        public void ShowMyUserControl(string user, long id, string Investment_type)
        {
            SessionData.Investment_Type = null;
            ShowLoading(true);
            engineeringManagementControl1.SetUserData(user, id); // Pass data to the control
            investmentsControl1.SetUserData(user, id, Investment_type);
            projectsControl1.SetUserData(user, id);
            engineeringManagementControl1.BringToFront();
            settingControl1.SendToBack();
            projectsControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            roadsControl1.SendToBack();
            this.dashBoardControl1.SendToBack();


            ShowLoading(false);
            engineeringManagementControl1.Visible = true;
            settingControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;
            dashBoardControl1.Visible = false;
            focusENG.Visible = true;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;
            focusBoard.Visible = false;

        }
        public void ShowMyUserControl2(string user, long id, string Investment_type)
        {
            SessionData.Investment_Type = null;
            settingControl1.SetUserData(user, id);
            investmentsControl1.SetUserData(user, id, Investment_type);
            settingControl1.BringToFront();
            projectsControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            roadsControl1.SendToBack();
            this.dashBoardControl1.SendToBack();
            settingControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;
            dashBoardControl1.Visible = false;
            focusSetting.Visible = true;
            foucsRoad.Visible = false;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;
            focusBoard.Visible = false;
        }
        public void ShowMyUserControl5(string user, long id, string Investment_type)//admin
        {
            SessionData.Investment_Type = null;
            ShowLoading(true);
            systemAdministratorControl1.SetUserData(user, id);


            investmentsControl1.SetUserData(user, id, Investment_type);
            systemAdministratorControl1.Visible = true;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;
            dashBoardControl1.Visible = false;
            systemAdministratorControl1.BringToFront();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            projectsControl1.SendToBack();
            this.dashBoardControl1.SendToBack();
            focusAdmin.Visible = true;
            focusENG.Visible = false;
            focusInvest.Visible = false;           
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            focusBoard.Visible = false;
            ShowLoading(false);
        }
        public void ShowMyUserControl3(string user, long id,string Investment_type)
        {
            SessionData.Investment_Type = Investment_type;
            
            investmentsControl1.SetUserData(user, id, Investment_type);
            projectsControl1.Visible = false;
            investmentsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            dashBoardControl1.Visible = false;
            projectsControl1.SendToBack();
            investmentsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            this.dashBoardControl1.SendToBack();
            focusInvest.Visible = true;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            focusBoard.Visible = false;
        }

        public void ShowMyUserControl7(string user, long id, string Investment_type)
        {
            SessionData.Investment_Type = Investment_type;

            investmentsControl1.SetUserData(user, id, Investment_type);
            projectsControl1.Visible = false;
            investmentsControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            dashBoardControl1.Visible = true;
            projectsControl1.SendToBack();
            investmentsControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            dashBoardControl1.BringToFront();
            focusBoard.Visible = true;
            focusAdmin.Visible = false;
            focusENG.Visible = false;
            focusInvest.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }
        public void ShowMyUserControl4(string user, long id, string Investment_type)
        {
            SessionData.Investment_Type = null;
            roadsControl1.SetUserData(user, id);
            investmentsControl1.SetUserData(user, id, Investment_type);
            roadsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;    
            projectsControl1.Visible = false;
            dashBoardControl1.Visible = false;
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.BringToFront();
            projectsControl1.SendToBack();
            this.dashBoardControl1.SendToBack();
            foucsRoad.Visible = true;
            focusInvest.Visible = false;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            focusBoard.Visible = false;
        }
        public void ShowMyUserControl6(string user, long id, string Investment_type)
        {
            SessionData.Investment_Type = null;
            investmentsControl1.SetUserData(user, id, Investment_type);
            projectsControl1.SetUserData(user, id);
            projectsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            dashBoardControl1.Visible = false;
            projectsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            dashBoardControl1.SendToBack();

            focusInvest.Visible = false;
            focusENG.Visible = true;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            focusBoard.Visible = false;
        }
        public Form2(string username, long user_id)
        {
            
                InitializeComponent();

                _username = username;
                _user_id = user_id;

                this.Text = String.Empty;
                this.ControlBox = false;
                this.DoubleBuffered = true;
                this.AutoScaleMode = AutoScaleMode.Dpi;
                this.MinimumSize = new Size(900, 600);
                this.WindowState = FormWindowState.Maximized;

                // 🔥 أهم جزء
                MakeControlsResponsive();

            
        }
        private void MakeControlsResponsive()
        {
            foreach (Control ctrl in this.Controls)
            {
                if (ctrl is UserControl)
                {
                    ctrl.Dock = DockStyle.Fill;
                }
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
            button1.Visible = false;
            button2.Visible = false;
            button3.Visible = false;
            button4.Visible = false;
            button5.Visible = false;
            button6.Visible = false;
            button7.Visible = false;
            button8.Visible = false;
            Menu.Visible = false;
            flowLayoutPanel1.Width = 36;
            guna2CircleButton1.Location = new System.Drawing.Point(0, 0);
            focusAdmin.Location = new Point(50, 2);
            focusENG.Location = new Point(50, 1);
            focusInvest.Location = new Point(50, 83);
            focusSetting.Location = new Point(50, 40);
            originalFormSize = this.Size;
            //SaveOriginalSizes(this);
            this.AutoScroll = true; // لو الشاشة صغيرة

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
            ShowMyUserControl(_username, _user_id, "كل الاستثمارات");
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
            ShowMyUserControl4(_username, _user_id, "كل الاستثمارات");
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
            Form1.LoginCache.Clear();
            Application.Restart();
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
            ShowMyUserControl5(_username, _user_id, "كل الاستثمارات");  
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ShowMyUserControl3(_username, _user_id, _investment_type);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            TrueFunction();
            ShowMyUserControl(_username, _user_id, "كل الاستثمارات");
            
            FalseFunction();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            TrueFunction();
            ShowMyUserControl2(_username, _user_id, "كل الاستثمارات");
            FalseFunction();
        }

        private void button7_Click(object sender, EventArgs e)
        {

            ShowMyUserControl6(_username, _user_id, "كل الاستثمارات");
           
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

        private void button8_Click(object sender, EventArgs e)
        {
            ShowMyUserControl7(_username, _user_id, _investment_type);
        }

        private void dashBoardControl1_Load(object sender, EventArgs e)
        {

        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            //DialogResult result = MessageBox.Show(
            //"هل تريد إغلاق البرنامج؟",
            //"تأكيد الإغلاق",
            //MessageBoxButtons.YesNo,
            //MessageBoxIcon.Question
            //);

            //if (result == DialogResult.No)
            //{
            //    e.Cancel = true; // ❌ prevent closing
            //}
        }
        //Drag Form
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);
        private void guna2Panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        bool isClicked = true;


        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            
            if (isClicked == true)
            {
                button1.Visible = true;
                button2.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
                button5.Visible = true;
                button6.Visible = true;
                button7.Visible = true;
                button8.Visible = false;
                Menu.Visible = true;
                isClicked = false;
                focusAdmin.Location = new Point(5,2);
                focusENG.Location = new Point(8,1);
                focusInvest.Location = new Point(5,83);
                focusSetting.Location = new Point(5,40);
                guna2CircleButton1.Location= new Point(143,0) ;
                flowLayoutPanel1.Width = 178;
                ApplyPermissions(this);
            }
            else {
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = false;
                button4.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button7.Visible = false;
                button8.Visible = false;
                Menu.Visible = false;
                focusAdmin.Location = new Point(50, 2);
                focusENG.Location = new Point(50, 1);
                focusInvest.Location = new Point(50, 83);
                focusSetting.Location = new Point(50, 40);
                guna2CircleButton1.Location = new System.Drawing.Point(0, 0);
                isClicked = true;
                flowLayoutPanel1.Width = 36;
            }
        }

        private void Form2_Resize(object sender, EventArgs e)
        {
            //float xRatio = (float)this.Width / originalFormSize.Width;
            //float yRatio = (float)this.Height / originalFormSize.Height;

            //ResizeControls(this, xRatio, yRatio);
        }
    }
}

