using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace DemoProject
{
    public partial class ENGReportForm : Form
    {
        private string _username;
        private long _user_id;
        public ENGReportForm(string username,long user_id)
        {
            InitializeComponent();
            _username = username;
            _user_id = user_id;
        }
        private void guna2GradientTileButton1_Click(object sender, EventArgs e)
        {
        }
        private void guna2GradientTileButton1_Click_1(object sender, EventArgs e)
        {
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = true;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = true;
            guna2GradientTileButton8.Visible = false;

        }
        private void guna2GradientTileButton4_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl2(_username, _user_id, "كل الاستثمارات");
            eng.Show();
            this.Hide();
            //this.Close(); 
        }
        private void guna2GradientTileButton3_Click(object sender, EventArgs e)
        {
            //Form2 eng = new Form2(_username, _user_id);
            //eng.ShowMyUserControl3(_username, _user_id);
            //eng.Show();
            //this.Hide();
            //this.Close();  
            guna2GradientTileButton1.Visible = false;
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton9.Visible = true;// كباري
            guna2GradientTileButton10.Visible = true;// خارج
            guna2GradientTileButton14.Visible = true;//داخل
            guna2GradientTileButton11.Visible = true;//Malls
            guna2GradientTileButton21.Visible=true;
            guna2GradientTileButton12.Visible = true;//مواقف
            guna2GradientTileButton13.Visible = true;
            guna2GradientTileButton21.Visible = true;
            guna2GradientTileButton23.Visible = true;
        }
        private void guna2GradientTileButton2_Click(object sender, EventArgs e)
        {
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton8.Visible = true;

            
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
        private void ENGReportForm_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dATABASE2DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.dATABASE2DataSet.roles);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.dATABASE2DataSet.users);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.pages' table. You can move, or remove it, as needed.
            this.pagesTableAdapter.Fill(this.dATABASE2DataSet.pages);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.functions' table. You can move, or remove it, as needed.
            this.functionsTableAdapter.Fill(this.dATABASE2DataSet.functions);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
            // TODO: This line of code loads data into the 'dATABASE2DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
            guna2GradientTileButton3.TextAlign = HorizontalAlignment.Center;
            guna2GradientTileButton3.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
            this.AutoScaleMode = AutoScaleMode.Dpi;  // Scale based on screen DPI
            this.AutoScaleDimensions = new SizeF(96F, 96F); // Standard DPI baseline
            guna2GradientTileButton3.Tag = "Page:Investment";
            guna2GradientTileButton2.Tag = "Page:Roads";
            guna2GradientTileButton1.Tag = "Page:Eng";
            guna2GradientTileButton6.Tag = "Page:Admin";
            guna2GradientTileButton4.Tag = "Page:Settings";
            // TODO: This line of code loads data into the 'database1DataSet.access' table. You can move, or remove it, as needed.      
            var userRow = this.dATABASE2DataSet.users.FirstOrDefault(u => u.id == _user_id);
            var role = this.dATABASE2DataSet.roles.FirstOrDefault(r=>r.user_id == _user_id);
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
        private void guna2GradientTileButton6_Click(object sender, EventArgs e)
        {

            try
            {
                Form2 eng = new Form2(_username, _user_id);
                eng.Show();
                this.Hide();
                //this.Close(); 
                eng.ShowMyUserControl5(_username, _user_id, "كل الاستثمارات");
            }
            catch (SqlException ex)
            {
                MessageBox.Show("خطأ في الاتصال بقاعدة البيانات من السيرفر: " + ex.Message,
                   "Database Error",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
           
        }
        private void guna2GradientTileButton5_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl(_username, _user_id, "كل الاستثمارات");
            eng.Show();
            this.Hide();
           // this.Close();
        }
        private void guna2GradientTileButton7_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl6(_username, _user_id, "كل الاستثمارات");
            eng.Show();
            this.Hide();
            //this.Close();
        }
        private void guna2GradientTileButton8_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl4(_username, _user_id, "كل الاستثمارات");
            eng.Show();
            this.Hide();
            //this.Close();
        }
        private void guna2GradientTileButton9_Click(object sender, EventArgs e)
        {
        }
        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {
            guna2CircleButton2.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton8.Visible = false;
            guna2GradientTileButton9.Visible = false;
            guna2GradientTileButton10.Visible = false;
            guna2GradientTileButton11.Visible = false;
            guna2GradientTileButton12.Visible = false;
            guna2GradientTileButton13.Visible = false;
            guna2GradientTileButton14.Visible = false;
            guna2GradientTileButton21.Visible = false;
            guna2GradientTileButton23.Visible = false;
            guna2GradientTileButton1.Visible = true;
            guna2GradientTileButton2.Visible = true;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = true;
            guna2GradientTileButton4.Visible = true;
            guna2GradientTileButton6.Visible = true;
            ApplyPermissions(this);
        }
        public void ShowMenuView55()
        {

            guna2CircleButton2.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton8.Visible = false;
            guna2GradientTileButton9.Visible = false;
            guna2GradientTileButton10.Visible = false;
            guna2GradientTileButton11.Visible = false;
            guna2GradientTileButton12.Visible = false;
            guna2GradientTileButton13.Visible = false;
            guna2GradientTileButton14.Visible = false;
            guna2GradientTileButton21.Visible = false;
            guna2GradientTileButton1.Visible = true;
            guna2GradientTileButton2.Visible = true;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = true;
            guna2GradientTileButton4.Visible = true;
            guna2GradientTileButton6.Visible = true;
            ApplyPermissions(this);
        }
        public void ShowMenuView()
        {
           
            ApplyPermissions(this);
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton21.Visible = true;
            guna2GradientTileButton13.Visible = true;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton8.Visible = false;
            guna2GradientTileButton18.Visible = false;
            guna2GradientTileButton19.Visible = false;
            guna2GradientTileButton20.Visible = false;
            guna2GradientTileButton23.Visible = true;
            guna2GradientTileButton9.Visible = true;// كباري
            guna2GradientTileButton10.Visible = true;// خارج
            guna2GradientTileButton14.Visible = true;//داخل
            guna2GradientTileButton11.Visible = true;//Malls
            guna2GradientTileButton12.Visible = true;//مواقف

        }
        public void ShowMenuView2()
        {

            ApplyPermissions(this);
            guna2CircleButton1.Visible = true;
            guna2CircleButton2.Visible = false;
            guna2GradientTileButton9.Visible = false;
            guna2GradientTileButton10.Visible = false;
            guna2GradientTileButton11.Visible = false;
            guna2GradientTileButton21.Visible = false;
            guna2GradientTileButton12.Visible = false;
            guna2GradientTileButton13.Visible = false;
            guna2GradientTileButton14.Visible = false;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton8.Visible = false;
            guna2GradientTileButton15.Visible = false;
            guna2GradientTileButton16.Visible = true;
            guna2GradientTileButton17.Visible = false;
            guna2GradientTileButton18.Visible = false;
            guna2GradientTileButton19.Visible = false;
            guna2GradientTileButton20.Visible = false;

        }
        public void ShowMenuView3()
        {

            ApplyPermissions(this);
            guna2CircleButton1.Visible = true;
            guna2CircleButton2.Visible = false;
            guna2GradientTileButton9.Visible = false;
            guna2GradientTileButton10.Visible = false;
            guna2GradientTileButton11.Visible = false;
            guna2GradientTileButton21.Visible = false;
            guna2GradientTileButton12.Visible = false;
            guna2GradientTileButton13.Visible = false;
            guna2GradientTileButton14.Visible = false;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton8.Visible = false;
            guna2GradientTileButton15.Visible = false;
            guna2GradientTileButton16.Visible = false;
            guna2GradientTileButton17.Visible = false;
            guna2GradientTileButton18.Visible = true;
            guna2GradientTileButton19.Visible = true;
            guna2GradientTileButton20.Visible = true;
            guna2GradientTileButton23.Visible = false;

        }
        public void ShowEngView()
        {
            ApplyPermissions(this);
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = true;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = true;
            guna2GradientTileButton8.Visible = false;

        }
        public void ShowRoadView()
        {
            ApplyPermissions(this);
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton8.Visible = true;

        }

        private void accessBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.accessBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dATABASE2DataSet);

        }

        private void accessBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {
            this.Validate();
            this.accessBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dATABASE2DataSet);

        }

        private void accessBindingNavigatorSaveItem_Click_2(object sender, EventArgs e)
        {
            this.Validate();
            this.accessBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dATABASE2DataSet);

        }

        private void accessBindingNavigatorSaveItem_Click_3(object sender, EventArgs e)
        {
            this.Validate();
            this.accessBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dATABASE2DataSet);

        }
        //اسفل كباري
        //استثمارات على الطرق
        //محلات شل اوت
        private void guna2GradientTileButton10_Click(object sender, EventArgs e)
        {
            guna2GradientTileButton18.Visible = true;
            guna2GradientTileButton19.Visible = true;
            guna2GradientTileButton20.Visible = true;

            guna2CircleButton1.Visible = true;
            guna2CircleButton2.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton9.Visible = false;
            guna2GradientTileButton10.Visible = false;
            guna2GradientTileButton14.Visible = false;
            guna2GradientTileButton11.Visible = false;
            guna2GradientTileButton21.Visible = false;
            guna2GradientTileButton12.Visible = false;
            guna2GradientTileButton13.Visible = false;
            guna2GradientTileButton16.Visible = false;
            guna2GradientTileButton17.Visible = false;
            guna2GradientTileButton15.Visible = false;
            guna2GradientTileButton23.Visible = false;
        }

        private void guna2GradientTileButton9_Click_1(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl3(_username, _user_id, "اسفل كباري");
            eng.Show();
            this.Hide();
            //this.Close();  
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton13_Click(object sender, EventArgs e)
        {

            //guna2GradientTileButton16.Visible = true;
            //guna2GradientTileButton17.Visible = false;
            //guna2GradientTileButton15.Visible = false;

            //guna2CircleButton1.Visible = true;
            //guna2CircleButton2.Visible =false;
            //guna2GradientTileButton2.Visible = false;
            //guna2GradientTileButton22.Visible = false;
            //guna2GradientTileButton3.Visible = false;
            //guna2GradientTileButton4.Visible = false;
            //guna2GradientTileButton5.Visible = false;
            //guna2GradientTileButton6.Visible = false;
            //guna2GradientTileButton7.Visible = false;
            //guna2GradientTileButton1.Visible = false;
            //guna2GradientTileButton9.Visible = false;
            //guna2GradientTileButton10.Visible = false;
            //guna2GradientTileButton14.Visible = false;
            //guna2GradientTileButton11.Visible = false;
            //guna2GradientTileButton21.Visible = false;
            //guna2GradientTileButton12.Visible = false;
            //guna2GradientTileButton13.Visible = false;

            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton15.Text;
            eng.ShowMyUserControl3(_username, _user_id, "استثمارات على الطرق");
            eng.Show();
            this.Hide();
            //this.Close();  
            guna2GradientTileButton1.Visible = false;

        }

        private void guna2GradientTileButton11_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl3(_username, _user_id, "مولات");
            eng.Show();
            this.Hide();
            //this.Close();
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton14_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl3(_username, _user_id, "محلات شل اوت داخل");
            eng.Show();
            this.Hide();
            //this.Close();
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton15_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton15.Text;
            eng.ShowMyUserControl3(_username, _user_id, "استثمارات على الطرق");
            eng.Show();
            this.Hide();
            //this.Close();  
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton16_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton16.Text;
            eng.ShowMyUserControl3(_username, _user_id, "استثمارات على الطرق");
            eng.Show();
            this.Hide();
            //this.Close(); 
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton17_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton17.Text;
            eng.ShowMyUserControl3(_username, _user_id, "استثمارات على الطرق");
            eng.Show();
            this.Hide();
            //this.Close();
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            guna2GradientTileButton16.Visible = false;
            guna2GradientTileButton17.Visible = false;
            guna2GradientTileButton15.Visible = false;

            guna2GradientTileButton18.Visible = false;
            guna2GradientTileButton19.Visible = false;
            guna2GradientTileButton20.Visible = false;
            guna2GradientTileButton23.Visible = true;
            guna2CircleButton1.Visible = false;
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton22.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton9.Visible = true;
            guna2GradientTileButton10.Visible = true;
            guna2GradientTileButton14.Visible = true;
            guna2GradientTileButton11.Visible = true;
            guna2GradientTileButton21.Visible = true;
            guna2GradientTileButton12.Visible = true;
            guna2GradientTileButton13.Visible = true;
        }

        private void guna2GradientTileButton20_Click(object sender, EventArgs e)
        {
            
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton20.Text;
            eng.ShowMyUserControl3(_username, _user_id, "محلات شل اوت خارج");
            eng.Show();
            this.Hide();
            //this.Close();
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton19_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton19.Text;
            eng.ShowMyUserControl3(_username, _user_id, "محلات شل اوت خارج");
            eng.Show();
            this.Hide();
            //this.Close(); 
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton18_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton18.Text;
            eng.ShowMyUserControl3(_username, _user_id, "محلات شل اوت خارج");
            eng.Show();
            this.Hide();
            ////this.Close();   
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton21_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            SessionData.Investment_para1 = guna2GradientTileButton21.Text;
            eng.ShowMyUserControl3(_username, _user_id, "كل الاستثمارات");
            eng.Show();
            this.Hide();
            //this.Close(); 
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton22_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl7(_username, _user_id, "كل الاستثمارات");
            eng.Show();
            this.Hide();
            //this.Close(); 
            guna2GradientTileButton1.Visible = false;
        }

        private void ENGReportForm_FormClosing(object sender, FormClosingEventArgs e)
        {
          //  DialogResult result = MessageBox.Show(
          //"هل تريد إغلاق البرنامج؟",
          //"تأكيد الإغلاق",
          //MessageBoxButtons.YesNo,
          //MessageBoxIcon.Question
          //);

          //  if (result == DialogResult.No)
          //  {
          //      e.Cancel = true; // ❌ prevent closing
          //  }
        }

        private void guna2GradientTileButton12_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl3(_username, _user_id, "مواقف");
            eng.Show();
            this.Hide();
            //this.Close();
            guna2GradientTileButton1.Visible = false;
        }

        private void guna2GradientTileButton23_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl3(_username, _user_id, "مناطق تنموية");
            eng.Show();
            this.Hide();
            //this.Close();
            guna2GradientTileButton1.Visible = false;

        }

        private void guna2Panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
