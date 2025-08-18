using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = true;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = true;
            guna2GradientTileButton8.Visible = false;
            guna2GradientTileButton9.Visible = false;
            guna2GradientTileButton10.Visible = false;
        }
        private void guna2GradientTileButton4_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl2(_username, _user_id);
            eng.Show();
            this.Hide();
            this.Close();         
        }
        private void guna2GradientTileButton3_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl3(_username, _user_id);
            eng.Show();
            this.Hide();
            this.Close();
        }
        private void guna2GradientTileButton2_Click(object sender, EventArgs e)
        {
            guna2CircleButton2.Visible = true;
            guna2GradientTileButton1.Visible = false;
            guna2GradientTileButton2.Visible = false;
            guna2GradientTileButton3.Visible = false;
            guna2GradientTileButton4.Visible = false;
            guna2GradientTileButton5.Visible = false;
            guna2GradientTileButton6.Visible = false;
            guna2GradientTileButton7.Visible = false;
            guna2GradientTileButton8.Visible = true;
            guna2GradientTileButton9.Visible = true;
            guna2GradientTileButton10.Visible = true;
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
            this.AutoScaleMode = AutoScaleMode.Dpi;  // Scale based on screen DPI
            this.AutoScaleDimensions = new SizeF(96F, 96F); // Standard DPI baseline
            // TODO: This line of code loads data into the 'database1DataSet.functions' table. You can move, or remove it, as needed.
            this.functionsTableAdapter.Fill(this.database1DataSet.functions);
            // TODO: This line of code loads data into the 'database1DataSet.pages' table. You can move, or remove it, as needed.
            this.pagesTableAdapter.Fill(this.database1DataSet.pages);
            guna2GradientTileButton3.Tag = "Page:Investment";
            guna2GradientTileButton2.Tag = "Page:Roads";
            guna2GradientTileButton1.Tag = "Page:Eng";
            guna2GradientTileButton6.Tag = "Page:Admin";
            guna2GradientTileButton4.Tag = "Page:Settings";
            // TODO: This line of code loads data into the 'database1DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.database1DataSet.access);
            // TODO: This line of code loads data into the 'database1DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.database1DataSet.users);
            // TODO: This line of code loads data into the 'database1DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.database1DataSet.roles);         
            var userRow = database1DataSet.users.FirstOrDefault(u => u.id == _user_id);
            var role = database1DataSet.roles.FirstOrDefault(r=>r.user_id == _user_id);
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
        private void guna2GradientTileButton6_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.Show();
            this.Hide();
            this.Close();
            eng.ShowMyUserControl5(_username, _user_id);
        }
        private void guna2GradientTileButton5_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl(_username, _user_id);
            eng.Show();
            this.Hide();
            this.Close();
        }
        private void guna2GradientTileButton7_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl6(_username, _user_id);
            eng.Show();
            this.Hide();
            this.Close();
        }
        private void guna2GradientTileButton8_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2(_username, _user_id);
            eng.ShowMyUserControl4(_username, _user_id);
            eng.Show();
            this.Hide();
            this.Close();
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
            guna2GradientTileButton1.Visible = true;
            guna2GradientTileButton2.Visible = true;
            guna2GradientTileButton3.Visible = true;
            guna2GradientTileButton4.Visible = true;
            guna2GradientTileButton6.Visible = true;
            ApplyPermissions(this);
        }
        public void ShowMenuView()
        {
            ApplyPermissions(this);
        }
    }
}
