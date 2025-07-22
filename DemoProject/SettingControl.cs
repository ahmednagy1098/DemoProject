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
    public partial class SettingControl : UserControl
    {
        public SettingControl()
        {
            InitializeComponent();
        }

        private void SettingControl_Load(object sender, EventArgs e)
        {
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
            //dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllHeaders;       
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

            MessageBox.Show("Approvals inserted successfully!");
            this.advancedDataGridView2.DataSource = this.approvalsTableAdapter.GetData();
            guna2TextBox3.Clear();
        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
          
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

       
        }
    }
}
