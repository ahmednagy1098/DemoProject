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
    public partial class SelectLocationDialog : Form
    {
        public string SelectedLandId { get; private set; }
        public string SelectedGovernorateId { get; private set; }
        public string SelectedNeighborhood { get; private set; }
        private DataTable _landsTable;
        public SelectLocationDialog(DataTable lands, DataTable governorates)
        {
            InitializeComponent();
            _landsTable = lands; // 👈 مهم
            // Load Lands
            cbLand.DataSource = _landsTable;
            cbLand.DisplayMember = "land_id"; // غيرها حسب اسم العمود عندك
            cbLand.ValueMember = "land_id";

            // Load Governorates
            cbGovernorate.DataSource = governorates;
            cbGovernorate.DisplayMember = "governorate";
            cbGovernorate.ValueMember = "governorate_id";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            SelectedLandId = cbLand.SelectedValue?.ToString() ?? "0";
            SelectedGovernorateId = cbGovernorate.SelectedValue?.ToString() ?? "0";
            SelectedNeighborhood = txtNeighborhood.Text;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void SelectLocationDialog_Load(object sender, EventArgs e)
        {

        }

        private void cbGovernorate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbGovernorate.SelectedValue == null) return;

            string govId = cbGovernorate.SelectedValue.ToString();

            DataView dv = new DataView(_landsTable);
            dv.RowFilter = $"governorate_fk = '{govId}'";

            cbLand.DataSource = dv;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectedLandId = cbLand.SelectedValue?.ToString() ?? "0";
            SelectedGovernorateId = cbGovernorate.SelectedValue?.ToString() ?? "0";
            SelectedNeighborhood = txtNeighborhood.Text;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void airButton1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
