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
        public ENGReportForm()
        {
            InitializeComponent();
        }

        private void guna2GradientTileButton1_Click(object sender, EventArgs e)
        {

        }

        private void guna2GradientTileButton1_Click_1(object sender, EventArgs e)
        {
            Form2 eng = new Form2();
            eng.Show();
            this.Hide();
            this.Close();
            eng.ShowMyUserControl();
        }

        private void guna2GradientTileButton4_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2();
            eng.Show();
            this.Hide();
            this.Close();
            eng.ShowMyUserControl2();
        }

        private void guna2GradientTileButton3_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2();
            eng.Show();
            this.Hide();
            this.Close();
            eng.ShowMyUserControl3();
        }

        private void guna2GradientTileButton2_Click(object sender, EventArgs e)
        {
            Form2 eng = new Form2();
            eng.Show();
            this.Hide();
            this.Close();
            eng.ShowMyUserControl4();
        }
    }
}
