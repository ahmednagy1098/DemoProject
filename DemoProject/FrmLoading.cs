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
    public partial class FrmLoading : Form
    {
        private Label lblText;
        private ProgressBar progressBar;
        private Panel panel;
        public FrmLoading()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.Black;
            this.Opacity = 0.75;
            this.Size = new Size(350, 120);
            this.TopMost = true;
            this.ShowInTaskbar = false;

            panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(15)
            };

            lblText = new Label
            {
                Text = "جاري تحميل البيانات...",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 11F, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60)
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 18,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };

            panel.Controls.Add(progressBar);
            panel.Controls.Add(lblText);

            this.Controls.Add(panel);

            // Rounded corners (اختياري)
            this.Load += (s, e) =>
            {
                var radius = 15;
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddArc(0, 0, radius, radius, 180, 90);
                path.AddArc(this.Width - radius, 0, radius, radius, 270, 90);
                path.AddArc(this.Width - radius, this.Height - radius, radius, radius, 0, 90);
                path.AddArc(0, this.Height - radius, radius, radius, 90, 90);
                path.CloseAllFigures();
                this.Region = new Region(path);
            };
        }

        /// <summary>
        /// تغيير نص التحميل أثناء التنفيذ
        /// </summary>
        public void SetText(string text)
        {
            if (lblText.InvokeRequired)
                lblText.Invoke(new Action(() => lblText.Text = text));
            else
                lblText.Text = text;
        
        }

    }
}
