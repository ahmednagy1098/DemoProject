using System;
using System.Drawing;
using System.Windows.Forms;
using static ReaLTaiizor.Controls.HopeNotify;

namespace DemoProject
{
    public partial class AlertForm : Form
    {
        private Timer timer;
        public AlertForm(string message, AlertType type)
        {
            InitializeComponent();
            SetupAlert(message, type);
        }
        public enum AlertType { Success, Warning, Error, Info }
        private void SetupAlert(string message, AlertType type)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.Size = new Size(300, 60);
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.Opacity = 1.0;

            var label = new Label
            {
                Text = message,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.White,
                Padding = new Padding(10, 0, 0, 0)
            };

            this.Controls.Add(label);

            switch (type)
            {
                case AlertType.Success:
                    this.BackColor = Color.FromArgb(40, 167, 69); // Green
                    break;
                case AlertType.Warning:
                    this.BackColor = Color.FromArgb(255, 193, 7); // Orange
                    break;
                case AlertType.Error:
                    this.BackColor = Color.FromArgb(220, 53, 69); // Red
                    break;
                case AlertType.Info:
                    this.BackColor = Color.FromArgb(0, 123, 255); // Blue
                    break;
            }
            this.Load += AlertForm_Load;
        }
        private void AlertForm_Load(object sender, EventArgs e)
        {
            // Start timer to auto close
            timer = new Timer();
            timer.Interval = 3000; // 3 seconds
            timer.Tick += (s, args) =>
            {
                timer.Stop();
                FadeOut();
            };
            timer.Start();
        }
        private async void FadeOut()
        {
            while (this.Opacity > 0.0)
            {
                await System.Threading.Tasks.Task.Delay(50);
                this.Opacity -= 0.05;
            }
            this.Close();
        }
    }
}
