using Guna.UI2.WinForms;
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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {//show
            if(PassWordTB.PasswordChar == '*')
            {
                guna2CircleButton2.Visible=true;
                guna2CircleButton1.Visible = false;
                guna2CircleButton2.BringToFront();
                guna2CircleButton1.SendToBack();
                PassWordTB.PasswordChar = '\0';
            }
        }

        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {//hide
            if (PassWordTB.PasswordChar == '\0')
            {
                guna2CircleButton2.Visible = false;
                guna2CircleButton1.Visible = true;
                guna2CircleButton1.BringToFront();
                guna2CircleButton2.SendToBack();
                PassWordTB.PasswordChar = '*';
            }
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            
            if (UserNameTB.Text=="المدير العام" && PassWordTB.Text =="123" )
            {            
                this.Hide();
                ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
                Form2 Menu = new Form2();
                Menu.Show();
             
            }
            else
            {
                ShowAlert("يرجي كتابة اسم المستخدم وكلمة السر بشكل صحيح", AlertForm.AlertType.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
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
        //private void btnSuccess_Click(object sender, EventArgs e)
        //{
        //    ShowAlert("Success Alert", AlertForm.AlertType.Success);
        //}

        //private void btnWarning_Click(object sender, EventArgs e)
        //{
        //    ShowAlert("Warning Alert", AlertForm.AlertType.Warning);
        //}

        //private void btnError_Click(object sender, EventArgs e)
        //{
        //    ShowAlert("Error Alert", AlertForm.AlertType.Error);
        //}

        //private void btnInfo_Click(object sender, EventArgs e)
        //{
        //    ShowAlert("Info Alert", AlertForm.AlertType.Info);
        //}
    }
}
