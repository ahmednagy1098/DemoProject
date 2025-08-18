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
using System.Media;

namespace DemoProject
{
    
    public partial class Form1 : Form
    {
        //SoundPlayer playeWelcome = new SoundPlayer(@"C:\Users\PC1\Downloads\welcomeLog_Manager1.wav");
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
        private Database1DataSet.usersRow AuthenticateUser(string username, string password)
        {
            var users = this.usersTableAdapter.GetData();
            
            return users.FirstOrDefault(u =>
                u.user_name == username &&
                u.password == password
                );
        }
        
        private void guna2Button1_Click(object sender, EventArgs e)
        {
            string inputUsername = UserNameTB.Text.Trim();
            string inputPassword = PassWordTB.Text.Trim();
            var matchedUser = AuthenticateUser(inputUsername, inputPassword);

            if (matchedUser != null)
            {
                SessionData.UserId = matchedUser.id;
                SessionData.UserName = matchedUser.user_name;
                this.Hide();
                //playeWelcome.Play();
                ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
                ENGReportForm menu = new ENGReportForm(matchedUser.user_name,matchedUser.id);
                menu.Show();
            }
            //else if (UserNameTB.Text=="المدير العام" && PassWordTB.Text =="123" )
            //{            
            //    this.Hide();
            //    ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
            //    ENGReportForm Menu = new ENGReportForm(matchedUser.user_name, matchedUser.password, matchedUser.id);
            //    Menu.Show();
             
            //}
            else
            {
                ShowAlert("يرجي كتابة اسم المستخدم وكلمة السر بشكل صحيح", AlertForm.AlertType.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'database1DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.database1DataSet.access);
            // TODO: This line of code loads data into the 'database1DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.database1DataSet.roles);
            // TODO: This line of code loads data into the 'database1DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.database1DataSet.users);

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

        private void PassWordTB_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '\r')
            {
                string inputUsername = UserNameTB.Text.Trim();
                string inputPassword = PassWordTB.Text.Trim();
                var matchedUser = AuthenticateUser(inputUsername, inputPassword);
                if (matchedUser!=null)
                {
                    SessionData.UserId = matchedUser.id;
                    SessionData.UserName = matchedUser.user_name;
                    this.Hide();
                    ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
                    ENGReportForm menu = new ENGReportForm(matchedUser.user_name, matchedUser.id);
                    menu.Show();
                }
               //else if (UserNameTB.Text == "المدير العام" && PassWordTB.Text == "123")
               // {
               //     this.Hide();
               //     ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
               //     ENGReportForm Menu = new ENGReportForm(matchedUser.user_name, matchedUser.password, matchedUser.id);

               //     Menu.Show();

               // }
                else
                {
                    ShowAlert("يرجي كتابة اسم المستخدم وكلمة السر بشكل صحيح", AlertForm.AlertType.Error);
                }
            }
        }

        private void UserNameTB_KeyPress(object sender, KeyPressEventArgs e)
        {

            if (e.KeyChar == '\r')
            {
                string inputUsername = UserNameTB.Text.Trim();
                string inputPassword = PassWordTB.Text.Trim();
                var matchedUser = AuthenticateUser(inputUsername, inputPassword);
                if (matchedUser != null)
                {
                    SessionData.UserId = matchedUser.id;
                    SessionData.UserName = matchedUser.user_name;
                    this.Hide();
                    ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
                    ENGReportForm menu = new ENGReportForm(matchedUser.user_name, matchedUser.id);
                    menu.Show();
                }
              /*else if (UserNameTB.Text == "المدير العام" && PassWordTB.Text == "123")
                {
                    this.Hide();
                    ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
                    ENGReportForm Menu = new ENGReportForm(matchedUser.user_name, matchedUser.password, matchedUser.id);

                    Menu.Show();

                }*/
                else
                {
                    ShowAlert("يرجي كتابة اسم المستخدم وكلمة السر بشكل صحيح", AlertForm.AlertType.Error);
                }
            }
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
