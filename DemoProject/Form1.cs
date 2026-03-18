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
using System.Data.SqlClient;
using Microsoft.AspNetCore.SignalR.Client;
using System.Runtime.Caching;
namespace DemoProject
{
    
    public partial class Form1 : Form
    {
        //SoundPlayer playeWelcome = new SoundPlayer(@"C:\Users\PC1\Downloads\welcomeLog_Manager1.wav");
        HubConnection hubConnection;
        public Form1()
        {
            InitializeComponent();
            //hubConnection = new HubConnectionBuilder().WithUrl("https://localhost:7212/ChatHub").Build();
            //hubConnection.Closed += HubConnection_Closed;
            //hubConnection.On<string, string>("ReceiveMessage", (user, message) =>
            //{
            //    var newMessage = $"{user} : {message}";
            //    ShowAlert(newMessage, AlertForm.AlertType.Success);
            //});
        }
        //private async Task HubConnection_Closed(Exception arg)
        //{
        //    await Task.Delay(new Random().Next(0,5) * 1000);
        //    await hubConnection.StartAsync();
        //}
        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        public static class LoginCache
        {
            private static readonly MemoryCache cache = MemoryCache.Default;
            private const string KEY = "LOGGED_USER";

            public static void Set(DATABASE2DataSet.usersRow user)
            {
                cache.Set(KEY, user, DateTimeOffset.Now.AddHours(12));

                // حفظ دائم
                Properties.Settings.Default.LastUserId = user.id;
                Properties.Settings.Default.Save();
            }

            public static DATABASE2DataSet.usersRow Get()
            {
                return cache.Get(KEY) as DATABASE2DataSet.usersRow;
            }

            public static void Clear()
            {
                cache.Remove(KEY);
                Properties.Settings.Default.LastUserId = 0;
                Properties.Settings.Default.Save();
            }
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
        private DATABASE2DataSet.usersRow AuthenticateUser(string username, string password)
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

                LoginCache.Set(matchedUser); // ✅ caching

                this.DialogResult = DialogResult.OK;
                this.Close();
      
                ShowAlert("مرحبا بك", AlertForm.AlertType.Success);
            }

            else
            {
                ShowAlert("يرجي كتابة اسم المستخدم وكلمة السر بشكل صحيح", AlertForm.AlertType.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                // 🔐 Auto Login
                long savedUserId = Properties.Settings.Default.LastUserId;

                if (savedUserId > 0)
                {
                    var users = this.usersTableAdapter.GetData();
                    var user = users.FirstOrDefault(u => u.id == savedUserId);

                    if (user != null)
                    {
                        SessionData.UserId = user.id;
                        SessionData.UserName = user.user_name;

                        ShowAlert("تم تسجيل الدخول تلقائياً", AlertForm.AlertType.Success);

                        this.DialogResult = DialogResult.OK;
                        this.Close();

                        return;
                    }
                }
                // TODO: This line of code loads data into the 'dATABASE2DataSet.functions' table. You can move, or remove it, as needed.
                this.functionsTableAdapter.Fill(this.dATABASE2DataSet.functions);
                // TODO: This line of code loads data into the 'dATABASE2DataSet.access' table. You can move, or remove it, as needed.
                this.accessTableAdapter.Fill(this.dATABASE2DataSet.access);
                // TODO: This line of code loads data into the 'dATABASE2DataSet.users' table. You can move, or remove it, as needed.
                this.usersTableAdapter.Fill(this.dATABASE2DataSet.users);
                // TODO: This line of code loads data into the 'dATABASE2DataSet.roles' table. You can move, or remove it, as needed.
                this.rolesTableAdapter.Fill(this.dATABASE2DataSet.roles);
                // TODO: This line of code loads data into the 'dATABASE2DataSet.pages' table. You can move, or remove it, as needed.
                this.pagesTableAdapter.Fill(this.dATABASE2DataSet.pages);
            }
            catch (SqlException ex)
            {
                MessageBox.Show("خطأ في الاتصال بقاعدة البيانات من السيرفر: " + ex.Message,
                   "Database Error",
                   MessageBoxButtons.OK,
                   MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show("حدث خطأ غير متوقع: " + ex.Message,
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }

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

                else
                {
                    ShowAlert("يرجي كتابة اسم المستخدم وكلمة السر بشكل صحيح", AlertForm.AlertType.Error);
                }
            }
        }

        private void pagesBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.pagesBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.dATABASE2DataSet);

        }
      
    }
}
