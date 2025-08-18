using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Timer = System.Windows.Forms.Timer;

namespace DemoProject
{
    public partial class Form2 : Form
    {
        private string _username;
        private long _user_id;

        public async void ShowMyUserControl(string user, long id)
        {
            engineeringManagementControl1.SetUserData(user, id); // Pass data to the control
            projectsControl1.SetUserData(user, id);
            engineeringManagementControl1.BringToFront();
            settingControl1.SendToBack();
            projectsControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            roadsControl1.SendToBack();
            ShowLoading(true);

            await engineeringManagementControl1.LoadSourceDataAsync();

            ShowLoading(false);
            engineeringManagementControl1.Visible = true;
            settingControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;

            focusENG.Visible = true;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;

        }
        public void ShowMyUserControl2(string user, long id)
        {
            settingControl1.SetUserData(user, id);
            settingControl1.BringToFront();
            projectsControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            roadsControl1.SendToBack();

            settingControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;

            focusSetting.Visible = true;
            foucsRoad.Visible = false;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;
        }
        public async void ShowMyUserControl5(string user, long id)//admin
        {
            systemAdministratorControl1.SetUserData(user, id);
            ShowLoading(true);
            await systemAdministratorControl1.LoadSystemAdministratorDataAsync();
            ShowLoading(false);
            systemAdministratorControl1.Visible = true;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;

            systemAdministratorControl1.BringToFront();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            projectsControl1.SendToBack();

            focusAdmin.Visible = true;
            focusENG.Visible = false;
            focusInvest.Visible = false;           
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }
        public async void ShowMyUserControl3(string user, long id)
        {
            investmentsControl1.SetUserData(user, id);
            ShowLoading(true);
            await investmentsControl1.LoadSourceDataAsync();
            ShowLoading(false);
            projectsControl1.Visible = false;
            investmentsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;

            projectsControl1.SendToBack();
            investmentsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();

            focusInvest.Visible = true;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }
        public async void ShowMyUserControl4(string user, long id)
        {
            roadsControl1.SetUserData(user, id);
            ShowLoading(true);
            await roadsControl1.LoadSourceDataAsync();
            ShowLoading(false);
            roadsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;    
            projectsControl1.Visible = false;

            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.BringToFront();
            projectsControl1.SendToBack();

            foucsRoad.Visible = true;
            focusInvest.Visible = false;
            focusENG.Visible = false;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;   
        }
        public async void ShowMyUserControl6(string user, long id)
        {
            projectsControl1.SetUserData(user, id);
            ShowLoading(true);
            await projectsControl1.LoadSourceDataAsync();
            ShowLoading(false);
            projectsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;

            projectsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();

            focusInvest.Visible = false;
            focusENG.Visible = true;
            focusAdmin.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }
        public Form2(string username, long user_id)
        {
            InitializeComponent();
            _username = username;
            _user_id = user_id;
        }
        private static int alertOffsetY = 0;
        bool MenuExpand = false;
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
        private void Form2_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'database1DataSet.roles' table. You can move, or remove it, as needed.
            this.rolesTableAdapter.Fill(this.database1DataSet.roles);
            // TODO: This line of code loads data into the 'database1DataSet.users' table. You can move, or remove it, as needed.
            this.usersTableAdapter.Fill(this.database1DataSet.users);
            // TODO: This line of code loads data into the 'database1DataSet.access' table. You can move, or remove it, as needed.
            this.accessTableAdapter.Fill(this.database1DataSet.access);
            // TODO: This line of code loads data into the 'database1DataSet.functions' table. You can move, or remove it, as needed.
            this.functionsTableAdapter.Fill(this.database1DataSet.functions);
            // TODO: This line of code loads data into the 'database1DataSet.pages' table. You can move, or remove it, as needed.
            this.pagesTableAdapter.Fill(this.database1DataSet.pages);
            //Investment
            //Roads
            //Settings
            //Admin
            //Eng
            
            button1.Tag = "Page:Admin";
            button2.Tag = "Page:Roads";
            Menu.Tag = "Page:Eng";
            button4.Tag = "Page:Investment";
            button5.Tag = "Page:Settings";
            var userRow = database1DataSet.users.FirstOrDefault(u => u.id == _user_id);
            var role = database1DataSet.roles.FirstOrDefault(r => r.user_id == _user_id);
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
        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void menutranstion_Tick(object sender, EventArgs e)
        {
            if (MenuExpand == false)
            {
                MenuContainer.Height += 10;
                if (MenuContainer.Height >= 108) {
                    menutranstion.Stop();
                    MenuExpand = true;
                }
            }
            else
            {
                MenuContainer.Height -= 10;
                if (MenuContainer.Height <= 35)
                {
                    menutranstion.Stop();
                    MenuExpand = false;
                }
            }
        }

        private async void Menu_Click(object sender, EventArgs e)
        {
            ShowLoading(true);

            await engineeringManagementControl1.LoadSourceDataAsync();

            ShowLoading(false);
                engineeringManagementControl1.Visible = true;
                systemAdministratorControl1.Visible = false;
                settingControl1.Visible = false;
                roadsControl1.Visible = false;
                projectsControl1.Visible = false;
                investmentsControl1.Visible = false;

                investmentsControl1.SendToBack();
                engineeringManagementControl1.BringToFront();
                systemAdministratorControl1.SendToBack();
                settingControl1.SendToBack();
                roadsControl1.SendToBack();
                projectsControl1.SendToBack();

                focusENG.Visible = true;
                focusAdmin.Visible = false;
                focusInvest.Visible = false;
                focusSetting.Visible = false;
                foucsRoad.Visible = false;
                menutranstion.Start();
           
        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void guna2DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void advancedDataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)advancedDataGridView1.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    e.ThrowException = false;
                    //((DataGridViewComboBoxColumn)advancedDataGridView1.Columns[e.ColumnIndex]).Items.Add(value);
                }
            }
        }

        private void Survying_TB_Load(object sender, EventArgs e)
        {
            var column = this.advancedDataGridView1.Columns["landcontractsDataGridViewTextBoxColumn1"];
        }
        private void UpdateColumnCounters()
        {
            /* Lands_TB.Text = "0";
             Survying_TB.Text = "0";
             license_8_TB.Text = "0";
             civil_defense_TB.Text = "0";
             environmental_TB.Text = "0";
             ministry_of_petroleum_TB.Text = "0";
             civil_aviation_authority_TB.Text = "0";
             traffic_study_TB.Text = "0";*/
            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                if (column.Name != "land_contracts" &&
                    column.Name != "license_type_8" &&
                    column.Name != "civil_defense_approval" &&
                    column.Name != "environmental_approval" &&
                    column.Name != "ministry_of_petroleum_approval" &&
                    column.Name != "civil_aviation_authority" &&
                    column.Name != "traffic_study" &&
                    column.Name != "Surveying_position")
                {
                    continue;
                }
                int count = 0;
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.Visible && row != null)
                    {
                        var value = row.Cells[column.Index].Value;
                        if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                        {
                            count++;
                            if (column.Name == "land_contracts") { Lands_TB.Text = count.ToString(); }
                            if (column.Name == "Surveying_position") { Survying_TB.Text = count.ToString(); }
                            if (column.Name == "license_type_8") { license_8_TB.Text = count.ToString(); }
                            if (column.Name == "civil_defense_approval") { civil_defense_TB.Text = count.ToString(); }
                            if (column.Name == "environmental_approval") { environmental_TB.Text = count.ToString(); }
                            if (column.Name == "ministry_of_petroleum_approval") { ministry_of_petroleum_TB.Text = count.ToString(); }
                            if (column.Name == "civil_aviation_authority") { civil_aviation_authority_TB.Text = count.ToString(); }
                            if (column.Name == "traffic_study") { traffic_study_TB.Text = count.ToString(); }
                        }
                    }
                }
            }
        }
        bool filterApplied = false;
        private void advancedDataGridView1_FilterStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.FilterEventArgs e)
        {
            filterApplied = true;

        }

        private void advancedDataGridView1_SortStringChanged(object sender, Zuby.ADGV.AdvancedDataGridView.SortEventArgs e)
        {
            //
        }

        private void advancedDataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (filterApplied)
            {
                UpdateColumnCounters(); // Your method using row.Visible
                filterApplied = false;
            }
        }

        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 0 &&
                e.RowIndex >= 0 &&
                (advancedDataGridView1.Columns[e.ColumnIndex].Name == "Pdf8"||
                advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfLand" ||
                advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfCivilDefense" ||
                advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfEnvironmental" ||
                advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfPetroleum" ||
                advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfCivilAviation" ||
                advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfTrafficStudy"
                ))
            {
                // Get the path from the "PdfPath" column in the same row
                string pdfPathLicense_8 = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_license_8"].Value?.ToString();
                string pdfPathLandContracts = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_land_contracts"].Value?.ToString();
                string pdfPathCivilDefense = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_civil_defense"].Value?.ToString();
                string pdfPathEnvironmental = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_environmental"].Value?.ToString();
                string pdfPathMinistryOfPetroleum = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_ministry_of_petroleum"].Value?.ToString();
                string pdfPathCivilAviation = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_civil_aviation"].Value?.ToString();
                string pdfPathTrafficStudy = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_traffic_study"].Value?.ToString();

                if (!string.IsNullOrEmpty(pdfPathLicense_8) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "Pdf8")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathLicense_8,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not open the PDF: " + ex.Message);
                    }
                }
                else if (!string.IsNullOrEmpty(pdfPathLandContracts) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfLand")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathLandContracts,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not open the PDF: " + ex.Message);
                    }
                }
                else if (!string.IsNullOrEmpty(pdfPathCivilDefense) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfCivilDefense")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathCivilDefense,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not open the PDF: " + ex.Message);
                    }
                }
                else if (!string.IsNullOrEmpty(pdfPathEnvironmental) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfEnvironmental")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathEnvironmental,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not open the PDF: " + ex.Message);
                    }
                }
                else if (!string.IsNullOrEmpty(pdfPathMinistryOfPetroleum) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfPetroleum")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathMinistryOfPetroleum,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not open the PDF: " + ex.Message);
                    }
                }
                else if (!string.IsNullOrEmpty(pdfPathCivilAviation) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfCivilAviation")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathCivilAviation,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Could not open the PDF: " + ex.Message);
                    }
                }
                else if (!string.IsNullOrEmpty(pdfPathTrafficStudy) && advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfTrafficStudy")
                {
                    try
                    {
                        // Open the PDF using the default associated application
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = pdfPathTrafficStudy,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        ShowAlert("لا يمكن فتح الملف" + ex.Message , AlertForm.AlertType.Error);
                    }
                }
                else
                {
                    ShowAlert("مسار المستند غير صحيح لم يتم العثور علية", AlertForm.AlertType.Error);
                }
            }
        }

        private void advancedDataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "Pdf8")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_license_8"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfLand")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_land_contracts"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfCivilDefense")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_civil_defense"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfEnvironmental")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_environmental"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfPetroleum")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_ministry_of_petroleum"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfCivilAviation")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_civil_aviation"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
            if (advancedDataGridView1.Columns[e.ColumnIndex].Name == "PdfTrafficStudy")
            {
                string pdfPath = advancedDataGridView1.Rows[e.RowIndex].Cells["Pdf_Path_traffic_study"]?.Value?.ToString();
                if (string.IsNullOrEmpty(pdfPath))
                {
                    e.Value = "غير متاح";  // Button text
                    e.CellStyle.ForeColor = Color.Gray;
                }
                else
                {
                    e.Value = "متاح";
                    e.CellStyle.ForeColor = Color.Black;
                }
            }
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            ShowLoading(true);

            await roadsControl1.LoadSourceDataAsync();

            ShowLoading(false);
            roadsControl1.Visible = true;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            projectsControl1.Visible = false;
            investmentsControl1.Visible = false;

            investmentsControl1.SendToBack();
            roadsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            projectsControl1.SendToBack();

            foucsRoad.Visible = true;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;
            focusENG.Visible = false;
            focusSetting.Visible = false;
            /*Form3 report = new Form3();
            report.Show();*/
        }

        private void Survying_TB_TextChanged(object sender, EventArgs e)
        {

        }

        private void guna2CheckBox4_CheckStateChanged(object sender, EventArgs e)
        {
            if (PdfEnviroment_CB.Checked == false)
            {
                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.Name == "PdfEnvironmental")
                    {
                        column.Visible = false;
                    }
                }
            }
            else
            {
                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.Name == "PdfEnvironmental")
                    {
                        column.Visible = true;
                    }
                }
            }
        }

        private void PdfStudy_CB_CheckStateChanged(object sender, EventArgs e)
        {
            if (PdfStudy_CB.Checked == false)
            {
                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.Name == "PdfTrafficStudy")
                    {
                        column.Visible = false;
                    }
                }
            }
            else
            {
                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.Name == "PdfTrafficStudy")
                    {
                        column.Visible = true;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form1 Login = new Form1();
            Login.Show();
            this.Close();
        }

        private void guna2TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {
            ShowLoading(true);
            await systemAdministratorControl1.LoadSystemAdministratorDataAsync();
            ShowLoading(false);
            systemAdministratorControl1.Visible = true;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;
            investmentsControl1.Visible = false;
            projectsControl1.Visible = false;

            systemAdministratorControl1.BringToFront();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();
            projectsControl1.SendToBack();
            investmentsControl1.SendToBack();


            focusAdmin.Visible = true;
            focusInvest.Visible = false;
            focusENG.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;

        }

        private async void button4_Click(object sender, EventArgs e)
        {
            ShowLoading(true);
            await investmentsControl1.LoadSourceDataAsync();
            ShowLoading(false);
            investmentsControl1.Visible = true;
            projectsControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;

            investmentsControl1.BringToFront();
            systemAdministratorControl1.SendToBack();
            projectsControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();

            focusInvest.Visible = true;
            focusAdmin.Visible = false;
            focusENG.Visible = false;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            engineeringManagementControl1.Visible = true;
            engineeringManagementControl1.BringToFront();
            systemAdministratorControl1.Visible = false;
            systemAdministratorControl1.SendToBack();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            settingControl1.Visible = true;
            engineeringManagementControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            roadsControl1.Visible = false;
            projectsControl1.Visible = false;
            investmentsControl1.Visible = false;

            investmentsControl1.SendToBack();
            settingControl1.BringToFront();
            engineeringManagementControl1.SendToBack();
            roadsControl1.SendToBack();
            projectsControl1.SendToBack();
            systemAdministratorControl1.SendToBack();

            focusSetting.Visible = true;
            focusAdmin.Visible = false;
            focusInvest.Visible = false;
            focusENG.Visible = false;
            foucsRoad.Visible = false;
        }

        private async void button7_Click(object sender, EventArgs e)
        {
            ShowLoading(true);
            await projectsControl1.LoadSourceDataAsync();
            ShowLoading(false);
            projectsControl1.Visible = true;
            investmentsControl1.Visible = false;
            systemAdministratorControl1.Visible = false;
            engineeringManagementControl1.Visible = false;
            settingControl1.Visible = false;
            roadsControl1.Visible = false;

            projectsControl1.BringToFront();
            investmentsControl1.SendToBack();
            systemAdministratorControl1.SendToBack();
            engineeringManagementControl1.SendToBack();
            settingControl1.SendToBack();
            roadsControl1.SendToBack();

            focusInvest.Visible = false;
            focusAdmin.Visible = false;
            focusENG.Visible = true;
            focusSetting.Visible = false;
            foucsRoad.Visible = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (guna2CircleProgressBar1.Value < 100)
            {
                guna2CircleProgressBar1.Value += 2; // speed of loading
            }
            else
            {
                guna2CircleProgressBar1.Value = 0; // reset for loop effect
            }
        }

        public void ShowLoading(bool show)
        {
            guna2CircleProgressBar1.Visible = show;
            guna2CircleProgressBar1.Value = 0;
            guna2CircleProgressBar1.BringToFront();

            if (show)

                timer1.Start();
            else
                timer1.Stop();
        }
    }
}

