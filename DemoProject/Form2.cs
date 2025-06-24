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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
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
        private void Form2_Load(object sender, EventArgs e)
        {
            //رقم المعاملة
            // TODO: This line of code loads data into the 'database1DataSet.Surveying_position_states' table. You can move, or remove it, as needed.
            this.surveying_position_statesTableAdapter.Fill(this.database1DataSet.Surveying_position_states);
            // TODO: This line of code loads data into the 'database1DataSet.status' table. You can move, or remove it, as needed.
            this.statusTableAdapter.Fill(this.database1DataSet.status);
            // TODO: This line of code loads data into the 'database1DataSet.status' table. You can move, or remove it, as needed.
            this.statusTableAdapter.Fill(this.database1DataSet.status);
            // TODO: This line of code loads data into the 'database1DataSet.Table' table. You can move, or remove it, as needed.
            this.tableTableAdapter.Fill(this.database1DataSet.Table);
            // TODO: This line of code loads data into the 'database1DataSet.Stations' table. You can move, or remove it, as needed.
            this.stationsTableAdapter.Fill(this.database1DataSet.Stations);
            var Survying = this.stationsTableAdapter.Count_All_Values_Surveying();
            var Lands = this.stationsTableAdapter.Count_All_Values_Lands();
            var license_8 = this.stationsTableAdapter.Count_All_Values_License_type_8();
            var civil_defense = this.stationsTableAdapter.Count_All_Values_civil_defense();
            var environmental = this.stationsTableAdapter.Count_All_Values_environmental();
            var ministry_of_petroleum = this.stationsTableAdapter.Count_All_Values_ministry_of_petroleum();
            var civil_aviation_authority = this.stationsTableAdapter.Count_All_Values_civil_aviation_authority();
            var traffic_study = this.stationsTableAdapter.Count_All_Values_traffic_study();
            if (Survying == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد الرفع المساحي", AlertForm.AlertType.Error);
                Survying_TB.Text = "0";
            }
            if (license_8 == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد نموذج 8", AlertForm.AlertType.Error);
                license_8_TB.Text = "0";
            }
            if (civil_defense == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد الحماية المدنية", AlertForm.AlertType.Error);
                civil_defense_TB.Text = "0";
            }
            if (environmental == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد وزارة البيئة", AlertForm.AlertType.Error);
                environmental_TB.Text = "0";
            }
            if (ministry_of_petroleum == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد وزارة البترول", AlertForm.AlertType.Error);
                ministry_of_petroleum_TB.Text = "0";
            }
            if (civil_aviation_authority == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد الطيران المدني", AlertForm.AlertType.Error);
                civil_aviation_authority_TB.Text = "0";
            }
            if (traffic_study == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد الدراسة المرورية", AlertForm.AlertType.Error);
                traffic_study_TB.Text = "0";
            }
            if (Lands == null)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد عقود الأراضي", AlertForm.AlertType.Error);
                Lands_TB.Text = "0";
            }
            else
            {
                //var column = this.advancedDataGridView1.Columns["landcontractsDataGridViewTextBoxColumn1"];
                Survying_TB.Text = Survying.ToString();
                Lands_TB.Text = Lands.ToString();
                license_8_TB.Text = license_8.ToString();
                civil_defense_TB.Text = civil_defense.ToString();
                civil_aviation_authority_TB.Text = civil_aviation_authority.ToString();
                traffic_study_TB.Text = traffic_study.ToString();
                environmental_TB.Text = environmental.ToString();
                ministry_of_petroleum_TB.Text = ministry_of_petroleum.ToString();
            }
        }
        private void flowLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void menutranstion_Tick(object sender, EventArgs e)
        {
            if(MenuExpand == false)
            {
                MenuContainer.Height += 10;
                if(MenuContainer.Height >= 108) {
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

        private void Menu_Click(object sender, EventArgs e)
        {
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
            if(e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)advancedDataGridView1.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    ((DataGridViewComboBoxColumn)advancedDataGridView1.Columns[e.ColumnIndex]).Items.Add(value);
                    e.ThrowException = false;
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
                            if(column.Name== "land_contracts"){Lands_TB.Text = count.ToString();}
                            if (column.Name == "Surveying_position"){Survying_TB.Text = count.ToString();}
                            if (column.Name == "license_type_8") {license_8_TB.Text= count.ToString(); }
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
    }
}

