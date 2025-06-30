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
            if (e.Exception.Message == "DataGridViewComboBoxCell value is not valid.")
            {
                object value = advancedDataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                if (!((DataGridViewComboBoxColumn)advancedDataGridView1.Columns[e.ColumnIndex]).Items.Contains(value))
                {
                    e.ThrowException = false;
                    ((DataGridViewComboBoxColumn)advancedDataGridView1.Columns[e.ColumnIndex]).Items.Add(value);
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
            if (e.RowIndex >= 0 &&
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

        private void button2_Click(object sender, EventArgs e)
        {
            Form3 report = new Form3();
            report.Show();
        }

        private void Survying_TB_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

