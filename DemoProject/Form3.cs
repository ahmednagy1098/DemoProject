using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoProject
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
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
        private void Form3_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'database1DataSet.Stations' table. You can move, or remove it, as needed.
            this.stationsTableAdapter.Fill(this.database1DataSet.Stations);
            // TODO: This line of code loads data into the 'database1DataSet.Stations' table. You can move, or remove it, as needed.
            this.stationsTableAdapter.Fill(this.database1DataSet.Stations);
            this.reportViewer1.RefreshReport();
        }
        private void guna2TextBox1_TextChanged(object sender, EventArgs e)
        {
        }
        private void guna2CircleButton2_Click(object sender, EventArgs e)
        {
            // Load all station data once
            var allData = stationsTableAdapter.GetData(); // Returns DataTable
            if (allData == null || allData.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات", AlertForm.AlertType.Error);
                return;
            }
            // Apply filters using LINQ
            var filteredRows = allData.AsEnumerable().Where(row =>
            {
                bool match = true;
                if (liscence_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("license_type_8") == "تم" || row.Field<string>("license_type_8") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("license_type_8") == guna2ComboBox2.Text;
                    }
                }
                if (Air_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("civil_aviation_authority") == "تم" || row.Field<string>("civil_aviation_authority") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("civil_aviation_authority") == guna2ComboBox2.Text;
                    }
                }
                if (Defense_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("civil_defense_approval") == "تم" || row.Field<string>("civil_defense_approval") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("civil_defense_approval") == guna2ComboBox2.Text;
                    }
                }
                if (Enviromental_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("environmental_approval") == "تم" || row.Field<string>("environmental_approval") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("environmental_approval") == guna2ComboBox2.Text;
                    }
                }
                if (Land_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("land_contracts") == "تم" || row.Field<string>("land_contracts") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("land_contracts") == guna2ComboBox2.Text;
                    }
                }
                if (Petrolum_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("ministry_of_petroleum_approval") == "تم" || row.Field<string>("ministry_of_petroleum_approval") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("ministry_of_petroleum_approval") == guna2ComboBox2.Text;
                    }
                }
                if (Study_CB.Checked)
                {
                    if (guna2ComboBox2.Text == "الكل")
                    {
                        match &= row.Field<string>("traffic_study") == "تم" || row.Field<string>("traffic_study") == "جاري";
                    }
                    else
                    {
                        match &= row.Field<string>("traffic_study") == guna2ComboBox2.Text;
                    }
                }
                return match;
            });
            var filteredTable = filteredRows.Any() ? filteredRows.CopyToDataTable() : null;
            if (filteredTable == null || filteredTable.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات مطابقة", AlertForm.AlertType.Warning);
                return;
            }
            // Bind to ReportViewer
            reportViewer1.LocalReport.DataSources.Clear();
            ReportDataSource rds = new ReportDataSource("StationsDataSet", filteredTable);
            reportViewer1.LocalReport.DataSources.Add(rds);
            reportViewer1.RefreshReport();
        }
        public void clear()
        {
            guna2ComboBox2.StartIndex = -1;
            guna2ComboBox3.StartIndex = -1;
            guna2TextBox1.Text = "";
            Land_CB.Checked = false;
            Air_CB.Checked = false;
            Defense_CB.Checked = false;
            liscence_CB.Checked = false;
            Petrolum_CB.Checked = false;
            Study_CB.Checked = false;
            Enviromental_CB.Checked = false;
        }
        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            var AllStations = this.stationsTableAdapter.GetData();
            reportViewer1.LocalReport.DataSources.Clear();
            // Create new data source for the report
            ReportDataSource rds = new ReportDataSource("StationsDataSet", (DataTable)AllStations);
            // "StationsDataSet" must match the name used inside the RDLC report (check it!)
            reportViewer1.LocalReport.DataSources.Add(rds);
            reportViewer1.RefreshReport();
            clear();
        }
    }
}
