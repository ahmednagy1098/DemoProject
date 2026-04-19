using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
namespace DemoProject
{
    public partial class DashBoardControl : UserControl
    {
        public DashBoardControl()
        {
            InitializeComponent();
           // DTimeStart.Value = DateTime.Today.AddDays(-7);
           // DTimeEnd.Value = DateTime.Now;
           // btnLast7Days.Select();
        }
        public class ChartItem
        {
            public string Name { get; set; }
            public int Value { get; set; }
            public string investor { get; set; }
        }
        public class AssetItem
        {
            public string Assets { get; set; }
            public int Units { get; set; }
        }
        int allAsits = 0;
        int rentedAsits = 0;
        int nRentedAsits = 0;
        int contractsWorking = 0;
        private void DashBoardControl_Load(object sender, EventArgs e)
        {
            var investments =this.investmentsTableAdapter.GetData();

            foreach (var shops in investments)
            {

                if (shops.Visable_Value == true)
                {
                allAsits += shops.Shops_Count;
                    if(shops.Activity_Name == "لا يوجد" && string.IsNullOrEmpty(shops.investment_name))
                    {
                        nRentedAsits += shops.Shops_Count;
                    }
                    if(shops.Activity_Name != "لا يوجد")
                    {
                        rentedAsits += shops.Shops_Count;
                    }
                }

            }

            AllAsits.Text = $"{allAsits}";
            RentedAsits.Text = $"{rentedAsits}";
            NRentedAsits.Text = $"{nRentedAsits}";

            DataTable dt = new DataTable();
            dt.Columns.Add("Assets", typeof(string));
            dt.Columns.Add("Units", typeof(int));
            dt.Rows.Add("شقق سكنية", 120);
            dt.Rows.Add("محلات تجارية", 45);
            dt.Rows.Add("مكاتب إدارية", 30);
            dt.Rows.Add("مخازن", 18);
            dt.Rows.Add("أراضي", 9);
            dt.Rows.Add("جراجات", 60);
            dt.Columns["Assets"].ColumnName = "الأصول";
            dt.Columns["Units"].ColumnName = "العدد";

            advancedDataGridView1.DataSource = dt;
            chart1.Series["شل اوت"].Points.Clear();
            chart1.Series["اسفل كباري"].Points.Clear();
            chart1.Series["مولات"].Points.Clear();
            chart1.Series["طرق"].Points.Clear();
            var dataShellOut = new List<ChartItem>
            {
                new ChartItem { Name = "يناير", Value = 12000 , investor = "احمد"},
                new ChartItem { Name = "فبراير", Value = 180000,investor = "نور" },
                new ChartItem { Name = "مارس", Value = 9000 ,investor = "ممدوح"},
                new ChartItem { Name = "ابريل", Value = 200000 ,investor = "عبدرحمان"},
                new ChartItem { Name = "مايو", Value = 150000,investor = "فاروق" }
            };

                        var dataUnderBridges = new List<ChartItem>
            {
                new ChartItem { Name = "يناير", Value = 8000 , investor = "احمد"},
                new ChartItem { Name = "فبراير", Value = 60000 , investor = "احمد"},
                new ChartItem { Name = "مارس", Value = 15000, investor = "احمد" },
                new ChartItem { Name = "ابريل", Value = 90000 , investor = "احمد"},
                new ChartItem { Name = "مايو", Value = 40000, investor = "اشرف" }
            };

                        var dataMalls = new List<ChartItem>
            {
                new ChartItem { Name = "يناير", Value = 30000,investor = "نور"  },
                new ChartItem { Name = "فبراير", Value = 120000 ,investor = "نور" },
                new ChartItem { Name = "مارس", Value = 70000,investor = "نور"  },
                new ChartItem { Name = "ابريل", Value = 160000 ,investor = "نور" },
                new ChartItem { Name = "مايو", Value = 110000 ,investor = "عبدرحمان"}
            };

                        var dataRoad = new List<ChartItem>
            {
                new ChartItem { Name = "يناير", Value = 5000,investor = "ممدوح" },
                new ChartItem { Name = "فبراير", Value = 25000,investor = "فاروق" },
                new ChartItem { Name = "مارس", Value = 18000,investor = "فاروق" },
                new ChartItem { Name = "ابريل", Value = 45000,investor = "ممدوح" },
                new ChartItem { Name = "مايو", Value = 30000,investor = "ممدوح" }
            };



            foreach (var item in dataShellOut)
                chart1.Series["شل اوت"].Points.AddXY(item.Name, item.Value);

            foreach (var item in dataUnderBridges)
                chart1.Series["اسفل كباري"].Points.AddXY(item.Name, item.Value);

            foreach (var item in dataMalls)
                chart1.Series["مولات"].Points.AddXY(item.Name, item.Value);

            foreach (var item in dataRoad)
                chart1.Series["طرق"].Points.AddXY(item.Name, item.Value);

            foreach (Series s in chart1.Series)
                s.BorderWidth = 2;
            //chart1.DataBind();

            var allData = dataShellOut
            .Concat(dataUnderBridges)
            .Concat(dataMalls)
            .Concat(dataRoad)
            .ToList();
            var topInvestors = allData
            .GroupBy(x => x.investor)
            .Select(g => new
            {
                Investor = g.Key,
                TotalValue = g.Sum(x => x.Value)
            })
            .OrderByDescending(x => x.TotalValue)
            .Take(5)
            .ToList();
            Series doughnutSeries = chart2.Series["Series1"];
            doughnutSeries.Points.Clear();
            doughnutSeries.ChartType = SeriesChartType.Doughnut;
            doughnutSeries.Label = "#PERCENT{P0}";
            doughnutSeries.LegendText = "#VALX";

            foreach (var item in topInvestors)
            {
                int pointIndex = doughnutSeries.Points.AddXY(item.Investor, item.TotalValue);

                // Tooltip (اختياري)
                doughnutSeries.Points[pointIndex].ToolTip =
                    $"{item.Investor} : {item.TotalValue:N0}";
            }

        }

        private void guna2CircleButton1_Click(object sender, EventArgs e)
        {
            string name = SessionData.UserName;
            long userId = SessionData.UserId;
            ENGReportForm menu = new ENGReportForm(name, userId);
            menu.Show();
            menu.ShowMenuView55();
            Form parentForm = this.FindForm();
            if (parentForm != null)
            {
                parentForm.Close(); // or parentForm.Hide(); if you just want to hide it
            }
        }

        private void headerLabel9_Click(object sender, EventArgs e)
        {
            DTimeEnd.Select();
            DTimeEnd.PerformClick();
        }

        private void headerLabel10_Click(object sender, EventArgs e)
        {
            DTimeStart.Select();
            DTimeStart.PerformClick();
        }

        private void DTimeEnd_ValueChanged(object sender, EventArgs e)
        {
            //headerLabel9.Text = DTimeEnd.Text;
            headerLabel9.Text = DTimeEnd.Value.ToString("yyyy-MM-dd");
          
        }

        private void DTimeStart_ValueChanged(object sender, EventArgs e)
        {
            // headerLabel10.Text = DTimeStart.Text;
            headerLabel10.Text = DTimeStart.Value.ToString("yyyy-MM-dd");
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            guna2Button6.Visible = true;

        }

        private void guna2Button4_Click(object sender, EventArgs e)
        {
            guna2Button6.Visible = false;
        }

        private void btnLast7Days_Click(object sender, EventArgs e)
        {
            guna2Button6.Visible = false;
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            guna2Button6.Visible = false;
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            guna2Button6.Visible = false;
        }

        private void advancedDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
