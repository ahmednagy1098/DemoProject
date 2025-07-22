using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace DemoProject
{
    public partial class EngineeringManagementControl : UserControl
    {
        public EngineeringManagementControl()
        {
            InitializeComponent();
        }
        public static DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();

            foreach (PropertyDescriptor prop in props)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in props)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }

            return table;
        }
        public void LoadSourceData()
        {
            var LandData = this.landsTableAdapter.GetData();
            var Governorates = this.governorateTableAdapter1.GetData();
            foreach (DataRow row in Governorates.Rows)
            {
                string name = row["governorate"].ToString(); // use column name exactly as in your table
                Governorate_COB.Items.Add(name);
            }
            // Perform the join using LINQ
            var joinedData = from land in LandData.AsEnumerable()
                             join gov in Governorates.AsEnumerable()
                             on land.Field<string>("governorate_fk") equals gov.Field<string>("governorate_id") into gj
                             from gov in gj.DefaultIfEmpty() // Left join in case some lands have no match
                             select new
                             {
                                 land_id = land.Field<string>("land_id"),
                                 land_name = land.Field<string>("land_name"),
                                 total_area = land.Field<string>("total_area"),
                                 Topographic_Survey_Status = land.Field<string>("Topographic_Survey_Status"),
                                 Land_Plate_Status = land.Field<string>("Land_Plate_Status"),
                                 coordinates_N = land.Field<string>("coordinates_N"),
                                 coordinates_E = land.Field<string>("coordinates_E"),
                                 serial_number = land.Field<string>("serial_number"),
                                 plate_number = land.Field<string>("plate_number"),
                                 governorate_fk = gov != null ? gov.Field<string>("governorate") : "", // use governorate name instead of FK
                                 Republican_Decree = land.Field<string>("Republican_Decree"),
                                 Republican_Decree_Status = land.Field<string>("Republican_Decree_Status"),
                                 consulting_Office = land.Field<string>("consulting_Office"),
                                 total_Land_Price = land.Field<decimal>("total_Land_Price").ToString(),
                                 Ownership_Authority = land.Field<string>("Ownership_Authority"),
                                 Address = land.Field<string>("Address")
                             };
            var joinedList = joinedData.ToList();
            DataTable table = ToDataTable(joinedList);
            BindingSource bindingSource = new BindingSource();
            bindingSource.DataSource = table;
            advancedDataGridView1.DataSource = bindingSource;

        }
        public void ApplyGuna2StyleToGrid(DataGridView dgv)
        {
            dgv.EnableHeadersVisualStyles = false;
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.RowHeadersVisible = false;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToResizeRows = false;

            dgv.BackgroundColor = Color.White;
            dgv.GridColor = Color.LightGray;

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 80);
            dgv.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 10F, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.ColumnHeadersHeight = 40;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.Black;
            dgv.DefaultCellStyle.Font = new System.Drawing.Font("Segoe UI", 9.5F, FontStyle.Regular);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(44, 96, 240);
            dgv.DefaultCellStyle.SelectionForeColor = Color.White;
            dgv.BorderStyle = BorderStyle.Fixed3D;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 250, 250);
            dgv.RowTemplate.Height = 30;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            //dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllHeaders;       
        }
        public void GenerativePanalFlow()
        {
            flowLayoutPanel1.AutoScroll = true;
            flowLayoutPanel1.WrapContents = true;  // Items will wrap to next row/column
            flowLayoutPanel1.FlowDirection = FlowDirection.LeftToRight; // Or TopDown
        }
        public void flagforHideColumnsEnG()
        {

           
        }
        private void EngineeringManagementControl_Load(object sender, EventArgs e)
        {
            LoadSourceData();
            ApplyGuna2StyleToGrid(advancedDataGridView1);
            GenerativePanalFlow();
            flagforHideColumnsEnG();
            LoadColumnsIntoCheckedListBox();
        }

        private void dungeonLabel8_Click(object sender, EventArgs e)
        {

        }

        private void guna2ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }
        List<string> lastAddedFilePaths = new List<string>(); // at class level
        public void Files()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "All files (*.*)|*.*";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                foreach (string filePath in openFileDialog.FileNames)
                {
                    string fileName = Path.GetFileName(filePath);
                    byte[] fileData = File.ReadAllBytes(filePath);

                    // SaveFileToDatabase(filePath, fileName, fileData);
                    lastAddedFilePaths.Add(filePath); // Track it

                    AddFileIconToPanel(filePath, fileName);
                }
            }
        }
        private void guna2ImageButton1_Click(object sender, EventArgs e)
        {
            Files();
        }
        private void AddFileIconToPanel(string filePath, string fileName)
        {
            PictureBox picBox = new PictureBox();
            picBox.Size = new Size(64, 64);
            picBox.SizeMode = PictureBoxSizeMode.Zoom;
            picBox.Cursor = Cursors.Hand;
            picBox.BackColor = Color.Transparent;
            // Try to load thumbnail if image
            try
            {
                picBox.Image = Icon.ExtractAssociatedIcon(filePath).ToBitmap();

                // Optional: Show real image if it's an image
                if (IsImageFile(filePath))
                {
                    picBox.Image = System.Drawing.Image.FromFile(filePath);
                }
            }
            catch
            {
                picBox.Image = SystemIcons.Question.ToBitmap(); // fallback icon
            }
            picBox.Click += (s, e) => OpenFile(filePath);
            Label lbl = new Label();
            lbl.Text = fileName;
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            lbl.AutoSize = false;
            lbl.Width = 100;
            lbl.Height = 30;
            lbl.Cursor = Cursors.Hand;
            lbl.Click += (s, e) => OpenFile(filePath); // Label also clickable
            // Container Panel for Icon + Label
            Panel container = new Panel();
            container.Size = new Size(100, 100);
            container.Controls.Add(picBox);
            container.Controls.Add(lbl);

            picBox.Location = new Point(18, 0);
            lbl.Location = new Point(0, 70);

            flowLayoutPanel1.Controls.Add(container);
        }
        private bool IsImageFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext == ".png" || ext == ".jpg" || ext == ".jpeg" || ext == ".bmp" || ext == ".gif";
        }

        private void guna2ImageButton2_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
            foreach (var path in lastAddedFilePaths)
            {
                //DeleteFileFromDatabase(path);
            }

            lastAddedFilePaths.Clear(); // reset
        }
        private void OpenFile(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = filePath,
                        UseShellExecute = true // ensures it opens with default app
                    });
                else
                    ShowAlert(filePath + "الملف غير موجود", AlertForm.AlertType.Error);
            }
            catch (Exception ex)
            {
                ShowAlert("فشل فتح الملف\n" + ex.Message, AlertForm.AlertType.Error);
                //MessageBox.Show("Failed to open file:\n" + ex.Message);
            }
        }

       

        private void skyButton2_Click(object sender, EventArgs e)
        {
            if (advancedDataGridView1.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات لحساب عدد الرفع المساحي", AlertForm.AlertType.Error);
                return;
            }

            string fileName = fileNameTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                ShowAlert("يرجى إدخال اسم للملف قبل التصدير", AlertForm.AlertType.Warning);
                MessageBox.Show("يرجى إدخال اسم للملف قبل التصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fullPath = Path.Combine(desktopPath, $"{fileName}.xlsx");

            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            var workbook = excelApp.Workbooks.Add(Type.Missing);
            var sheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
            sheet.Name = fileName;

            // ✅ Set sheet direction to RTL
            sheet.DisplayRightToLeft = true;

            int excelCol = 1;

            // ✅ Write headers and align right
            for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
            {
                var gridCol = advancedDataGridView1.Columns[col];
                if (gridCol.Visible)
                {
                    var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[1, excelCol];
                    cell.Value = gridCol.HeaderText;
                    cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    cell.Font.Bold = true;
                    excelCol++;
                }
            }

            // ✅ Write data and align right
            int excelRow = 2;
            foreach (DataGridViewRow row in advancedDataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                excelCol = 1;
                for (int col = 0; col < advancedDataGridView1.Columns.Count; col++)
                {
                    var gridCol = advancedDataGridView1.Columns[col];
                    if (!gridCol.Visible) continue;

                    var value = row.Cells[col].Value;
                    var cell = (Microsoft.Office.Interop.Excel.Range)sheet.Cells[excelRow, excelCol];
                    cell.Value = value != null ? value.ToString() : "";
                    cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
                    excelCol++;
                }

                excelRow++;
            }

            // ✅ Auto fit and formatting
            sheet.Columns.AutoFit();
            sheet.Rows.AutoFit();
            sheet.Cells.Font.Size = 12;

            // ✅ Apply plain borders and remove styling
            int totalRows = excelRow - 1;
            int totalCols = excelCol - 1;
            var fullRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[totalRows, totalCols]];

            // Set borders
            fullRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            fullRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

            // Set background to white (remove alternating rows, etc.)
            fullRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            try
            {
                workbook.SaveAs(fullPath);
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                ShowAlert($"حدث خطأ أثناء حفظ الملف:\n{ex.Message}",AlertForm.AlertType.Error);
                //MessageBox.Show($"حدث خطأ أثناء حفظ الملف:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

        private void skyButton3_Click(object sender, EventArgs e)
        {
            if (advancedDataGridView1.Rows.Count == 0)
            {
                ShowAlert("لا يوجد بيانات للتصدير إلى Word", AlertForm.AlertType.Error);
                return;
            }

            string fileName = fileNameTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                ShowAlert("يرجى إدخال اسم للملف قبل التصدير",AlertForm.AlertType.Warning);
               // MessageBox.Show("يرجى إدخال اسم للملف قبل التصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string fullPath = Path.Combine(desktopPath, $"{fileName}.docx");

            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Add();

                // Set paragraph format to RTL globally
                foreach (Microsoft.Office.Interop.Word.Paragraph para in doc.Paragraphs)
                {
                    para.Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                    para.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                }

                var tableData = new List<List<string>>();

                // Add visible headers
                var headers = new List<string>();
                foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                {
                    if (col.Visible)
                        headers.Add(col.HeaderText);
                }
                tableData.Add(headers);

                // Add visible row data
                foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                {
                    if (row.IsNewRow) continue;

                    var rowData = new List<string>();
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (col.Visible)
                        {
                            var value = row.Cells[col.Index].Value;
                            rowData.Add(value?.ToString() ?? "");
                        }
                    }
                    tableData.Add(rowData);
                }

                int rowCount = tableData.Count;
                int colCount = headers.Count;

                var range = doc.Range(0, 0);
                var table = doc.Tables.Add(range, rowCount, colCount);
                table.Borders.Enable = 1;
                table.Range.Font.Size = 11;
                table.Range.Font.Name = "Segoe UI";

                // Set RTL for the whole table
                table.Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                table.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

                // Fill table cells
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        table.Cell(i + 1, j + 1).Range.Text = tableData[i][j];
                        table.Cell(i + 1, j + 1).Range.ParagraphFormat.ReadingOrder = Microsoft.Office.Interop.Word.WdReadingOrder.wdReadingOrderRtl;
                        table.Cell(i + 1, j + 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }

                doc.SaveAs2(fullPath);
                wordApp.Visible = true;
            }
            catch (Exception ex)
            {
                ShowAlert($"حدث خطأ أثناء التصدير إلى Word:\n{ex.Message}",AlertForm.AlertType.Error);
                MessageBox.Show($"حدث خطأ أثناء التصدير إلى Word:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void skyButton1_Click(object sender, EventArgs e)
        {
            
                if (advancedDataGridView1.Rows.Count == 0)
                {
                    ShowAlert("لا يوجد بيانات للتصدير إلى PDF", AlertForm.AlertType.Error);
                    return;
                }

                string fileName = fileNameTextBox.Text.Trim();
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    ShowAlert("يرجى إدخال اسم للملف قبل التصدير",AlertForm.AlertType.Error);
                   // MessageBox.Show("يرجى إدخال اسم للملف قبل التصدير", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string fullPath = Path.Combine(desktopPath, $"{fileName}.pdf");

                try
                {
                    // PDF setup
                    Document pdfDoc = new Document(PageSize.A4, 20f, 20f, 20f, 20f);
                    PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fullPath, FileMode.Create));
                    pdfDoc.Open();

                    // ✅ Support Arabic + RTL
                    string arialPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
                    BaseFont bf = BaseFont.CreateFont(arialPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    iTextSharp.text.Font font = new iTextSharp.text.Font(bf, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                    iTextSharp.text.Font headerFont = new iTextSharp.text.Font(bf, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK);

                    // Count visible columns
                    int visibleColCount = advancedDataGridView1.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible);
                    PdfPTable table = new PdfPTable(visibleColCount);
                    table.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                    table.WidthPercentage = 100;

                    // ✅ Add header cells
                    foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                    {
                        if (!col.Visible) continue;

                        PdfPCell headerCell = new PdfPCell(new Phrase(col.HeaderText, headerFont));
                        headerCell.BackgroundColor = BaseColor.WHITE;
                        headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        headerCell.BorderWidth = 1;
                        headerCell.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                        table.AddCell(headerCell);
                    }

                    // ✅ Add data rows
                    foreach (DataGridViewRow row in advancedDataGridView1.Rows)
                    {
                        if (row.IsNewRow) continue;

                        foreach (DataGridViewColumn col in advancedDataGridView1.Columns)
                        {
                            if (!col.Visible) continue;

                            string cellText = row.Cells[col.Index].Value?.ToString() ?? "";
                            PdfPCell dataCell = new PdfPCell(new Phrase(cellText, font));

                            dataCell.BorderWidth = 1;
                            dataCell.BackgroundColor = BaseColor.WHITE;
                            dataCell.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
                            table.AddCell(dataCell);
                        }
                    }

                    pdfDoc.Add(table);
                    pdfDoc.Close();
                    writer.Close();
                    ShowAlert("تم حفظ الملف بنجاح على سطح المكتب",AlertForm.AlertType.Success);
                    //MessageBox.Show("تم حفظ الملف بنجاح على سطح المكتب", "تم", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    ShowAlert($"حدث خطأ أثناء حفظ ملف PDF:\n{ex.Message}",AlertForm.AlertType.Error);
                    MessageBox.Show($"حدث خطأ أثناء حفظ ملف PDF:\n{ex.Message}", "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            var data = this.landsTableAdapter.GetData().FindByland_id(serial_number_TB.Text);
            if (data != null)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
            }
            bool hasError = false;

            if (string.IsNullOrWhiteSpace(serial_number_TB.Text))
            {
                Error_Serial.Visible = true;
                serial_number_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(plate_number_TB.Text))
            {
                Error_plate.Visible = true;
                plate_number_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Land_Name_TB.Text))
            {
                Error_Land.Visible = true;
                Land_Name_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Address_TB.Text))
            {
                Error_Address.Visible = true;
                Address_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Governorate_COB.Text))
            {
                Error_Governorate.Visible = true;
                Governorate_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Ownership_Authority_TB.Text))
            {
                Error_Ownership.Visible = true;
                Ownership_Authority_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(total_land_price_TB.Text))
            {
                Error_total_land.Visible = true;
                total_land_price_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(total_area_TB.Text))
            {
                Error_total_area.Visible = true;
                total_area_TB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(coordinates_E_TB.Text))
            {
                Error_E.Visible = true;
                coordinates_E_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(coordinates_N_TB.Text))
            {
                Error_N.Visible = true;
                coordinates_N_TB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Topographic_Survey_Status_COB.Text))
            {
                Error_Topographic.Visible = true;
                Topographic_Survey_Status_COB.Focus();
                hasError = true;
            }

            if (string.IsNullOrWhiteSpace(Land_Plate_Status_COB.Text))
            {
                Error_Land_Plate.Visible = true;
                Land_Plate_Status_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(consulting_Office_COB.Text))
            {
                Error_consulting.Visible = true;
                consulting_Office_COB.Focus();
                hasError = true;
            }
            if (string.IsNullOrWhiteSpace(Republican_Decree_Status_COB.Text))
            {
                Error_Republican_Decree.Visible = true;
                Republican_Decree_Status_COB.Focus();
                hasError = true;
            }
            if (hasError)
            {
                ShowAlert("يرجى تعبئة الحقول المطلوبة", AlertForm.AlertType.Error);
                return;
            }
            else
            {
                var IDGovernorate = this.governorateTableAdapter1.GetDataByGovernorate(Governorate_COB.Text).First().governorate_id;

                int IDLand = int.Parse(this.landsTableAdapter.GetData().Last().land_id.ToString());
                IDLand++;

                this.landsTableAdapter.Insert(IDLand.ToString(),
                    Land_Name_TB.Text
                    , total_area_TB.Text
                    , Topographic_Survey_Status_COB.Text
                    , Land_Plate_Status_COB.Text,
                    coordinates_N_TB.Text,
                    coordinates_E_TB.Text,
                    serial_number_TB.Text,
                    plate_number_TB.Text,
                    IDGovernorate,
                    lastAddedFilePaths[0].ToString(),
                    Republican_Decree_Status_COB.Text,
                    consulting_Office_COB.Text,
                    decimal.Parse(total_land_price_TB.Text),
                    Address_TB.Text,
                    Ownership_Authority_TB.Text);
                var x = this.landsTableAdapter.GetData();
                this.advancedDataGridView1.DataSource = x;
                ShowAlert("تمت العملية بنجاح", AlertForm.AlertType.Success);
            }
        }
        public void LoadColumnsIntoCheckedListBox()
        {
            // Clear previous items
            checkedListBox1.Items.Clear();

            // Make sure the DataGridView has a DataSource
            if (advancedDataGridView1.DataSource == null) return;

            // Loop through the columns and add their HeaderText or Name
            foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
            {
                checkedListBox1.Items.Add(column.HeaderText, column.Visible); // Show as checked if visible
                if (column.HeaderText == "مستند القرار الجمهوري")
                {
                    checkedListBox1.Items.Remove(column.HeaderText);
                }
            }
        }
        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                string header = checkedListBox1.Items[e.Index].ToString();

                foreach (DataGridViewColumn column in advancedDataGridView1.Columns)
                {
                    if (column.HeaderText == "مستند القرار الجمهوري")
                    {
                        break;                    
                    }
                    else if (column.HeaderText == header)
                    {
                        column.Visible = checkedListBox1.GetItemChecked(e.Index);
                        break;
                    }
                }
            });
        }

        private void advancedDataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Ensure click is not on header row
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            var clickedColumn = advancedDataGridView1.Columns[e.ColumnIndex];

            // Check if the clicked column is the "pdf" column
            if (clickedColumn.Name == "Republican_Decree_Status")
            {
                // Get the file path from the cell value
                string filePath = advancedDataGridView1.Rows[e.RowIndex].Cells["Republican_Decree"].Value?.ToString();

                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(filePath); // Opens with default PDF reader
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("تعذر فتح الملف: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("الملف غير موجود أو المسار فارغ.");
                }
            }
        }
    }
}
