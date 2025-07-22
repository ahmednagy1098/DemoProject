namespace DemoProject
{
    partial class Form3
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.stationsBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.database1DataSet = new DemoProject.Database1DataSet();
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.stationsTableAdapter = new DemoProject.Database1DataSetTableAdapters.StationsTableAdapter();
            this.guna2CircleButton2 = new Guna.UI2.WinForms.Guna2CircleButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.Petrolum_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.liscence_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.Defense_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.Land_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.Air_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.Study_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.guna2CircleButton1 = new Guna.UI2.WinForms.Guna2CircleButton();
            this.Enviromental_CB = new Guna.UI2.WinForms.Guna2CheckBox();
            this.guna2Panel1 = new Guna.UI2.WinForms.Guna2Panel();
            this.guna2TextBox1 = new Guna.UI2.WinForms.Guna2TextBox();
            this.guna2ComboBox2 = new Guna.UI2.WinForms.Guna2ComboBox();
            this.guna2ComboBox3 = new Guna.UI2.WinForms.Guna2ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.stationsBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.database1DataSet)).BeginInit();
            this.guna2Panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // stationsBindingSource
            // 
            this.stationsBindingSource.DataMember = "Stations";
            this.stationsBindingSource.DataSource = this.database1DataSet;
            // 
            // database1DataSet
            // 
            this.database1DataSet.DataSetName = "Database1DataSet";
            this.database1DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "StationsDataSet";
            reportDataSource1.Value = this.stationsBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "DemoProject.Report1.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 84);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.ServerReport.BearerToken = null;
            this.reportViewer1.Size = new System.Drawing.Size(908, 366);
            this.reportViewer1.TabIndex = 0;
            // 
            // stationsTableAdapter
            // 
            this.stationsTableAdapter.ClearBeforeFill = true;
            // 
            // guna2CircleButton2
            // 
            this.guna2CircleButton2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.guna2CircleButton2.Animated = true;
            this.guna2CircleButton2.BackColor = System.Drawing.Color.Transparent;
            this.guna2CircleButton2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("guna2CircleButton2.BackgroundImage")));
            this.guna2CircleButton2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.guna2CircleButton2.DisabledState.BorderColor = System.Drawing.Color.DarkGray;
            this.guna2CircleButton2.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray;
            this.guna2CircleButton2.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(169)))), ((int)(((byte)(169)))));
            this.guna2CircleButton2.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(141)))), ((int)(((byte)(141)))), ((int)(((byte)(141)))));
            this.guna2CircleButton2.FillColor = System.Drawing.Color.Transparent;
            this.guna2CircleButton2.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.guna2CircleButton2.ForeColor = System.Drawing.Color.White;
            this.guna2CircleButton2.IndicateFocus = true;
            this.guna2CircleButton2.Location = new System.Drawing.Point(867, 40);
            this.guna2CircleButton2.Name = "guna2CircleButton2";
            this.guna2CircleButton2.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle;
            this.guna2CircleButton2.Size = new System.Drawing.Size(38, 31);
            this.guna2CircleButton2.TabIndex = 13;
            this.guna2CircleButton2.UseTransparentBackground = true;
            this.guna2CircleButton2.Click += new System.EventHandler(this.guna2CircleButton2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(114, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 20);
            this.label1.TabIndex = 18;
            this.label1.Text = "وثائق للبحث";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(388, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 20);
            this.label2.TabIndex = 19;
            this.label2.Text = "قيم البحث بل وثائق";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(555, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 20);
            this.label3.TabIndex = 20;
            this.label3.Text = "الرفع المساحي";
            // 
            // Petrolum_CB
            // 
            this.Petrolum_CB.AutoSize = true;
            this.Petrolum_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Petrolum_CB.CheckedState.BorderRadius = 4;
            this.Petrolum_CB.CheckedState.BorderThickness = 0;
            this.Petrolum_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Petrolum_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Petrolum_CB.Location = new System.Drawing.Point(118, 54);
            this.Petrolum_CB.Name = "Petrolum_CB";
            this.Petrolum_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Petrolum_CB.Size = new System.Drawing.Size(119, 17);
            this.Petrolum_CB.TabIndex = 21;
            this.Petrolum_CB.Text = "موافقة وزارة البترول";
            this.Petrolum_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Petrolum_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.Petrolum_CB.UncheckedState.BorderRadius = 4;
            this.Petrolum_CB.UncheckedState.BorderThickness = 0;
            this.Petrolum_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // liscence_CB
            // 
            this.liscence_CB.AutoSize = true;
            this.liscence_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.liscence_CB.CheckedState.BorderRadius = 4;
            this.liscence_CB.CheckedState.BorderThickness = 0;
            this.liscence_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.liscence_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.liscence_CB.Location = new System.Drawing.Point(326, 54);
            this.liscence_CB.Name = "liscence_CB";
            this.liscence_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.liscence_CB.Size = new System.Drawing.Size(63, 17);
            this.liscence_CB.TabIndex = 22;
            this.liscence_CB.Text = "نموذج 8";
            this.liscence_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.liscence_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.liscence_CB.UncheckedState.BorderRadius = 4;
            this.liscence_CB.UncheckedState.BorderThickness = 0;
            this.liscence_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // Defense_CB
            // 
            this.Defense_CB.AutoSize = true;
            this.Defense_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Defense_CB.CheckedState.BorderRadius = 4;
            this.Defense_CB.CheckedState.BorderThickness = 0;
            this.Defense_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Defense_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Defense_CB.Location = new System.Drawing.Point(118, 31);
            this.Defense_CB.Name = "Defense_CB";
            this.Defense_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Defense_CB.Size = new System.Drawing.Size(122, 17);
            this.Defense_CB.TabIndex = 23;
            this.Defense_CB.Text = "موافقة الحماية المدنية";
            this.Defense_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Defense_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.Defense_CB.UncheckedState.BorderRadius = 4;
            this.Defense_CB.UncheckedState.BorderThickness = 0;
            this.Defense_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // Land_CB
            // 
            this.Land_CB.AutoSize = true;
            this.Land_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Land_CB.CheckedState.BorderRadius = 4;
            this.Land_CB.CheckedState.BorderThickness = 0;
            this.Land_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Land_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Land_CB.Location = new System.Drawing.Point(243, 54);
            this.Land_CB.Name = "Land_CB";
            this.Land_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Land_CB.Size = new System.Drawing.Size(77, 17);
            this.Land_CB.TabIndex = 24;
            this.Land_CB.Text = "عقود الأرض";
            this.Land_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Land_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.Land_CB.UncheckedState.BorderRadius = 4;
            this.Land_CB.UncheckedState.BorderThickness = 0;
            this.Land_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // Air_CB
            // 
            this.Air_CB.AutoSize = true;
            this.Air_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Air_CB.CheckedState.BorderRadius = 4;
            this.Air_CB.CheckedState.BorderThickness = 0;
            this.Air_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Air_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Air_CB.Location = new System.Drawing.Point(18, 31);
            this.Air_CB.Name = "Air_CB";
            this.Air_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Air_CB.Size = new System.Drawing.Size(91, 17);
            this.Air_CB.TabIndex = 25;
            this.Air_CB.Text = "الطيران المدني";
            this.Air_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Air_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.Air_CB.UncheckedState.BorderRadius = 4;
            this.Air_CB.UncheckedState.BorderThickness = 0;
            this.Air_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // Study_CB
            // 
            this.Study_CB.AutoSize = true;
            this.Study_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Study_CB.CheckedState.BorderRadius = 4;
            this.Study_CB.CheckedState.BorderThickness = 0;
            this.Study_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Study_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Study_CB.Location = new System.Drawing.Point(11, 54);
            this.Study_CB.Name = "Study_CB";
            this.Study_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Study_CB.Size = new System.Drawing.Size(98, 17);
            this.Study_CB.TabIndex = 26;
            this.Study_CB.Text = "الدراسة المرورية";
            this.Study_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Study_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.Study_CB.UncheckedState.BorderRadius = 4;
            this.Study_CB.UncheckedState.BorderThickness = 0;
            this.Study_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // guna2CircleButton1
            // 
            this.guna2CircleButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.guna2CircleButton1.Animated = true;
            this.guna2CircleButton1.BackColor = System.Drawing.Color.Transparent;
            this.guna2CircleButton1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("guna2CircleButton1.BackgroundImage")));
            this.guna2CircleButton1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.guna2CircleButton1.DisabledState.BorderColor = System.Drawing.Color.DarkGray;
            this.guna2CircleButton1.DisabledState.CustomBorderColor = System.Drawing.Color.DarkGray;
            this.guna2CircleButton1.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(169)))), ((int)(((byte)(169)))));
            this.guna2CircleButton1.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(141)))), ((int)(((byte)(141)))), ((int)(((byte)(141)))));
            this.guna2CircleButton1.FillColor = System.Drawing.Color.Transparent;
            this.guna2CircleButton1.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.guna2CircleButton1.ForeColor = System.Drawing.Color.White;
            this.guna2CircleButton1.IndicateFocus = true;
            this.guna2CircleButton1.Location = new System.Drawing.Point(867, 3);
            this.guna2CircleButton1.Name = "guna2CircleButton1";
            this.guna2CircleButton1.ShadowDecoration.Mode = Guna.UI2.WinForms.Enums.ShadowMode.Circle;
            this.guna2CircleButton1.Size = new System.Drawing.Size(38, 31);
            this.guna2CircleButton1.TabIndex = 27;
            this.guna2CircleButton1.UseTransparentBackground = true;
            this.guna2CircleButton1.Click += new System.EventHandler(this.guna2CircleButton1_Click);
            // 
            // Enviromental_CB
            // 
            this.Enviromental_CB.AutoSize = true;
            this.Enviromental_CB.CheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Enviromental_CB.CheckedState.BorderRadius = 4;
            this.Enviromental_CB.CheckedState.BorderThickness = 0;
            this.Enviromental_CB.CheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.Enviromental_CB.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Enviromental_CB.Location = new System.Drawing.Point(266, 31);
            this.Enviromental_CB.Name = "Enviromental_CB";
            this.Enviromental_CB.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Enviromental_CB.Size = new System.Drawing.Size(110, 17);
            this.Enviromental_CB.TabIndex = 28;
            this.Enviromental_CB.Text = "موافقة وزارة البيئه";
            this.Enviromental_CB.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Enviromental_CB.UncheckedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            this.Enviromental_CB.UncheckedState.BorderRadius = 4;
            this.Enviromental_CB.UncheckedState.BorderThickness = 0;
            this.Enviromental_CB.UncheckedState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(125)))), ((int)(((byte)(137)))), ((int)(((byte)(149)))));
            // 
            // guna2Panel1
            // 
            this.guna2Panel1.BackColor = System.Drawing.Color.White;
            this.guna2Panel1.Controls.Add(this.Enviromental_CB);
            this.guna2Panel1.Controls.Add(this.guna2CircleButton1);
            this.guna2Panel1.Controls.Add(this.Study_CB);
            this.guna2Panel1.Controls.Add(this.Air_CB);
            this.guna2Panel1.Controls.Add(this.Land_CB);
            this.guna2Panel1.Controls.Add(this.Defense_CB);
            this.guna2Panel1.Controls.Add(this.liscence_CB);
            this.guna2Panel1.Controls.Add(this.Petrolum_CB);
            this.guna2Panel1.Controls.Add(this.label3);
            this.guna2Panel1.Controls.Add(this.label2);
            this.guna2Panel1.Controls.Add(this.label1);
            this.guna2Panel1.Controls.Add(this.guna2ComboBox3);
            this.guna2Panel1.Controls.Add(this.guna2ComboBox2);
            this.guna2Panel1.Controls.Add(this.guna2CircleButton2);
            this.guna2Panel1.Controls.Add(this.guna2TextBox1);
            this.guna2Panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.guna2Panel1.Location = new System.Drawing.Point(0, 0);
            this.guna2Panel1.Name = "guna2Panel1";
            this.guna2Panel1.Size = new System.Drawing.Size(908, 84);
            this.guna2Panel1.TabIndex = 1;
            // 
            // guna2TextBox1
            // 
            this.guna2TextBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.guna2TextBox1.Animated = true;
            this.guna2TextBox1.AutoRoundedCorners = true;
            this.guna2TextBox1.BorderRadius = 13;
            this.guna2TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.guna2TextBox1.DefaultText = "";
            this.guna2TextBox1.DisabledState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(208)))), ((int)(((byte)(208)))));
            this.guna2TextBox1.DisabledState.FillColor = System.Drawing.Color.FromArgb(((int)(((byte)(226)))), ((int)(((byte)(226)))), ((int)(((byte)(226)))));
            this.guna2TextBox1.DisabledState.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(138)))), ((int)(((byte)(138)))));
            this.guna2TextBox1.DisabledState.PlaceholderForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(138)))), ((int)(((byte)(138)))), ((int)(((byte)(138)))));
            this.guna2TextBox1.FocusedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.guna2TextBox1.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.guna2TextBox1.HoverState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.guna2TextBox1.Location = new System.Drawing.Point(642, 42);
            this.guna2TextBox1.Name = "guna2TextBox1";
            this.guna2TextBox1.PlaceholderText = "هنا يتم ادخال اسم او رقم للبحث";
            this.guna2TextBox1.SelectedText = "";
            this.guna2TextBox1.Size = new System.Drawing.Size(219, 29);
            this.guna2TextBox1.Style = Guna.UI2.WinForms.Enums.TextBoxStyle.Material;
            this.guna2TextBox1.TabIndex = 14;
            this.guna2TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.guna2TextBox1.TextChanged += new System.EventHandler(this.guna2TextBox1_TextChanged);
            // 
            // guna2ComboBox2
            // 
            this.guna2ComboBox2.AutoRoundedCorners = true;
            this.guna2ComboBox2.BackColor = System.Drawing.Color.Transparent;
            this.guna2ComboBox2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.guna2ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.guna2ComboBox2.FocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.guna2ComboBox2.FocusedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.guna2ComboBox2.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.guna2ComboBox2.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(88)))), ((int)(((byte)(112)))));
            this.guna2ComboBox2.ItemHeight = 30;
            this.guna2ComboBox2.Items.AddRange(new object[] {
            "تم",
            "جاري",
            "الكل"});
            this.guna2ComboBox2.Location = new System.Drawing.Point(405, 35);
            this.guna2ComboBox2.Name = "guna2ComboBox2";
            this.guna2ComboBox2.Size = new System.Drawing.Size(87, 36);
            this.guna2ComboBox2.Style = Guna.UI2.WinForms.Enums.TextBoxStyle.Material;
            this.guna2ComboBox2.TabIndex = 16;
            this.guna2ComboBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // guna2ComboBox3
            // 
            this.guna2ComboBox3.AutoRoundedCorners = true;
            this.guna2ComboBox3.BackColor = System.Drawing.Color.Transparent;
            this.guna2ComboBox3.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.guna2ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.guna2ComboBox3.FocusedColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.guna2ComboBox3.FocusedState.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(94)))), ((int)(((byte)(148)))), ((int)(((byte)(255)))));
            this.guna2ComboBox3.Font = new System.Drawing.Font("Segoe UI", 10F);
            this.guna2ComboBox3.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(68)))), ((int)(((byte)(88)))), ((int)(((byte)(112)))));
            this.guna2ComboBox3.ItemHeight = 30;
            this.guna2ComboBox3.Items.AddRange(new object[] {
            "موجود",
            "x",
            "✔",
            "الكل",
            "x موجود و ",
            "✔ موجود و",
            "x و ✔"});
            this.guna2ComboBox3.Location = new System.Drawing.Point(498, 35);
            this.guna2ComboBox3.Name = "guna2ComboBox3";
            this.guna2ComboBox3.Size = new System.Drawing.Size(138, 36);
            this.guna2ComboBox3.Style = Guna.UI2.WinForms.Enums.TextBoxStyle.Material;
            this.guna2ComboBox3.TabIndex = 17;
            this.guna2ComboBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(908, 450);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.guna2Panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form3";
            this.Text = "Form3";
            this.Load += new System.EventHandler(this.Form3_Load);
            ((System.ComponentModel.ISupportInitialize)(this.stationsBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.database1DataSet)).EndInit();
            this.guna2Panel1.ResumeLayout(false);
            this.guna2Panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private Database1DataSet database1DataSet;
        private System.Windows.Forms.BindingSource stationsBindingSource;
        private Database1DataSetTableAdapters.StationsTableAdapter stationsTableAdapter;
        private Guna.UI2.WinForms.Guna2CircleButton guna2CircleButton2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private Guna.UI2.WinForms.Guna2CheckBox Petrolum_CB;
        private Guna.UI2.WinForms.Guna2CheckBox liscence_CB;
        private Guna.UI2.WinForms.Guna2CheckBox Defense_CB;
        private Guna.UI2.WinForms.Guna2CheckBox Land_CB;
        private Guna.UI2.WinForms.Guna2CheckBox Air_CB;
        private Guna.UI2.WinForms.Guna2CheckBox Study_CB;
        private Guna.UI2.WinForms.Guna2CircleButton guna2CircleButton1;
        private Guna.UI2.WinForms.Guna2CheckBox Enviromental_CB;
        private Guna.UI2.WinForms.Guna2Panel guna2Panel1;
        private Guna.UI2.WinForms.Guna2ComboBox guna2ComboBox3;
        private Guna.UI2.WinForms.Guna2ComboBox guna2ComboBox2;
        private Guna.UI2.WinForms.Guna2TextBox guna2TextBox1;
    }
}