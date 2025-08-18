// MaterialStyleComboBox.cs
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace DemoProject
{
    public class MaterialStyleComboBox : ComboBox
    {
        public MaterialStyleComboBox()
        {
            this.DropDownStyle = ComboBoxStyle.DropDown;
            this.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            this.AutoCompleteSource = AutoCompleteSource.ListItems;

            this.BackColor = Color.White;
            this.ForeColor = Color.FromArgb(68, 88, 112);
            this.Font = new Font("Segoe UI", 10F);
            this.FlatStyle = FlatStyle.Flat;
            this.MaxDropDownItems = 8;
            this.DropDownHeight = 106;
            this.DrawMode = DrawMode.OwnerDrawFixed;
            this.RightToLeft = RightToLeft.Yes;
            this.IntegralHeight = false;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();

            using (Brush brush = new SolidBrush(this.ForeColor))
            {
                e.Graphics.DrawString(this.Items[e.Index].ToString(), this.Font, brush, e.Bounds);
            }

            e.DrawFocusRectangle();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Draw only bottom border
            using (Pen pen = new Pen(Color.FromArgb(217, 221, 226), 1))
            {
                e.Graphics.DrawLine(pen, 0, this.Height - 1, this.Width, this.Height - 1);
            }
        }
    }

}
