using System;
using System.Collections.Generic;
using System.Drawing;
namespace PolycyConverter
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
            this.label1 = new System.Windows.Forms.Label();
            this.alphaBlendTextBox2 = new ZBobb.AlphaBlendTextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(316, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(147, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Не удалось найти объекты:";

            // 
            // alphaBlendTextBox2
            // 
            this.alphaBlendTextBox2.BackAlpha = 0;
            this.alphaBlendTextBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.alphaBlendTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.alphaBlendTextBox2.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.alphaBlendTextBox2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.alphaBlendTextBox2.Location = new System.Drawing.Point(319, 25);
            this.alphaBlendTextBox2.Multiline = true;
            this.alphaBlendTextBox2.Name = "alphaBlendTextBox2";
            this.alphaBlendTextBox2.ReadOnly = true;
            this.alphaBlendTextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.alphaBlendTextBox2.Size = new System.Drawing.Size(283, 361);
            this.alphaBlendTextBox2.TabIndex = 1;
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LimeGreen;
            this.BackgroundImage = global::PolycyConverter.Properties.Resources.bad;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(614, 398);
            this.ControlBox = false;

            if (WindowsFormsApplication1.Form1.CurrentErrorObjResolved.Count != 0)
            {
                foreach (List<string> OneErrorObj in WindowsFormsApplication1.Form1.CurrentErrorObjResolved)
                {
                    this.alphaBlendTextBox2.AppendText(OneErrorObj[0] + " " + OneErrorObj[1] + " " + OneErrorObj[2] + Environment.NewLine);
                }
                
                this.Controls.Add(this.alphaBlendTextBox2);
                this.Controls.Add(this.label1);
            }

            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "Form3";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form3";
            this.TopMost = true;
            this.TransparencyKey = System.Drawing.Color.LimeGreen;
            this.MouseClick += new System.Windows.Forms.MouseEventHandler(this.Form3_MouseClick);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private ZBobb.AlphaBlendTextBox alphaBlendTextBox2;
    }
}