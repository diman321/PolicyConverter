
using System.Drawing;
using System.Windows.Forms;
namespace WindowsFormsApplication1
{

    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        /// 

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.alphaBlendTextBox2 = new ZBobb.AlphaBlendTextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.Filter = "HTML Files|*.html";
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.Transparent;
            this.button1.Location = new System.Drawing.Point(13, 13);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(152, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Открыть политику в HTML";
            this.toolTip1.SetToolTip(this.button1, "Чтобы начать - откройте экспортированную в HTML политику МСЭ");
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(13, 43);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(371, 23);
            this.progressBar1.TabIndex = 1;
            this.toolTip1.SetToolTip(this.progressBar1, "Прогресс заполнения таблицы ExCel");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(12, 99);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(174, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Пока ничего не делаю";
            this.toolTip1.SetToolTip(this.label1, "Тут показано текущее действие программы");
            // 
            // progressBar2
            // 
            this.progressBar2.Location = new System.Drawing.Point(13, 72);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(371, 23);
            this.progressBar2.TabIndex = 4;
            this.toolTip1.SetToolTip(this.progressBar2, "Прогресс текущего действия программы");
            // 
            // alphaBlendTextBox2
            // 
            this.alphaBlendTextBox2.BackAlpha = 0;
            this.alphaBlendTextBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.alphaBlendTextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.alphaBlendTextBox2.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.alphaBlendTextBox2.ForeColor = System.Drawing.SystemColors.WindowText;
            this.alphaBlendTextBox2.Location = new System.Drawing.Point(13, 118);
            this.alphaBlendTextBox2.Multiline = true;
            this.alphaBlendTextBox2.Name = "alphaBlendTextBox2";
            this.alphaBlendTextBox2.ReadOnly = true;
            this.alphaBlendTextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.alphaBlendTextBox2.Size = new System.Drawing.Size(371, 150);
            this.alphaBlendTextBox2.TabIndex = 0;
            this.toolTip1.SetToolTip(this.alphaBlendTextBox2, "Тут ведётся лог всех действий программы");
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.Transparent;
            this.button2.Location = new System.Drawing.Point(172, 13);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(42, 23);
            this.button2.TabIndex = 5;
            this.button2.Text = "Стоп";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // checkBox1
            // 
            this.checkBox1.AutoEllipsis = true;
            this.checkBox1.AutoSize = true;
            this.checkBox1.BackColor = System.Drawing.Color.Transparent;
            this.checkBox1.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.checkBox1.Location = new System.Drawing.Point(220, 17);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(129, 17);
            this.checkBox1.TabIndex = 6;
            this.checkBox1.Text = "Удобочитаемый вид";
            this.toolTip1.SetToolTip(this.checkBox1, "Создаёт для групповых объектов с количеством\r\nчленов более 15 отдельный лист");
            this.checkBox1.UseVisualStyleBackColor = false;
            // 
            // Form1
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = global::PolycyConverter.Properties.Resources.заливко41;
            this.ClientSize = new System.Drawing.Size(396, 280);
            this.Controls.Add(this.checkBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.alphaBlendTextBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(412, 314);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(412, 287);
            this.Name = "Form1";
            this.Text = "PolicyConverter";
            this.toolTip1.SetToolTip(this, "Это Программа конвертирования экспортированных в HTML политик МСЭ в ExCel");
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.ToolTip toolTip1;

        private ZBobb.AlphaBlendTextBox alphaBlendTextBox2;
        private Button button2;
        private CheckBox checkBox1;

    }
}

