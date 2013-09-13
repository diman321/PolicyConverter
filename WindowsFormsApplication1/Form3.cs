using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PolycyConverter
{
    public partial class Form3 : Form
    {
        public Form3()
        {

            InitializeComponent();

            if (WindowsFormsApplication1.Form1.CurrentErrorObjResolved.Count != 0)
            {
                foreach (List<string> OneErrorObj in WindowsFormsApplication1.Form1.CurrentErrorObjResolved)
                {
                    this.richTextBox1.AppendText(OneErrorObj[0] + Environment.NewLine);
                }
            }
            else
            {
                richTextBox1.Visible = false;
                label1.Visible = false;
                this.ClientSize = new System.Drawing.Size(310, 398);
            }
        }

        private void Form3_MouseClick(object sender, MouseEventArgs e)
        {
            try
            {
                Form3.ActiveForm.Close();
            }
            catch
            {
                return;
            }
        }

    }
}
