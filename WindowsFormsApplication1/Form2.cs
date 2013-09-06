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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_MouseClick(object sender, MouseEventArgs e)
        {
            /*if (WindowsFormsApplication1.Form1.t.ThreadState == System.Threading.ThreadState.Aborted ||
                WindowsFormsApplication1.Form1.t.ThreadState == System.Threading.ThreadState.Stopped ||
                WindowsFormsApplication1.Form1.t.ThreadState == System.Threading.ThreadState.StopRequested)
                return;
            try
            {
                WindowsFormsApplication1.Form1.t.Abort();
            }
            catch 
            {
                return;
            }*/
            try
            {
                Form2.ActiveForm.Close();
            }
            catch
            {
                return;
            }
        }

    }
}
