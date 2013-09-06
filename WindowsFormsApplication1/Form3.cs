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
        }

        private void Form3_MouseClick(object sender, MouseEventArgs e)
        {
            /*if (WindowsFormsApplication1.Form1.tbad.ThreadState == System.Threading.ThreadState.Aborted ||
                WindowsFormsApplication1.Form1.tbad.ThreadState == System.Threading.ThreadState.Stopped ||
                WindowsFormsApplication1.Form1.tbad.ThreadState == System.Threading.ThreadState.StopRequested)
                return;
            try
            {
                WindowsFormsApplication1.Form1.tbad.Abort();
            }
            catch
            {
                return;
            }*/
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
