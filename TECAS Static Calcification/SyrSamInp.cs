using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TECAS_Static_Calcification
{
    public partial class SyrSamInp : Form
    {
        public SyrSamInp(string Units)
        {
            InitializeComponent();
            label2.Text = Units;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                TECAS.SampleVolume = Convert.ToDouble(textBox1.Text);
            }
            catch (Exception h)
            {
                MessageBox.Show("Please provide numbers only!");
                return;
            }
            TECAS.State = 4;
            this.Hide();
        }
    }
}
