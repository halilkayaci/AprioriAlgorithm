using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AprioriAlgorithm
{
    public partial class Help : Form
    {
        public Help()
        {
            InitializeComponent();
        }
        private void Help_Load(object sender, EventArgs e)
        {
	    // Form yüklendiğinde adım 1 resmi gösterilip label rengi değiştiriliyor.
            image_Help.Image = Properties.Resources.Step1;
            lbl_Step1.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step1_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step1;
            lbl_Step1.ForeColor = Color.FromArgb(0, 128, 64);
        }
 
        private void rbtn_Step2_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step2;
            lbl_Step2.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step3_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step3;
            lbl_Step3.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step4_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step4;
            lbl_Step4.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step5_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step5;
            lbl_Step5.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step6_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step6;
            lbl_Step6.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step7_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step7;
            lbl_Step7.ForeColor = Color.FromArgb(0, 128, 64);
        }

        private void rbtn_Step8_Click(object sender, EventArgs e)
        {
            changeAllLabelColor();
            image_Help.Image = Properties.Resources.Step8;
            lbl_Step8.ForeColor = Color.FromArgb(0, 128, 64);
        }
	
	// Tüm label renklerini siyah yapan metod.
        private void changeAllLabelColor() {
            lbl_Step1.ForeColor = Color.Black;
            lbl_Step2.ForeColor = Color.Black;
            lbl_Step3.ForeColor = Color.Black;
            lbl_Step4.ForeColor = Color.Black;
            lbl_Step5.ForeColor = Color.Black;
            lbl_Step6.ForeColor = Color.Black;
            lbl_Step7.ForeColor = Color.Black;
            lbl_Step8.ForeColor = Color.Black;
        }
    }
}
