using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing.Drawing2D;

namespace AprioriAlgorithm
{
    class RoundedButton:Button
    {
	// Kullanım kılavuzunda kullanılmak üzere buton biçimini override eden metod.
	// Butonlar bu metod sayesinde elips şeklinde çizilmektedir.
        protected override void OnPaint(System.Windows.Forms.PaintEventArgs e)
        {
            GraphicsPath grPath = new GraphicsPath();
            grPath.AddEllipse(1, 1, ClientSize.Width - 4, ClientSize.Height - 4);
            this.Region = new System.Drawing.Region(grPath);
            base.OnPaint(e);
        }
    }
}
