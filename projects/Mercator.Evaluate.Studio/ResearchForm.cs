using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Mercator.Evaluate.Studio
{
    public partial class ResearchForm : Form
    {
        public double JZZWCL, JZZWCB;
        public double ZDZW1CL, ZDZW1CB;
        public double ZDZW2CL, ZDZW2CB;

        public ResearchForm()
        {
            InitializeComponent();
        }

        private void OKButton_Click(object sender, EventArgs e)
        {
            JZZWCL = Convert.ToDouble(JZZWCLTextBox.Text);
            JZZWCB = Convert.ToDouble(JZZWCBTextBox.Text);
            ZDZW1CL = Convert.ToDouble(ZDZW1CLTextBox.Text);
            ZDZW1CB = Convert.ToDouble(ZDZW1CBTextBox.Text);
            ZDZW2CL = Convert.ToDouble(ZDZW2CLTextBox.Text);
            ZDZW2CB = Convert.ToDouble(ZDZW2CBTextBox.Text);

            this.DialogResult = DialogResult.OK;
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
