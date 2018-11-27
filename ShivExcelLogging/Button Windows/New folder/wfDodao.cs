using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShivExcelLogging
{
    public partial class wfDodao : Form
    {
        int countProcess = 0;
        Timer timerProcess = new Timer();
        public wfDodao()
        {
            InitializeComponent();
            timerProcess.Interval = 100;
            timerProcess.Tick += incProcess;
            timerProcess.Start();
        }

        private void incProcess(object sender, EventArgs e)
        {
            countProcess += 1;
            if (countProcess > 8) countProcess = 0;
            lblProcess.Text = "";
            for (int i = 0; i < countProcess; i++)
            {
                lblProcess.Text += "_";
            }
        }

        private void wfKheho_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerProcess.Stop();
        }

        private void lblCloseKheho_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
