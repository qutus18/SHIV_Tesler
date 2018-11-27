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
        public delegate void StringCap(string x);
        public event StringCap stringDoneDodao;
        int countProcess = 0;
        Timer timerProcess = new Timer();
        Timer timerProcessClose = new Timer();
        public wfDodao()
        {
            InitializeComponent();
            timerProcess.Interval = 100;
            timerProcess.Tick += incProcess;
            timerProcess.Start();

            timerProcessClose.Interval = 200;
            timerProcessClose.Tick += checkCloseForm;
        }

        private void checkCloseForm(object sender, EventArgs e)
        {
            if (textBox1.Text.IndexOf("DT100") >= 0)
            {
                if (stringDoneDodao != null) stringDoneDodao(textBox1.Text);
                Task.Delay(100);
                timerProcessClose.Stop();
                this.Close();
            }
            else
            {
                textBox1.Text = "DT10000-000004.0345M" + "DT10002-000004.1245M" + "DT10001-000004.4345M";
            }

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

        private void wfDodao_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerProcess.Stop();
        }

        private void lblCloseDodao_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 15) timerProcessClose.Start();
        }
    }
}
