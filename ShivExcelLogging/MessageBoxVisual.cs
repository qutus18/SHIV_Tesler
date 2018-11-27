using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Timers;

namespace ShivExcelLogging
{
    public partial class MessageBoxVisual : Form
    {
        public System.Timers.Timer timerNew;

        public MessageBoxVisual()
        {
            InitializeComponent();
        }

        private void btnCloseForm2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void MessageBoxVisual_Load(object sender, EventArgs e)
        {
            timerNew = new System.Timers.Timer();
            timerNew.Elapsed += new System.Timers.ElapsedEventHandler(OntimeEvent);
            timerNew.Interval = 2000;
            timerNew.Start();
        }

        private void OntimeEvent(object source, ElapsedEventArgs e)
        {
            if (InvokeRequired)
                Invoke(new MethodInvoker(delegate { this.Hide(); }));
            
        }
    }
}
