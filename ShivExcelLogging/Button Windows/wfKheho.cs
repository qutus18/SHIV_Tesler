using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Ports;
using ActUtlTypeLib;

namespace ShivExcelLogging
{
    public partial class wfKheho : Form
    {
        public delegate void StringCap(string x);
        public event StringCap stringDoneKheho;
        int countProcess = 0;
        Timer timerProcess = new Timer();
        int countAllowEnd = 0;
        string _receiveString = "0.00";
        SerialPort ComKheho;
        private static float valueMax, valueMin;
        ActUtlType plcRef;
        private string bufferString;

        public wfKheho()
        {
            InitializeComponent();
            timerProcess.Interval = 100;
            timerProcess.Tick += incProcess;
            timerProcess.Start();

            this.KeyPreview = true;
            this.KeyDown += CheckKeydown;

            // Giá trị mặc định
            valueMax = (float)0.0001;
            valueMin = 0;
        }

        private void CheckKeydown(object sender, KeyEventArgs e)
        {
            switch (e.KeyData)
            {
                case Keys.Space:
                    if (stringDoneKheho != null) stringDoneKheho((valueMax - valueMin).ToString("0.00"));
                    this.Close();
                    break;
                default:
                    break;
            }
        }

        public wfKheho(ref ActUtlType PLC, ref SerialPort COM) : this()
        {
            plcRef = PLC;

            ComKheho = COM;
            ComKheho.DataReceived -= ProcessComMessage;
            ComKheho.DataReceived += ProcessComMessage;

            // Sent CMD Set Value Message
            Task.Delay(100);
            ComKheho.WriteLine("PRE +0\r\n");
            ComKheho.Write("OUT1\r\n");
        }

        private void ProcessComMessage(object sender, SerialDataReceivedEventArgs e)
        {
            // Chuyển đổi từ String thành Float 2 chữ số sau dấu phẩy
            bufferString += ComKheho.ReadExisting();
            if (bufferString.IndexOf("\r") > 0)
            {
                string tempStringReceive = bufferString;
                bufferString = "";
                Console.WriteLine(tempStringReceive + "--------");
                try
                {
                    float tempF = float.Parse(tempStringReceive);
                    if (valueMax < tempF) valueMax = tempF;
                    if (valueMin > tempF) valueMin = tempF;
                    Invoke(new MethodInvoker(delegate { lblKheho.Text = (valueMax - valueMin).ToString("0.000"); }));
                }
                catch { }
            }
        }

        private void incProcess(object sender, EventArgs e)
        {
            countProcess += 1;
            countAllowEnd += 1;
            if (countProcess > 8) countProcess = 0;
            lblProcess.Text = "";
            for (int i = 0; i < countProcess; i++)
            {
                lblProcess.Text += "_";
            }

            // Check Button ReClick
            if ((plcRef != null) && (countAllowEnd > 20))
            {
                int buttonRead;
                var iret = plcRef.GetDevice("M5", out buttonRead);
                if (buttonRead == 1)
                {
                    if (stringDoneKheho != null) stringDoneKheho((valueMax - valueMin).ToString("0.00"));
                    this.Close();
                }
            }
        }

        private void wfKheho_FormClosing(object sender, FormClosingEventArgs e)
        {
            timerProcess.Stop();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            ComKheho.Write("?\r\n");
        }

        private void lblCloseKheho_Click(object sender, EventArgs e)
        {
            if (stringDoneKheho != null) stringDoneKheho((valueMax - valueMin).ToString("0.00"));
            this.Close();
        }
    }
}
