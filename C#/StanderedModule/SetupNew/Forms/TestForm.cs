using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SetupNew.Forms
{
    public partial class TestForm : Form
    {
        private System.Threading.Timer timer1;
        private TimerCallback timer1Delegate;
        private AutoResetEvent autoevent1 = new AutoResetEvent(false);
        private ClsPlcSLMP clsPlcSLMP;
        public TestForm()
        {
            InitializeComponent();
        }

        private void TestForm_Load(object sender, EventArgs e)
        {
            timer1Delegate = new TimerCallback(timer1_tick);
            timer1 = new System.Threading.Timer(timer1Delegate, autoevent1, 1000, 1000);
            clsPlcSLMP = new ClsPlcSLMP();
            clsPlcSLMP.SLMPModel = new Models.clsSLMP
            {
                IPAddress = "192.168.1.37",
                PortNo = 1232,
            };
            clsPlcSLMP.Connect();
            for (int i = 0; i < 10000; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.Cells.Add(new DataGridViewTextBoxCell { Value = "D" +  i.ToString() });
                //row.Cells.Add(new DataGridViewTextBoxCell { Value = "0" });
                DGV1.Rows.Add(row);
            }

        }
        private void timer1_tick(object sender)
        {
            timer1.Change(Timeout.Infinite, Timeout.Infinite);
            //for (int i = 0; i < 2000; i++)
            //{
            //    DGV1[1,i].Value = clsPlcSLMP.PLCData[i];
            //}
            timer1.Change(1000,1000);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            clsPlcSLMP.PLCData[int.Parse(txtDataRegister.Text)] = int.Parse(txtData.Text) ;
        }
    }
}
