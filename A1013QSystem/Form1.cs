using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace A1013QSystem
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //将数字键上下箭头去掉
            speedChip1.UpDownButton.Visible = false;
            speedChip2.UpDownButton.Visible = false;

            //           

            //显示时间
            Timer time = new Timer();
            time.Interval = 1000;
            time.Tick += Fn_ShowTime;
            time.Enabled = true;
        }

        public void Fn_ShowTime(object sender,EventArgs e)
        {
            dateStatusLabel.Text = "当前时间：" + System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void speedChip1_ValueChanged(object sender, EventArgs e)
        {
            var speedValue = (sender as NumericUpDown).Value;
            label14.Text = label15.Text = label16.Text = label17.Text = speedValue + "M/s";
        }

        private void speedChip2_ValueChanged(object sender, EventArgs e)
        {
            var speedValue = (sender as NumericUpDown).Value;
            label34.Text = label35.Text = label36.Text = label37.Text = speedValue + "M/s";
        }      

        private void reciveChip1_CheckedChanged(object sender, EventArgs e)
        {
            var signRS = (sender as RadioButton).Checked;
            if (signRS)
            {
                label19.Text = label20.Text = label21.Text = label22.Text = "收";
            }          
        }

        private void sendChip1_CheckedChanged(object sender, EventArgs e)
        {
            var signRS = (sender as RadioButton).Checked;
            if (signRS)
            {
                label19.Text = label20.Text = label21.Text = label22.Text = "发";
            }
        }

        private void reciveChip2_CheckedChanged(object sender, EventArgs e)
        {
            var signRS = (sender as RadioButton).Checked;
            if (signRS)
            {
                label39.Text = label40.Text = label41.Text = label42.Text = "收";
            }
        }
        private void sendChip2_CheckedChanged(object sender, EventArgs e)
        {
            var signRS = (sender as RadioButton).Checked;
            if (signRS)
            {
                label39.Text = label40.Text = label41.Text = label42.Text = "发";
            }
        }
    }
}
