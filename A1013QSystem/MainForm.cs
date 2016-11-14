using A1013QSystem.Common;
using A1013QSystem.DriverCommon;
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
    public partial class MainForm : Form
    {
        public MainForm()
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

        private void MainForm_Load(object sender, EventArgs e)
        {
            //居中显示
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Fixed3D;
            //显示应用程序在任务栏中的图标
            this.ShowInTaskbar = true;

            //首先要从ini文件读取仪器的参数信息
            CDll.GetAlltheInstumentsParasFromIniFile();
            ipAddressControl.Text = CGloabal.g_InstrPowerModule.ipAdress;
            port.Value = CGloabal.g_InstrPowerModule.port;

            ipAddressControl2.Text = CGloabal.g_InstrPowerModule.ipAdress;
            port2.Value= CGloabal.g_InstrScopeModule.port;

            multiNum.Value = CGloabal.g_InstrMultimeterModule.port;
        }

        private void btnElect_Click(object sender, EventArgs e)
        {
            string strIP;
            UInt32 nPort;
            string resourceName;
            int error = 0;
            if (btnElect.Text == "打开")//用户要连接仪器
            {
                strIP = this.ipAddressControl.Text;
                nPort = (UInt32)this.port.Value;
                //连接设备
                resourceName = "TCPIP0::" + strIP + "::inst0::INSTR";

                error = CDll.ConnectSpecificInstrument(CGloabal.g_curInstrument.strInstruName, resourceName);
                if (error < 0)//连接失败
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, "电源打开失败！");
                    btnElect.Text = "打开";
                }
                else//连接成功,则要将当前用户输入的IP地址和端口号保存到ini文件中
                {
                    CDll.SaveInputNetInforsToIniFile(CGloabal.g_curInstrument.strInstruName, strIP, nPort);
                    btnElect.Text = "关闭";
                }
            }
            else//此时用户要断开连接
            {
                error = CDll.CloseSpecificInstrument(CGloabal.g_curInstrument.strInstruName);
                if (error < 0)//断开失败，则还要将switchConnect恢复为连接状态      
                {
                    btnElect.Text = "关闭";
                }
                else {
                    btnElect.Text = "打开";
                }
            }
        }

        private void btnOscill_Click(object sender, EventArgs e)
        {
            string strIP;
            UInt32 nPort;
            string resourceName;
            int error = 0;
            if (btnOscill.Text == "打开")//用户要连接仪器
            {
                strIP = this.ipAddressControl2.Text;
                nPort = (UInt32)this.port2.Value;
                //连接设备
                resourceName = "TCPIP0::" + strIP + "::inst0::INSTR";

                error = CDll.ConnectSpecificInstrument(CGloabal.g_curInstrument.strInstruName, resourceName);
                if (error < 0)//连接失败
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, "示波器打开失败！");
                    btnOscill.Text = "打开";
                }
                else//连接成功,则要将当前用户输入的IP地址和端口号保存到ini文件中
                {
                    CDll.SaveInputNetInforsToIniFile(CGloabal.g_curInstrument.strInstruName, strIP, nPort);
                    btnOscill.Text = "关闭";
                }
            }
            else//此时用户要断开连接
            {
                error = CDll.CloseSpecificInstrument(CGloabal.g_curInstrument.strInstruName);
                if (error < 0)//断开失败，则还要将switchConnect恢复为连接状态      
                {
                    btnOscill.Text = "关闭";
                }
                else
                {
                    btnOscill.Text = "打开";
                }
            }
        }

        private void btnMulti_Click(object sender, EventArgs e)
        {
            int nGpibAddr = 0;
            int error = 0;
            nGpibAddr = (Int32)multiNum.Value;
            string strResource = "GPIB0::" + nGpibAddr + "::INSTR"; // GPIB1 to GPIB0 changed 16.09.03 by msq
            if (btnMulti.Text == "打开")//用户要连接万用表   //changed 16.09.03 by msq
            {
                error = Multimeter_Driver.Connect_Multimeter(strResource);
                if (error < 0)
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, "万用表连接失败！");
                    btnMulti.Text = "打开";
                }
                else
                {
                    btnMulti.Text = "关闭";
                }
            }
            else//断开万用表
            {
                Multimeter_Driver.Close();
                btnMulti.Text = "打开";
            }
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

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            var tabSelect = (sender as TabControl).SelectedTab;
           if (tabSelect.Text == "试验用例")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage4);
            }
            if (tabSelect.Text == "结果查看") {
                DataTable dt = new DataTable();
                DataRow dr = dt.NewRow();
               
                dt.Columns.Add("编号");
                dt.Columns.Add("时间");
                dt.Columns.Add("内容");
                dt.Columns.Add("结果");
               
                dr["编号"] = "AX";
                dr["时间"] = "AX";
                dr["内容"] = "AX";
                dr["结果"] = "AX";

                dt.Rows.Add(dr);
                dataView.AllowUserToAddRows = false;
                dataView.DataSource = dt;
                dataView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }

            tabControl2.TabPages.Clear();
            tabControl2.TabPages.Add(tabPage3);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            var SelectNode = (sender as TreeView).SelectedNode;
            if (SelectNode.Text == "基本功能测试")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage3);
            }
            else if (SelectNode.Text == "寄存器测试")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage4);
            }           
        }

        private void bunTuple1_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道1";
            bunTuple1.BaseColor = Color.DeepSkyBlue;
            bunTuple2.BaseColor = Color.LightGray;
            bunTuple3.BaseColor = Color.LightGray;
            bunTuple4.BaseColor = Color.LightGray;

        }

        private void bunTuple2_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道2";
            bunTuple1.BaseColor = Color.LightGray;
            bunTuple2.BaseColor = Color.DeepSkyBlue;
            bunTuple3.BaseColor = Color.LightGray;
            bunTuple4.BaseColor = Color.LightGray;
        }

        private void bunTuple3_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道3";
            bunTuple1.BaseColor = Color.LightGray;
            bunTuple2.BaseColor = Color.LightGray;
            bunTuple3.BaseColor = Color.DeepSkyBlue;
            bunTuple4.BaseColor = Color.LightGray;
        }

        private void bunTuple4_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道4";
            bunTuple1.BaseColor = Color.LightGray;
            bunTuple2.BaseColor = Color.LightGray;
            bunTuple3.BaseColor = Color.LightGray;
            bunTuple4.BaseColor = Color.DeepSkyBlue;
        }

      
    }
}
