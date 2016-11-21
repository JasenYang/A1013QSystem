using A1013QSystem.Common;
using A1013QSystem.DriverCommon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace A1013QSystem
{
    public partial class MainForm : Form
    {
        public System.IntPtr nHandle = (System.IntPtr)0;
        public MainForm()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;

            //将数字键上下箭头去掉
            speedChip1.UpDownButton.Visible = false;
            speedChip2.UpDownButton.Visible = false;                        
            //           

            //显示时间
            Timer time = new Timer();
            time.Interval = 1000;
            time.Tick += Fn_ShowTime;
            time.Enabled = true;

            //
            //显示串口的默认参数
            //首先检测本机是否含有串口
            string[] str = SerialPort.GetPortNames();
            if (str.Length == 0)
            {
                MessageBox.Show("本机没有串口！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            CGloabal.g_serialPorForUUT = new SerialPort();
            //将或的串口在cmbCom中显示
            foreach (string s in System.IO.Ports.SerialPort.GetPortNames())
            {//获取有多少个COM口  
                cmbCom.Items.Add(s);
            }
            //串口设置默认选择项  
            cmbCom.SelectedIndex = 0;

            //设置其他参数的默认值
            cmbBoundrate.SelectedIndex = 3;
            cmbDataBit.SelectedIndex = 0;
            cmbStopBit.SelectedIndex = 0;
            cmbEvenBit.SelectedIndex = 0;

            CGloabal.CurSerialPortFlag = false; //告诉系统当前的串口配置为板卡的串口配置  
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

        //UUT的串口收到数据回调函数
        private List<byte> buffer = new List<byte>(4096);     
        private void g_serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {           
            int length;
            byte[] ReceiveBytes = new byte[24];           

            CGloabal.g_bIsComRecvedDataFlag = true;  //串口是否收到数据
            length = CGloabal.g_serialPorForUUT.BytesToRead;                 

            byte[] Receivebuf = new byte[length];                             
            CGloabal.ReadCom(CGloabal.g_serialPorForUUT, Receivebuf, length);    
            //1、缓存数据
            buffer.AddRange(Receivebuf);                                        
            //2、完整性判断
            while (buffer.Count >= 24)
            {
                if (buffer[0] == 0xAA)
                {
                    //得到完整的数据，复制到ReceiveBytes中进行校验
                    buffer.CopyTo(0, ReceiveBytes, 0, 24);
                    buffer.RemoveRange(0, 24);

                    /////执行其他代码，对数据进行处理。
                    CGloabal.g_ComReadBuf = ReceiveBytes;

                    int error = 0;
                    int nRtnID = 0;
                    error = CmdsAnalysis(CGloabal.g_ComReadBuf, ref nRtnID);
                }
                else //帧头不正确时，记得清除
                {
                    buffer.RemoveAt(0);
                }
            }
        }
        int tupSend1 = 0; int tupSend2 = 0;int tupSend3 = 0;int tupSend4 =0;
        int tupReceive1 = 0; int tupReceive2 = 0; int tupReceive3 = 0; int tupReceive4 = 0;
        int tupError11 = 0; int tupError12 = 0; int tupError13 = 0; int tupError14 = 0;       
        int tupError21 = 0; int tupError22 = 0; int tupError23 = 0; int tupError24 = 0;

        private int  CmdsAnalysis(byte[] CmdBuf, ref int nCmdID)
        {
            int nTmpVal;
            int nArrayIndex = 0;
            string strErrMsg = "";

            if (CmdBuf[0] != 0xAA)
            {
                return -1;
            }
            if (CmdBuf[23] != 0xBB)
            {
                return -1;
            }

            //波特率
            if (CmdBuf[2] == 104) {
                label14.Text = label15.Text = label16.Text = label17.Text="9600Bps";
                label34.Text = label35.Text = label36.Text = label37.Text = "9600Bps";
            }
            if (CmdBuf[2] == 2)
            {
                label14.Text = label15.Text = label16.Text = label17.Text = "500KBps";
                label34.Text = label35.Text = label36.Text = label37.Text = "500Bps";
            }
            //通道1
            tupSend1 += CmdBuf[6];
            tupError11 += CmdBuf[7];
            tupReceive1 += CmdBuf[8];
            tupError21 += CmdBuf[9];

            label24.Text = label39.Text = tupSend1.ToString();
            label70.Text = tupError11.ToString();

            label19.Text = label44.Text = tupReceive1.ToString();
            label66.Text = tupError21.ToString();


            //通道2
            tupSend2 += CmdBuf[10];
            tupError12 += CmdBuf[11];
            tupReceive2 += CmdBuf[12];
            tupError22 += CmdBuf[13];

            label20.Text = label45.Text = tupSend2.ToString();
            label67.Text = tupError12.ToString();

            label25.Text = label40.Text = tupReceive2.ToString();
            label71.Text = tupError22.ToString();
            //通道3
            tupSend3 += CmdBuf[10];
            tupError13 += CmdBuf[11];
            tupReceive3 += CmdBuf[12];
            tupError23 += CmdBuf[13];

            label26.Text = label41.Text = tupSend3.ToString();
            label72.Text = tupError13.ToString();

            label21.Text = label46.Text = tupReceive3.ToString();
            label68.Text = tupError23.ToString();
            //通道4
            tupSend4 += CmdBuf[12];
            tupError14 += CmdBuf[13];
            tupReceive4 += CmdBuf[12];
            tupError24 += CmdBuf[13];

            label22.Text = label42.Text = tupSend4.ToString();
            label69.Text = tupError14.ToString();

             label27.Text = label47.Text = tupReceive4.ToString();
            label73.Text = tupError24.ToString();

            return 0;
        }

        private void btnTest_Click(object sender, EventArgs e)
        {
          
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

        private void btnSerial_Click(object sender, EventArgs e)
        {
            if (CGloabal.g_serialPorForUUT.IsOpen == false)
            {
                OpenCom(ref CGloabal.g_serialPorForUUT);
                if (CGloabal.g_serialPorForUUT.IsOpen == true)
                {//注册回掉函数    

                    btnSerial.Text = "关闭";
                    CGloabal.g_serialPorForUUT.ReceivedBytesThreshold = CGloabal.nCOM_RECV_NUMS;
                  
                    CGloabal.g_serialPorForUUT.DataReceived += new SerialDataReceivedEventHandler(g_serialPort_DataReceived); //打开串口后开始接收数据
                }
                else
                {
                    MessageBox.Show("串口打开失败!");
                }
            }
            else
            {
                CGloabal.g_serialPorForUUT.Close();
                if (CGloabal.g_serialPorForUUT.IsOpen == false)
                {
                    btnSerial.Text = "打开";
                }
                else
                {
                    MessageBox.Show("串口关闭失败!");
                }
            }
        }

        //连接板卡
        private int OpenCom(ref SerialPort comPort)
        {
            //如果串口已经打开，首先要将其关闭
            if (comPort.IsOpen == true)
            {
                comPort.Close();
            }

            string serialName = cmbCom.SelectedItem.ToString();
            comPort.PortName = serialName;

            //波特率
            comPort.BaudRate = Convert.ToInt32(cmbBoundrate.Text);
            //数据位
            comPort.DataBits = Convert.ToInt32(cmbDataBit.Text);
            //停止位
            switch (cmbStopBit.Text)
            {
                case "1":
                    comPort.StopBits = StopBits.One;
                    break;
                case "2":
                    comPort.StopBits = StopBits.Two;
                    break;
                case "1.5":
                    comPort.StopBits = StopBits.OnePointFive;
                    break;
                default:
                    MessageBox.Show("参数不正确!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    break;

            }
            //奇偶位
            switch (cmbEvenBit.Text)
            {
                case "无":
                    comPort.Parity = Parity.None;
                    break;
                case "奇校验":
                    comPort.Parity = Parity.Odd;
                    break;
                case "偶校验":
                    comPort.Parity = Parity.Even;
                    break;
                default:
                    break;
            }

            comPort.Open();     //打开串口 

            if (comPort.IsOpen != true)
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.hint, "串口打开失败！");
            }         

            return 0;
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
            //var signRS = (sender as RadioButton).Checked;
            //if (signRS)
            //{
            //    label19.Text = label20.Text = label21.Text = label22.Text = "收";
            //}          
        }

        private void sendChip1_CheckedChanged(object sender, EventArgs e)
        {
            //var signRS = (sender as RadioButton).Checked;
            //if (signRS)
            //{
            //    label19.Text = label20.Text = label21.Text = label22.Text = "发";
            //}
        }

        private void reciveChip2_CheckedChanged(object sender, EventArgs e)
        {
            //var signRS = (sender as RadioButton).Checked;
            //if (signRS)
            //{
            //    label39.Text = label40.Text = label41.Text = label42.Text = "收";
            //}
        }
        private void sendChip2_CheckedChanged(object sender, EventArgs e)
        {
            //var signRS = (sender as RadioButton).Checked;
            //if (signRS)
            //{
            //    label39.Text = label40.Text = label41.Text = label42.Text = "发";
            //}
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
            bunTuple1.BaseColor = Color.CornflowerBlue;
            bunTuple2.BaseColor = Color.LightGray;
            bunTuple3.BaseColor = Color.LightGray;
            bunTuple4.BaseColor = Color.LightGray;

        }

        private void bunTuple2_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道2";
            bunTuple1.BaseColor = Color.LightGray;
            bunTuple2.BaseColor = Color.CornflowerBlue;
            bunTuple3.BaseColor = Color.LightGray;
            bunTuple4.BaseColor = Color.LightGray;
        }

        private void bunTuple3_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道3";
            bunTuple1.BaseColor = Color.LightGray;
            bunTuple2.BaseColor = Color.LightGray;
            bunTuple3.BaseColor = Color.CornflowerBlue;
            bunTuple4.BaseColor = Color.LightGray;
        }

        private void bunTuple4_Click(object sender, EventArgs e)
        {
            CGloabal.g_curTupple = "通道4";
            bunTuple1.BaseColor = Color.LightGray;
            bunTuple2.BaseColor = Color.LightGray;
            bunTuple3.BaseColor = Color.LightGray;
            bunTuple4.BaseColor = Color.CornflowerBlue;
        }

       
    }
}
