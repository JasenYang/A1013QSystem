using A1013QSystem.Common;
using A1013QSystem.DriverCommon;
using A1013QSystem.Model;
using DigitalCircuitSystem.DriverDAL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace A1013QSystem
{
    public partial class MainForm : Form
    {       
        public RecordModel RModel = new RecordModel();

        public System.IntPtr nHandle = (System.IntPtr)0;
        DataTable dt = new DataTable();
        public MainForm()
        {
            InitializeComponent();
            Control.CheckForIllegalCrossThreadCalls = false;
            
            GraphicsPath buttonPath = new GraphicsPath();
            buttonPath.AddEllipse(5, 5, 20, 20);

            btnLight1.Region = new System.Drawing.Region(buttonPath);
            btnLight2.Region = new System.Drawing.Region(buttonPath);
            btnLight3.Region = new System.Drawing.Region(buttonPath);
            btnLight4.Region = new System.Drawing.Region(buttonPath);
            btnLight5.Region = new System.Drawing.Region(buttonPath);
            btnLight6.Region = new System.Drawing.Region(buttonPath);
            btnLight7.Region = new System.Drawing.Region(buttonPath);
            btnLight8.Region = new System.Drawing.Region(buttonPath);
           

            //将数字键上下箭头去掉
            volNum1.UpDownButton.Visible = false;
            eleNum1.UpDownButton.Visible = false;
            volNum2.UpDownButton.Visible = false;
            eleNum2.UpDownButton.Visible = false;
            volNum3.UpDownButton.Visible = false;
            eleNum3.UpDownButton.Visible = false;
            volNum4.UpDownButton.Visible = false;
            eleNum4.UpDownButton.Visible = false;

            speedChip1.UpDownButton.Visible = false;
            speedChip2.UpDownButton.Visible = false;                        
            //           

            //显示时间
            Timer time = new Timer();
            time.Interval = 1000;
            time.Tick += Fn_ShowTime;
            time.Enabled = true;

            //读取电压、电流
            Timer tSave = new Timer();
            tSave.Interval = 2000;
            tSave.Tick += Fn_ReadVolAndEle;
            tSave.Enabled = true;


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

          
            dt.Columns.Add("时间"); dt.Columns.Add("电压1"); dt.Columns.Add("电流1"); dt.Columns.Add("电压2"); dt.Columns.Add("电流2"); dt.Columns.Add("速率");
            dt.Columns.Add("芯片1通道1发"); dt.Columns.Add("芯片1通道1收"); dt.Columns.Add("芯片1通道1错误"); dt.Columns.Add("芯片2通道1发"); dt.Columns.Add("芯片2通道1收"); dt.Columns.Add("芯片2通道1错误");
            dt.Columns.Add("芯片1通道2发"); dt.Columns.Add("芯片1通道2收"); dt.Columns.Add("芯片1通道2错误"); dt.Columns.Add("芯片2通道2发"); dt.Columns.Add("芯片2通道2收"); dt.Columns.Add("芯片2通道2错误");
            dt.Columns.Add("芯片1通道3发"); dt.Columns.Add("芯片1通道3收"); dt.Columns.Add("芯片1通道3错误"); dt.Columns.Add("芯片2通道3发"); dt.Columns.Add("芯片2通道3收"); dt.Columns.Add("芯片2通道3错误");
            dt.Columns.Add("芯片1通道4发"); dt.Columns.Add("芯片1通道4收"); dt.Columns.Add("芯片1通道4错误"); dt.Columns.Add("芯片2通道4发"); dt.Columns.Add("芯片2通道4收"); dt.Columns.Add("芯片2通道4错误");

            //tabControl1.TabPages.Clear();
            //tabControl1.TabPages.Add(tabPage1);
            //tabControl1.TabPages.Add(tabPage2);


            //变化电压3
            Timer tVolChange = new Timer();
            tVolChange.Interval = 300000;
            tVolChange.Tick += Fn_CircleVol;
            tVolChange.Enabled = true;

            //变化电压4
            Timer tEleChange = new Timer();
            tEleChange.Interval =  1200000;
            tEleChange.Tick += Fn_CircleEle;
            tEleChange.Enabled = true;
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
        int cireVolNum = 4;
        public void Fn_CircleVol(object sender, EventArgs e)
        {
            switch (cireVolNum)
            {
                case 4:
                    volNum3.Value = (decimal)3.3;
                    cireVolNum--;
                    break;
                case 3:
                    volNum3.Value = (decimal)3.0;
                    cireVolNum--;
                    break;
                case 2:
                    volNum3.Value = (decimal)3.3;
                    cireVolNum--;
                    break;
                case 1:
                    volNum3.Value = (decimal)3.6;
                    cireVolNum = 4;
                    break;
            }
            int error = Power_Driver.SetOutputVoltage(CGloabal.g_InstrPowerModule.nHandle, 3, (double)volNum3.Value, "");
        }

        int cireEleNum = 4;
        public void Fn_CircleEle(object sender, EventArgs e)
        {
            switch (cireEleNum)
            {
                case 4:
                    volNum4.Value = (decimal)1.2;
                    cireEleNum--;
                    break;
                case 3:
                    volNum4.Value = (decimal)1.08;
                    cireEleNum--;
                    break;
                case 2:
                    volNum4.Value = (decimal)1.2;
                    cireEleNum--;
                    break;
                case 1:
                    volNum4.Value = (decimal)1.32;
                    cireEleNum = 4;
                    break;
            }
            int error = Power_Driver.SetOutputVoltage(CGloabal.g_InstrPowerModule.nHandle, 4, (double)volNum4.Value, "");
        }

        //UUT的串口收到数据回调函数
        private List<byte> buffer = new List<byte>(4096);     
        private void g_serialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            RModel.curDate = System.DateTime.Now;
            RModel.volVal1 =double.Parse(string.IsNullOrEmpty(vol1.Text)?"0": vol1.Text);
            RModel.eleVal1 = double.Parse(string.IsNullOrEmpty(ele1.Text) ? "0" : ele1.Text);
            RModel.volVal2 = double.Parse(string.IsNullOrEmpty(vol2.Text) ? "0" : vol2.Text);
            RModel.eleVal2 = double.Parse(string.IsNullOrEmpty(ele2.Text) ? "0" : ele2.Text);

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
                RModel.rate = 9600;
                label14.Text = label15.Text = label16.Text = label17.Text="9600bps";
                label34.Text = label35.Text = label36.Text = label37.Text = "9600bps";
            }
            if (CmdBuf[2] == 2)
            {
                RModel.rate = 500000;
                label14.Text = label15.Text = label16.Text = label17.Text = "500Kbps";
                label34.Text = label35.Text = label36.Text = label37.Text = "500Kbps";
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
            //数据保存
            RModel.sendVal11 = RModel.receivVal21 = CmdBuf[6];
            RModel.errorVal11 = CmdBuf[7];
            RModel.receivVal11 = RModel.sendVal21 = CmdBuf[8];
            RModel.errorVal21 = CmdBuf[9];

            //通道2
            tupSend2 += CmdBuf[10];
            tupError12 += CmdBuf[11];
            tupReceive2 += CmdBuf[12];
            tupError22 += CmdBuf[13];

            label25.Text = label40.Text = tupSend2.ToString();
            label71.Text = tupError12.ToString();

            label20.Text = label45.Text = tupReceive2.ToString();
            label67.Text = tupError22.ToString();

            //数据保存
            RModel.sendVal12 = RModel.receivVal22 = CmdBuf[10];
            RModel.errorVal12 = CmdBuf[11];
            RModel.receivVal12 = RModel.sendVal22 = CmdBuf[12];
            RModel.errorVal22 = CmdBuf[13];
            //通道3
            tupSend3 += CmdBuf[14];
            tupError13 += CmdBuf[15];
            tupReceive3 += CmdBuf[16];
            tupError23 += CmdBuf[17];

            label26.Text = label41.Text = tupSend3.ToString();
            label72.Text = tupError13.ToString();
            label21.Text = label46.Text = tupReceive3.ToString();
            label68.Text = tupError23.ToString();

            //数据保存
            RModel.sendVal13 = RModel.receivVal23 = CmdBuf[14];
            RModel.errorVal13 = CmdBuf[15];
            RModel.receivVal13 = RModel.sendVal23 = CmdBuf[16];
            RModel.errorVal23 = CmdBuf[17];
            //通道4
            tupSend4 += CmdBuf[18];
            tupError14 += CmdBuf[19];
            tupReceive4 += CmdBuf[20];
            tupError24 += CmdBuf[21];
            
            label27.Text = label47.Text =  tupSend4.ToString();
            label73.Text = tupError24.ToString();
            label22.Text = label42.Text = tupReceive4.ToString();
            label69.Text = tupError14.ToString();
            //数据保存
            RModel.sendVal14 = RModel.receivVal24 = CmdBuf[18];
            RModel.errorVal14 = CmdBuf[19];
            RModel.receivVal14 = RModel.sendVal24 = CmdBuf[20];
            RModel.errorVal24 = CmdBuf[21];

            CGloabal.LModel.Add(RModel);
                      
            DataRow dr = dt.NewRow();            
            foreach (var item in CGloabal.LModel)
            {
                dr["时间"] = item.curDate;dr["电压1"] = item.volVal1; dr["电流1"] = item.eleVal1;dr["电压2"] = item.volVal2;  dr["电流2"] = item.eleVal2; dr["速率"] = item.rate;
                dr["芯片1通道1发"] = item.sendVal11; dr["芯片1通道1收"] = item.receivVal11; dr["芯片1通道1错误"] = item.errorVal11;
                dr["芯片2通道1发"] = item.sendVal21; dr["芯片2通道1收"] = item.receivVal21; dr["芯片2通道1错误"] = item.errorVal21;

                dr["芯片1通道2发"] = item.sendVal12; dr["芯片1通道2收"] = item.receivVal12; dr["芯片1通道2错误"] = item.errorVal12;
                dr["芯片2通道2发"] = item.sendVal22; dr["芯片2通道2收"] = item.receivVal22; dr["芯片2通道2错误"] = item.errorVal22;

                dr["芯片1通道3发"] = item.sendVal13; dr["芯片1通道3收"] = item.receivVal13; dr["芯片1通道3错误"] = item.errorVal13;
                dr["芯片2通道3发"] = item.sendVal23; dr["芯片2通道3收"] = item.receivVal23; dr["芯片2通道3错误"] = item.errorVal23;

                dr["芯片1通道4发"] = item.sendVal14; dr["芯片1通道4收"] = item.receivVal14; dr["芯片1通道4错误"] = item.errorVal14;
                dr["芯片2通道4发"] = item.sendVal24; dr["芯片2通道4收"] = item.receivVal24; dr["芯片2通道4错误"] = item.errorVal24;
            }

            dt.Rows.Add(dr);
            // dt.Select("1=1", " 时间 desc");
            if (dt.Rows.Count == 22)
            {
                dt.Rows[0].Delete();
            }
            //dataView.DataSource = dt;
            ////dataView.Sort(dataView.Columns[0], ListSortDirection.Descending);

            //dataView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dataView.AllowUserToAddRows = false;    
            dataView.ScrollBars = ScrollBars.Both;

            dataView.DataSource = dt;// CGloabal.LModel;     
            dataView.Columns[0].Width = 160;

            //数据保存到Excel中
            int kk = CDll.DataSaveExcel(CGloabal.LModel);
            if (kk > 0) {
                CGloabal.LModel.Clear();
            }
            return 0;
        }

        public void Fn_ReadVolAndEle(object sender,EventArgs e) {
            string strErrMsg = "";
            int error;
            double voltage1=0,eletage1=0;
            double voltage2 = 0, eletage2 = 0;
            if (btnElect.Text=="关闭")
            {
                error = Power_Driver.ReadVoltage(CGloabal.g_InstrPowerModule.nHandle, 3, strErrMsg, ref voltage1);
                error = Power_Driver.ReadCurrent(CGloabal.g_InstrPowerModule.nHandle, 3, strErrMsg, ref eletage1);
                vol1.Text = voltage1.ToString("0.00000");
                ele1.Text = eletage1.ToString("0.00000");

                error = Power_Driver.ReadVoltage(CGloabal.g_InstrPowerModule.nHandle, 4, strErrMsg, ref voltage2);
                error = Power_Driver.ReadCurrent(CGloabal.g_InstrPowerModule.nHandle, 4, strErrMsg, ref eletage2);

                vol2.Text = voltage2.ToString("0.00000");
                ele2.Text = eletage2.ToString("0.00000");
            }
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

                error = CDll.ConnectSpecificInstrument(CGloabal.g_InstrPowerModule.strInstruName, resourceName);
                if (error < 0)//连接失败
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, "电源打开失败！");
                    btnElect.Text = "打开";
                }
                else//连接成功,则要将当前用户输入的IP地址和端口号保存到ini文件中
                {
                    CDll.SaveInputNetInforsToIniFile(CGloabal.g_InstrPowerModule.strInstruName, strIP, nPort);
                    btnElect.Text = "关闭";
                }
            }
            else//此时用户要断开连接
            {
                error = CDll.CloseSpecificInstrument(CGloabal.g_InstrPowerModule.strInstruName);
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

                error = CDll.ConnectSpecificInstrument(CGloabal.g_InstrScopeModule.strInstruName, resourceName);
                if (error < 0)//连接失败
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, "示波器打开失败！");
                    btnOscill.Text = "打开";
                }
                else//连接成功,则要将当前用户输入的IP地址和端口号保存到ini文件中
                {
                    CDll.SaveInputNetInforsToIniFile(CGloabal.g_InstrScopeModule.strInstruName, strIP, nPort);
                    btnOscill.Text = "关闭";
                }
            }
            else//此时用户要断开连接
            {
                error = CDll.CloseSpecificInstrument(CGloabal.g_InstrScopeModule.strInstruName);
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
                    //CGloabal.g_serialPorForUUT.ReceivedBytesThreshold = CGloabal.nCOM_RECV_NUMS;                  
                    //CGloabal.g_serialPorForUUT.DataReceived += new SerialDataReceivedEventHandler(g_serialPort_DataReceived); //打开串口后开始接收数据
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
                tabControl2.TabPages.Add(tabPage3);                
            }
            if (tabSelect.Text == "结果查看") {
                DataTable dt = new DataTable();              

            //    dt.Columns.Add("时间");
            //    dt.Columns.Add("电压1");
            //    dt.Columns.Add("电流1");
            //    dt.Columns.Add("电压2");
            //    dt.Columns.Add("电流2");
            //    dt.Columns.Add("速率");
            //    dt.Columns.Add("11"); dt.Columns.Add("12电流2电流2"); dt.Columns.Add("13"); dt.Columns.Add("14");
            //    dt.Columns.Add("21"); dt.Columns.Add("22电流2电流2"); dt.Columns.Add("23"); dt.Columns.Add("24");
            //    dt.Columns.Add("31"); dt.Columns.Add("32电流2电流2"); dt.Columns.Add("33"); dt.Columns.Add("34");
            //    dt.Columns.Add("41"); dt.Columns.Add("42电流2电流2"); dt.Columns.Add("43"); dt.Columns.Add("44");
            //    dt.Columns.Add("51"); dt.Columns.Add("52电流2电流2"); dt.Columns.Add("53"); dt.Columns.Add("54");

            //    for (int i=0;i<10;i++)
            //    {
            //        DataRow dr = dt.NewRow();
            //        dr["时间"] = "2016/11/22 12:00:"+i;
            //        dr["电压1"] = "3233";
            //        dr["电流1"] = "AX"+i;
            //        dr["电压2"] = "AX";
            //        dr["电流2"] = "AX";
            //        dr["速率"] = "AX";
            //        dr["11"] = "34000"; dr["12电流2电流2"] = "34000"; dr["13"] = "34000"; dr["14"] = "34000";
            //        dr["21"] = "34000"; dr["22电流2电流2"] = "34000"; dr["23"] = "34000"; dr["24"] = "34000";
            //        dr["31"] = "34000"; dr["32电流2电流2"] = "34000"; dr["33"] = "34000"; dr["34"] = "34000";
            //        dr["41"] = "34000"; dr["42电流2电流2"] = "34000"; dr["43"] = "34000"; dr["44"] = "34000";
            //        dr["51"] = "34000"; dr["52电流2电流2"] = "34000"; dr["53"] = "34000"; dr["54"] = "34000";

            //        dt.Rows.Add(dr);
            //    }
            
            //    dataView.AllowUserToAddRows = false;
            //    dataView.DataSource = dt;// CGloabal.LModel;             
            //    dataView.Columns[0].Width = 160;
            //    dataView.ScrollBars = ScrollBars.Both;    

            }

            
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            var SelectNode = (sender as TreeView).SelectedNode;
            if (SelectNode.Text == "稳定性测试")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage3);
            }
            else if (SelectNode.Text == "寄存器测试")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage4);
            }
            else if (SelectNode.Text == "芯片设置")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage6);

                //芯片系统默认设置
                chipSelect.SelectedIndex = 0;
                pathSelect.SelectedIndex = 0;

            }
            else if (SelectNode.Text == "基本功能测试")
            {
                tabControl2.TabPages.Clear();
                tabControl2.TabPages.Add(tabPage7);
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

        private void button1_Click(object sender, EventArgs e)
        {
          // CDll.DataSaveExcel("tese", CGloabal.LModel);
        }

        private void btnSet_Click(object sender, EventArgs e)
        {
            int error;
            string strMsg = "";

            //btnLight1.BackColor = Color.LightGreen;
            //btnLight2.BackColor = Color.LightGreen;
            //btnLight3.BackColor = Color.LightGreen;
            //btnLight4.BackColor = Color.LightGreen;
            //btnLight5.BackColor = Color.LightGreen;
            //btnLight6.BackColor = Color.LightGreen;
            //btnLight7.BackColor = Color.LightGreen;
            //btnLight8.BackColor = Color.LightGreen;
            if (btnElect.Text == "关闭")//用户要连接仪器
            {            
                for (int ChanID = 1; ChanID <= 4; ChanID++)
                {
                    error = Power_Driver.isEnableChannel(CGloabal.g_InstrPowerModule.nHandle, ChanID, 1, ref strMsg);
                    if (error < 0)
                    {
                        strMsg = string.Format("电源通道{0}打开失败", ChanID);
                        CCommonFuncs.ShowHintInfor(eHintInfoType.error, strMsg);
                        return;
                    }
                    else
                    {
                        switch (ChanID)
                        {
                            case 1:
                                error = Power_Driver.SetOutputVoltage(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)volNum1.Value, strMsg);
                               btnLight1.BackColor = Color.LightGreen;
                                error = Power_Driver.SetMaxElectricityVal(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)eleNum1.Value, strMsg);
                                btnLight2.BackColor = Color.LightGreen;
                                break;
                            case 2:
                                error = Power_Driver.SetOutputVoltage(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)volNum2.Value, strMsg);
                                btnLight3.BackColor = Color.LightGreen;
                                error = Power_Driver.SetMaxElectricityVal(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)eleNum2.Value, strMsg);
                                btnLight4.BackColor = Color.LightGreen;
                                break;
                            case 3:
                                error = Power_Driver.SetOutputVoltage(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)volNum3.Value, strMsg);
                                btnLight5.BackColor = Color.LightGreen;
                                error = Power_Driver.SetMaxElectricityVal(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)eleNum3.Value, strMsg);
                                btnLight6.BackColor = Color.LightGreen;
                                break;
                            case 4:
                                error = Power_Driver.SetOutputVoltage(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)volNum4.Value, strMsg);
                                btnLight7.BackColor = Color.LightGreen;
                                error = Power_Driver.SetMaxElectricityVal(CGloabal.g_InstrPowerModule.nHandle, ChanID, (double)eleNum4.Value, strMsg);
                                btnLight8.BackColor = Color.LightGreen;
                                break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("请先打开电源");
            }
        }

        private void btnOff_Click(object sender, EventArgs e)
        {
            int error = 0;
            string strMsg = "";
            //btnLight1.BackColor = Color.WhiteSmoke;
            //btnLight2.BackColor = Color.WhiteSmoke;
            //btnLight3.BackColor = Color.WhiteSmoke;
            //btnLight4.BackColor = Color.WhiteSmoke;
            //btnLight5.BackColor = Color.WhiteSmoke;
            //btnLight6.BackColor = Color.WhiteSmoke;
            //btnLight7.BackColor = Color.WhiteSmoke;
            //btnLight8.BackColor = Color.WhiteSmoke;

            for (int ChanID = 1; ChanID <= 4; ChanID++)
            {
                error = Power_Driver.isEnableChannel(CGloabal.g_InstrPowerModule.nHandle, ChanID, 0, ref strMsg);
                if (error < 0)
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, strMsg);
                }
                else
                {
                    switch (ChanID)
                    {
                        case 1:
                            btnLight1.BackColor = Color.WhiteSmoke;
                            btnLight2.BackColor = Color.WhiteSmoke;

                            break;
                        case 2:
                            btnLight3.BackColor = Color.WhiteSmoke;
                            btnLight4.BackColor = Color.WhiteSmoke;
                            break;
                        case 3:
                            btnLight5.BackColor = Color.WhiteSmoke;
                            btnLight6.BackColor = Color.WhiteSmoke;
                            break;
                        case 4:
                            btnLight7.BackColor = Color.WhiteSmoke;
                            btnLight8.BackColor = Color.WhiteSmoke;
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// 稳定性测试——执行测试
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTest_Click(object sender, EventArgs e)
        {
            CGloabal.g_serialPorForUUT.ReceivedBytesThreshold = CGloabal.nCOM_RECV_NUMS;
            CGloabal.g_serialPorForUUT.DataReceived += new SerialDataReceivedEventHandler(g_serialPort_DataReceived); //打开串口后开始接收数据
        }

        private void btnRead1_Click(object sender, EventArgs e)
        {
            this.lsrList.GridLines = true; //显示表格线
            this.lsrList.View = View.Details;//显示表格细节
            this.lsrList.LabelEdit = true; //是否可编辑,ListView只可编辑第一列。
            this.lsrList.Scrollable = true;//有滚动条
            this.lsrList.HeaderStyle = ColumnHeaderStyle.Clickable;//对表头进行设置
            this.lsrList.FullRowSelect = true;//是否可以选择行

            //this.listView1.HotTracking = true;// 当选择此属性时则HoverSelection自动为true和Activation属性为oneClick
            //this.listView1.HoverSelection = true;
            //this.listView1.Activation = ItemActivation.Standard; //

            lsrList.Columns.Clear();
            //添加表头
            for (int i = 0; i < 8; i++)
            {
                this.lsrList.Columns.Add((i+1).ToString(), 41);
            }

            //添加各项
            ListViewItem[] p = new ListViewItem[1];
            p[0] = new ListViewItem(new string[] { "00", "aa", "bb" ,"cc","dd","ee","ff"});
           // p[1] = new ListViewItem(new string[] { "", "cc", "ggg" });
            //p[0].SubItems[0].BackColor = Color.Red; //用于设置某行的背景颜色

            this.lsrList.Items.AddRange(p);         
        }

        private void btnChipSet_Click(object sender, EventArgs e)
        {
            ChipModel CHIPMODEL = new ChipModel();
            CHIPMODEL.chipSelect = chipSelect.SelectedIndex;
            CHIPMODEL.pathSelect = pathSelect.SelectedIndex;

            CHIPMODEL.baudRate = baudRate.Text;
            CHIPMODEL.parityCheck = parityCheck.SelectedItem.ToString();
            CHIPMODEL.stopBit = stopBit.SelectedItem.ToString();
            CHIPMODEL.byteLength = byteLength.SelectedItem.ToString();

            CHIPMODEL.FIFOSelect = FIFOSelect.SelectedItem.ToString();

            CHIPMODEL.DMAPattern = DMAPattern.SelectedItem.ToString();
            CHIPMODEL.receiveFIFO = receiveFIFO.SelectedItem.ToString();
            CHIPMODEL.sendTarget = sendTarget.SelectedItem.ToString();

            CHIPMODEL.receiveInterrupt = receiveInterrupt.SelectedIndex;
            CHIPMODEL.sendInterrupt = sendInterrupt.SelectedIndex;
            CHIPMODEL.receiveCache = receiveCache.SelectedIndex;

            int error = CDll.ChipSet(CHIPMODEL);
            if (error>=0)
            {

            }

            //var chipID = chipSelect.SelectedText;
            //var pathID = pathSelect.SelectedText;

            //double baudRateStr = Convert.ToDouble(baudRate.Text);
            //var parityCheckStr = parityCheck.SelectedText;
            //var stopBitStr = stopBit.SelectedText;
            //var byteLengthStr = byteLength.SelectedText;
            //var FIFOSelectStr = FIFOSelect.SelectedText;
            //var DMAPatternStr = DMAPattern.SelectedText;
            //var receiveFIFOStr = receiveFIFO.SelectedText;
            //var sendTargetStr = sendTarget.SelectedText;
            //var receiveInterruptStr = receiveInterrupt.SelectedText;
            //var sendInterruptStr = sendInterrupt.SelectedText;
            //var receiveCacheStr = receiveCache.SelectedText;


        }

        private void chipReset_Click(object sender, EventArgs e)
        {
            var chip =(byte) chipSelect.SelectedIndex;
            var path = (byte)pathSelect.SelectedIndex;
            Byte[] cmdByte = new Byte[10] { 0XAA, 0X00, 0X00, 0X00, 0X00, 0X00, 0X00, 0X00, 0X00, 0XBB };
            cmdByte[1] = chip;
            cmdByte[2] = path;

            int  error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);
            if (error < 0)
            {
                MessageBox.Show("芯片复位失败");
                return;
            }

        }
    }
}
