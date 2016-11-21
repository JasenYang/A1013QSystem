using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace A1013QSystem.Common
{
     public class CGloabal
    {
        //定义仪器管理的参数类
        public class InstrMentsParas
        {
            public string strInstruName;  //仪器名字
            public string ipAdress;      //ip地址
            public int port;             //端口号
            public bool bInternet;       //连接状态， 0断开  1连接
            public int nHandle;        //句柄
            public Image imag;         //各自仪器的图片         


            public InstrMentsParas(string name)
            {
                this.strInstruName = name;
                this.ipAdress = "172.141.10.30";
                this.nHandle = 8000;
                this.bInternet = false;
                this.nHandle = 0; //默认为0

            }
        };

        //定义三个仪器的对象
        public static InstrMentsParas g_InstrPowerModule = new InstrMentsParas("电源");
        public static InstrMentsParas g_InstrScopeModule = new InstrMentsParas("示波器");
        public static InstrMentsParas g_InstrMultimeterModule = new InstrMentsParas("万用表");

        public static InstrMentsParas g_curInstrument; //表征当前正在操作的仪器

        public static string g_curTupple;

        //定义常量
        public const int nCOM_RECV_NUMS = 24;//定义串口接收到个数时，触发接收事件

        //定义数据list
        public static  List<RecordModel> LModel = new List<RecordModel>();

        //导入系统库
        [DllImport("user32.dll", EntryPoint = "SendMessage")]//dll导入
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll", EntryPoint = "PostMessage")]//dll导入
        public static extern int PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        #region //串口函数
        public static SerialPort g_serialPorForUUT;//定义用于UUT通信的串口
        public static bool g_bIsComRecvedDataFlag = false; //串口是否收到数据
        public static byte[] g_ComReadBuf = new byte[nCOM_RECV_NUMS];//存储串口读到的数据


        //定义一个专用于万用表模块串口通信的串口
        public static SerialPort g_serialPortForMultimeter;

        public static bool CurSerialPortFlag = false;//当前串口标识：false，板卡串口；true，万用表串口。
        //关闭串口
        public static int CloseCom(ref SerialPort comPort)
        {
            comPort.Close();
            return 0;
        }

        //向串口写入数据
        public static int WriteToCom(SerialPort comPort, byte[] cmdBuf, int WriteLen)
        {

            try
            {
                comPort.Write(cmdBuf, 0, WriteLen);
            }
            catch
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.error, "串口写出错！");
                return -1;
            }

            return 0;

        }

        //串口读
        public static int ReadCom(SerialPort comPort, byte[] cmdBuf, int ReadLen)
        {
            try
            {
                comPort.Read(cmdBuf, 0, ReadLen);
            }
            catch
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.error, "串口读出错！");
                return -1;
            }

            return 0;
        }




        #endregion

    }
}
