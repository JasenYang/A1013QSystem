using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
       
    }
}
