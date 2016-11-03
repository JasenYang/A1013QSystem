using A1013QSystem.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace A1013QSystem.DriverCommon
{
     public class Multimeter_Driver
    {
        const int BUFF_SIZE = 512;
        static int MultimeterDefaultRM;
        static int nMultimeterHandle; //万用表句柄

        public static int Connect_Multimeter(string resourceName)
        {

            int nReturnStatus; //模块化电源底层初始化函数返回状态

            if ((nReturnStatus = visa32.viOpenDefaultRM(out MultimeterDefaultRM)) < 0)
                return nReturnStatus;

            if ((nReturnStatus = visa32.viOpen(MultimeterDefaultRM, resourceName, 0, 0, out nMultimeterHandle)) < 0)
            {
                visa32.viClose(MultimeterDefaultRM);
                nMultimeterHandle = 0;
                return nReturnStatus;
            }

            visa32.viSetAttribute(nMultimeterHandle, visa32.VI_ATTR_TERMCHAR_EN, visa32.VI_TRUE);//终止符使能
            visa32.viSetAttribute(nMultimeterHandle, visa32.VI_ATTR_SEND_END_EN, visa32.VI_TRUE);//终止符使能	
            visa32.viSetAttribute(nMultimeterHandle, visa32.VI_ATTR_TERMCHAR, 0xA);//终止符设置0xA
            visa32.viSetBuf(nMultimeterHandle, visa32.VI_READ_BUF, 500);//RECVMAXLEN+4
            visa32.viSetAttribute(nMultimeterHandle, visa32.VI_ATTR_TMO_VALUE, 2000); //超时2000ms

            return 0;
        }

        //IDN查询
        public static int IDNQuery(out string strIDN)
        {
            int error = 0;
            int retCnt;

            string Commands = "*IDN?";
            error = visa32.viWrite(nMultimeterHandle, System.Text.Encoding.Default.GetBytes(Commands), Commands.Length, out retCnt);
            if (error < 0)
            {
                strIDN = string.Empty;
                return error;
            }
            //读取
            byte[] byteArray = new byte[100];
            error = visa32.viRead(nMultimeterHandle, byteArray, BUFF_SIZE, out retCnt);
            if (error < 0)
            {
                strIDN = string.Empty;
                return error;
            }
            else
            {
                strIDN = System.Text.Encoding.Default.GetString(byteArray);
            }
            return error;
        }
        //复位
        public static int Reset()
        {
            int error = 0;
            int retCnt;

            string commands = "*RST";
            error = visa32.viWrite(nMultimeterHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                return error;
            }
            Thread.Sleep(6000); //wait the ocillocope to ready
            return error;
        }



        //关闭模块
        public static int Close()
        {

            if (nMultimeterHandle != 0)
            {
                visa32.viClose(nMultimeterHandle);  //关闭指定的session,事件或查表(find list)
            }
            if (MultimeterDefaultRM != 0)
            {
                visa32.viClose(MultimeterDefaultRM);
            }
            return 0;
        }

        //设置电流测试量程为自动
        public static int SetMeasureRangeAuto()
        {
            int error = 0;
            int retCnt;

            string commands = ":CONFigure:CURRent:DC 0.2 1e-06"; //changed 16.09.04 by msq
            error = visa32.viWrite(nMultimeterHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                return error;
            }
            return error;
        }

        //读取当前电流测试结果
        public static int ReadElectricity(out double dVal)
        {
            int error = 0;
            int retCnt = 0;
            string strVal;

            //设置
            string commands = ":MEASure:CURRent:DC? 0.2,1e-06";     //changed 16.09.04 by msq
            error = visa32.viWrite(nMultimeterHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                dVal = -10000.0;
                return error;
            }

            //读取
            byte[] byteArray = new byte[100];
            error = visa32.viRead(nMultimeterHandle, byteArray, BUFF_SIZE, out retCnt);
            if (error < 0)
            {
                dVal = -10000.0;
                return error;
            }
            else
            {
                strVal = System.Text.Encoding.Default.GetString(byteArray);
                dVal = Convert.ToDouble(strVal);
            }
            return 0;
        }
    }
}
