using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using A1013QSystem;
using A1013QSystem.Common;

namespace DigitalCircuitSystem.DriverDAL
{
    class Power_Driver
    {
        static int PowerDefaultRM;
        const int BUFF_SIZE = 512;  

        /*************************************************
        函数原型：Init (ViChar resourceName[],HANDLE *pnHandle,ViString pErrMsg)  
        函数功能：电源模块初始化
        输入参数：
        输出参数：
        返 回 值：
        */
        public static int Init(string resourceName, ref int nHandle,string pErrMsg)
        {
            int nReturnStatus; //模块化电源底层初始化函数返回状态

            if ((nReturnStatus = visa32.viOpenDefaultRM(out PowerDefaultRM)) < 0)
                return nReturnStatus;

            if ((nReturnStatus = visa32.viOpen(PowerDefaultRM, resourceName, 0, 0, out nHandle)) < 0)
            {
                visa32.viClose(PowerDefaultRM);
                nHandle = 0;
                pErrMsg = "电源模块初始化失败！";
                return nReturnStatus;
            }


            visa32.viSetAttribute(nHandle, visa32.VI_ATTR_TERMCHAR_EN, visa32.VI_TRUE);//终止符使能
            visa32.viSetAttribute(nHandle, visa32.VI_ATTR_SEND_END_EN, visa32.VI_TRUE);//终止符使能	
            visa32.viSetAttribute(nHandle, visa32.VI_ATTR_TERMCHAR, 0xa);//终止符设置0xA

            visa32.viSetBuf(nHandle, visa32.VI_READ_BUF, 500);//RECVMAXLEN+4

            visa32.viSetAttribute(nHandle, visa32.VI_ATTR_TMO_VALUE, 500); //超时2000ms

            return 0;
        }

        //IDN查询
        public static int IDNQuery(int nInstrumentHandle, int nSimulateFlag, ref string pErrMsg)
        {
            int error = 0;
            int retCnt;


            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;

            string Commands = "*IDN?";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(Commands), Commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置查询模块ID失败";
                return error;
            }



            //读取
            byte[] byteArray = new byte[100];
            error = visa32.viRead(nInstrumentHandle, byteArray, BUFF_SIZE, out retCnt);

            if (error < 0)
            {
                pErrMsg = "读取电源模块ID失败";
                return error;
            }
            else
            {
                string strIDN = System.Text.Encoding.Default.GetString(byteArray);
                pErrMsg = strIDN;
            }
            return error;

        }

        //复位
        public static int Reset(int nInstrumentHandle, int nSimulateFlag, string strErrMsg)
        {
            int status = 0;
            string commands;
            int retCnt;

            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;

            commands = "*RST";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            return status;
        }


        //关闭模块
        public static int Close(int nInstrumentHandle, int nSimulateFlag, string strErrMsg)
        {

            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;


            if (nInstrumentHandle == 0)
            {
                strErrMsg = "电源模块已经关闭或初始化失败!";
                return -1;
            }
            if (nInstrumentHandle > 0)
            {
                visa32.viClose(nInstrumentHandle);  //关闭指定的session,事件或查表(find list)
            }
            if (PowerDefaultRM > 0)
            {
                visa32.viClose(PowerDefaultRM);
            }
            return 0;
        }

        /***********************************************************************************
           函数原型：private int WriteCmdToInstrument(int nInstrumentHandle, string StrCmd)  
           函数功能：向仪器写命令
           输入参数：
           输出参数：
           返 回 值：
           */
        public static int WriteCmdToInstrument(int nInstrumentHandle, string StrCmd)
        {
            int retCnt = 0;
            int error = 0;
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(StrCmd), StrCmd.Length, out retCnt);
            return error;
        }




        /*************************************************
        函数原型：SetOutputVoltage(HANDLE nInstrumentHandle,double dVoltageV,ViString pErrMsg)    
        函数功能：设置电源模块的通道的电压
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int SetOutputVoltage(int nInstrumentHandle,  int ChanID, double dVoltageV, string pErrMsg)
        {
            int error = 0;
            string CmdBuf;
            int retCnt;
            
            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            CmdBuf = string.Format("VOLT {0},(@{1})", dVoltageV, ChanID);
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置通道{0}输出电压失败!", ChanID);
                return error;
            }

            return error;


        }

        /*************************************************
        函数原型：SetSafeVoltage(HANDLE nInstrumentHandle,double dVoltageV,ViString pErrMsg)     
        函数功能：设置电源模块的通道的输出过压保护
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int SetSafeVoltage(int nInstrumentHandle, int nSimulateFlag, int ChanID, double dVoltageV, string pErrMsg)
        {
            int error = 0;
            string CmdBuf;
            int retCnt;

            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;


            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            CmdBuf = string.Format("VOLT:PROT:LEV {0},(@{1})", dVoltageV, ChanID);
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置通道{0}的过压保护电压值失败!", ChanID);
                return error;
            }

            return error;
        }

        /*************************************************
        函数原型： SetMaxElectricityVal(HANDLE nInstrumentHandle,double dValA,ViString pErrMsg)     
        函数功能：设置电源模块的通道的最大输出电流
        输入参数： dValA,单位A
        输出参数：
        返 回 值：
        */
        public static int SetMaxElectricityVal(int nInstrumentHandle, int ChanID, double dValA, string pErrMsg)
        {
            int error = 0;
            string CmdBuf;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            CmdBuf = string.Format("CURR {0},(@{1})", dValA, ChanID);
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置通道{0}的最大输出电流失败!", ChanID);
                return error;
            }

            return error;
        }

        /*************************************************
        函数原型：isEnableSafeElectricity(HANDLE nInstrumentHandle,int nFlag,ViString pErrMsg)    
        函数功能：设置是否启动通道的过流保护
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int isEnableSafeElectricity(int nInstrumentHandle, int nSimulateFlag, int ChanID, int nFlag, string pErrMsg)
        {
            int error = 0;
            string CmdBuf;


            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;


            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            if (nFlag == 1)
            {
                CmdBuf = string.Format("CURR:PROT:STAT ON,(@{0})", ChanID);
            }
            else
            {
                CmdBuf = string.Format("CURR:PROT:STAT OFF,(@{0})", ChanID);
            }
            error = WriteCmdToInstrument(nInstrumentHandle, CmdBuf);
            if (error < 0)
            {
                pErrMsg = string.Format("设置通道{0}是否启动过流保护失败!", ChanID);
                return error;
            }

            return error;
        }

        /*************************************************
        函数原型：isEnableChannel(HANDLE nInstrumentHandle,int nSimulateFlag,int ChanID,int nFlag,ViString pErrMsg)   
        函数功能：是否使能通道
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int isEnableChannel(int nInstrumentHandle,  int ChanID, int nFlag, ref string pErrMsg)
        {
            int error = 0;
            string CmdBuf;
            
            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            if (nFlag == 1)
            {
                CmdBuf = string.Format("OUTP ON,(@{0})", ChanID);
            }
            else
            {
                CmdBuf = string.Format("OUTP OFF,(@{0})", ChanID);
            }
            error = WriteCmdToInstrument(nInstrumentHandle, CmdBuf);
            if (error < 0)
            {
                pErrMsg = string.Format("设置是否使能通{0}道失败!", ChanID);
                return error;
            }

            return error;
        }

        public static int ReadVoltage(int nInstrumentHandle, int channel,  string strErrMsg, ref double voltage)
        {
            int status = 0;
            string commands;
            string strVal;
            int retCnt;
            byte[] byteArray = new byte[100];
          
            //:MEASure:SCALar:VOLTage:DC? (@1)
            commands = ":MEASure:SCALar:VOLTage:DC? (@" + channel.ToString() + ")";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            status = visa32.viRead(nInstrumentHandle, byteArray, BUFF_SIZE, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            strVal = System.Text.Encoding.Default.GetString(byteArray);
            voltage = Convert.ToDouble(strVal);
            if (voltage < 0.0001)
                voltage = 0;
            return status;
        }
        public static int ReadCurrent(int nInstrumentHandle, int channel, string strErrMsg, ref double voltage)
        {
            int status = 0;
            string commands;
            string strVal;
            int retCnt;
            byte[] byteArray = new byte[100];

            //:MEASure:SCALar:VOLTage:DC? (@1)
            commands = ":MEASure:SCALar:CURRent:DC? (@" + channel.ToString() + ")";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            status = visa32.viRead(nInstrumentHandle, byteArray, BUFF_SIZE, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            strVal = System.Text.Encoding.Default.GetString(byteArray);
            voltage = Convert.ToDouble(strVal);
            if (voltage < 0.0001)
                voltage = 0;
            return status;
        }

    }
}
