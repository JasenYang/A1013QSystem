using A1013QSystem;
using A1013QSystem.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace DigitalCircuitSystem.DriverDAL
{
    public class Oscilloscope_Driver
    {
        private static string strError;
        const int BUFF_SIZE = 512;
        static int ScopeDefaultRM;

        public static void TrigChannelSetting(int TrigChanID, int displayOn, double Times, double ScaleVoltV, double offsetVolt, double TrigLevel, int TrigEdge, int TrigMode)
        {
            //控制示波器
            isEnableWork(CGloabal.g_InstrScopeModule.nHandle, displayOn, TrigChanID, 1, ref strError);
            SetImpedance(CGloabal.g_InstrScopeModule.nHandle,  TrigChanID, ref strError);
            //时间轴 unit: s
            AjustTimeAxis(CGloabal.g_InstrScopeModule.nHandle,  TrigChanID, Times, ref strError);
            //垂直刻度
            AjustVoltAxis(CGloabal.g_InstrScopeModule.nHandle,   TrigChanID, ScaleVoltV, ref strError);
            //set offset voltage
            //            double offSetVlot = 3;
            SetBiasVolt(CGloabal.g_InstrScopeModule.nHandle,   TrigChanID, offsetVolt, ref strError);
            //通道触发
            SetTrigerSource(CGloabal.g_InstrScopeModule.nHandle,   TrigChanID, ref strError);
            //触发电平

            SetTrigeVolt(CGloabal.g_InstrScopeModule.nHandle,   TrigChanID, TrigLevel, ref strError);
            //触发边沿是上升沿还是下降沿
            //int rise_fall_edge = 1; // 0  positive edge   1 : negetive edge 
            SetTrigerEdge(CGloabal.g_InstrScopeModule.nHandle,   TrigChanID, TrigEdge, ref strError);
            //触发模式
            //int nTrigMode = 1;
            SetTrigeMode(CGloabal.g_InstrScopeModule.nHandle,   TrigMode, ref strError);
            //设置测量阈值add by msq
            SetTHResholds(CGloabal.g_InstrScopeModule.nHandle,   TrigMode, ref strError);
        }

        public static void ChannelSetting(int ChanID, int displayOn, double Times, double ScaleVoltV, double offsetVolt)
        {
            //控制示波器
            isEnableWork(CGloabal.g_InstrScopeModule.nHandle, displayOn, ChanID, 1, ref strError);
            SetImpedance(CGloabal.g_InstrScopeModule.nHandle,   ChanID, ref strError);
            //时间轴 unit: s
            AjustTimeAxis(CGloabal.g_InstrScopeModule.nHandle,   ChanID, Times, ref strError);
            //垂直刻度
            AjustVoltAxis(CGloabal.g_InstrScopeModule.nHandle,   ChanID, ScaleVoltV, ref strError);
            //set offset voltage
            //            double offSetVlot = 3;
            SetBiasVolt(CGloabal.g_InstrScopeModule.nHandle,   ChanID, offsetVolt, ref strError);
            //通道触发
            //            SetTrigerSource(CGloabal.g_InstrScopeModule.nHandle,   ChanID, ref strError);
            //触发电平

        }


        /***********************************************************************************
        函数原型：ReadDoubleValFromInstrument(int nInstrumentHandle, out double dRtnVal)
        函数功能：从指定的仪器中读数
        输入参数：
        输出参数：
        返 回 值：
        */
        public static int ReadDoubleValFromInstrument(int nInstrumentHandle, out double dRtnVal)
        {
            int error = 0;
            int retCnt;
            byte[] ReadBuf = new byte[100];
            error = visa32.viRead(nInstrumentHandle, ReadBuf, 100, out retCnt);
            if (error < 0)
            {
                dRtnVal = -1.0;
                return error;
            };
            ReadBuf[retCnt] = 0; //字符串结尾加0

            string str = System.Text.Encoding.Default.GetString(ReadBuf);
            dRtnVal = Convert.ToDouble(str);

            return 0;

        }


        /******************************************************************************************
        * 函数原型：ConnectSpecificInstrument(string strInstruName,string resourceName)
        * 函数功能：根据输入的仪器名进行连接。连接后的句柄存储在相应的句柄参数中
        * 输入参数：strInstruName，仪器名字；resourceName，资源名字，用于VISA连接
        * 输出参数：
        * 创 建 者：yzx
        * 创建日期：2016.7.27
        * 修改说明：
        * */
        public static int ConnectScope(string resourceName)
        {
            int error = 0;

            if (CGloabal.g_InstrScopeModule.nHandle == 0)
            {
                error = Oscilloscope_Driver.Init(resourceName,   ref CGloabal.g_InstrScopeModule.nHandle, ref strError);
                if (error < 0)
                {
                    CGloabal.g_InstrScopeModule.bInternet = false;
                    CCommonFuncs.ShowHintInfor(eHintInfoType.error, CGloabal.g_InstrScopeModule.strInstruName + "连接失败");
                }
                else
                {
                    CGloabal.g_InstrScopeModule.bInternet = true;
                }
            }
            else
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.hint, CGloabal.g_InstrScopeModule.strInstruName + "已经处于连接状态");
            }

            return error;

        }

        /*************************************************
        函数原型：ScopeModule_Init (ViChar resourceName[],HANDLE *pnHandle,ViString pErrMsg)  
        函数功能：示波器模块初始化
        输入参数：
        输出参数：
        返 回 值：
        */
        public static int Init(string resourceName, ref int pnHandle, ref string pErrMsg)
        {
            int nReturnStatus; //模块化电源底层初始化函数返回状态
          
            if ((nReturnStatus = visa32.viOpenDefaultRM(out ScopeDefaultRM)) < 0)
                return nReturnStatus;

            if ((nReturnStatus = visa32.viOpen(ScopeDefaultRM, resourceName, 0, 0, out pnHandle)) < 0)
            {
                visa32.viClose(ScopeDefaultRM);
                pnHandle = 0;
                pErrMsg = "电源模块初始化失败！";
                return nReturnStatus;
            }

            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_TERMCHAR_EN, visa32.VI_TRUE);//终止符使能
            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_SEND_END_EN, visa32.VI_TRUE);//终止符使能	
            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_TERMCHAR, 0xA);//终止符设置0xA
            visa32.viSetBuf(pnHandle, visa32.VI_READ_BUF, 500);//RECVMAXLEN+4
            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_TMO_VALUE, 2000); //超时2000ms

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
            string strIDN = System.Text.Encoding.Default.GetString(byteArray);
            if (error < 0)
            {
                pErrMsg = "读取电源模块ID失败";
                return error;
            }
            else
            {
                //sprintf (pErrMsg,"%s",Commands);    //将独处来的IDN信息存放到pErrMsg中进行返回
                pErrMsg = strIDN;
            }
            return error;
        }
        //STB查询
        public static int STBQuery(int nInstrumentHandle, ref short status)
        {
            visa32.viReadSTB(nInstrumentHandle, ref  status);
            return 0;
        }
        //STB清除
        public static int STBClear(int nInstrumentHandle)
        {
            int retCnt;
            string commands = "*CLS";
            visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            return 0;
        }


        //复位
        public static int Reset(int nInstrumentHandle, int nSimulateFlag, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;
            string commands = "*RST";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            Thread.Sleep(6000); //wait the ocillocope to ready
            return status;
        }

        //关闭模块
        public static int defaultSetting(int nInstrumentHandle, int nSimulateFlag, ref string strErrMsg)
        {
            int status = 0;
            string commands;
            int retCnt;
            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;
            if (0 == nInstrumentHandle)
            {
                strErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            if (nInstrumentHandle != 0)
            {
                visa32.viClose(nInstrumentHandle);  //关闭指定的session,事件或查表(find list)
            }
            commands = ":SYSTem:PRESet DEFAULT";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "VISA函数错误!";
                return status;
            }
            return 0;
        }

        //关闭模块
        public static int Close(int nInstrumentHandle, int nSimulateFlag, ref string strErrMsg)
        {

            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;


            if (0 == nInstrumentHandle)
            {
                strErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            if (nInstrumentHandle != 0)
            {
                visa32.viClose(nInstrumentHandle);  //关闭指定的session,事件或查表(find list)
            }
            if (ScopeDefaultRM != 0)
            {
                visa32.viClose(ScopeDefaultRM);
            }
            return 0;
        }



        ///*************************************************
        //函数原型：ScopeModule_isOpenChannel(HANDLE nInstrumentHandle,int nSimulateFlag,int ChanID,int nFlag,char* pErrMsg)     
        //函数功能：是否打开通道选择的通道
        //输入参数：nFlag：0，关闭；1，打开；
        //输出参数：
        //返 回 值：
        //*/
        public static int isOpenChannel(int nInstrumentHandle, int nSimulateFlag, int ChanID, int nFlag, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;


            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;


            if (nInstrumentHandle == 0)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }


            //设置
            if (nFlag == 1)
            {
                commands = ":CHAN" + ChanID.ToString() + ":DISPlay ON";
            }
            else
            {
                commands = ":CHAN" + ChanID.ToString() + ":DISPlay OFF";
            }
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置通道" + ChanID.ToString() + "是否打开失败!";
                return error;
            }
            return error;
        }

        /*************************************************
        函数原型：ViInt32 __stdcall  ScopeModule_SetVolt_couple(HANDLE nInstrumentHandle,int nSimulateFlag,int ChanID,double dTimeS, char* pErrMsg)   
        函数功能：设置通道的测试 
        输入参数： dTimeS，数值，单位s
        输出参数： 
        返 回 值：
        说明：
        DSO6054A：:CHAN1:IMP ONEM
            大示波器 :CHAN1:INP DC
        */
        public static int SetVolt_couple(int nInstrumentHandle, int nSimulateFlag, int ChanID, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;
            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;
            if (nInstrumentHandle == 0)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            //设置 耦合方式
            commands = ":CHAN" + ChanID.ToString() + ":COUP DC";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置耦合方式" + ChanID.ToString() + "时间坐标轴失败!";
                return error;
            }
            //设置 测量电压 
            commands = ":CHAN" + ChanID.ToString() + ":UNIT VOLT";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置测量电压" + ChanID.ToString() + "时间坐标轴失败!";
                return error;
            }
            return error;

        }



        /*************************************************
        函数原型：ViInt32 __stdcall  ScopeModule_SetImpedance(HANDLE nInstrumentHandle,int nSimulateFlag,int ChanID,double dTimeS, char* pErrMsg)   
        函数功能：调整指定通道的时间轴 
        输入参数： dTimeS，数值，单位s
        输出参数： 
        返 回 值：
        说明：
        DSO6054A：:CHAN1:IMP ONEM
            大示波器 :CHAN1:INP DC
        */
        public static int SetImpedance(int nInstrumentHandle, int ChanID, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;
          
            if (nInstrumentHandle == 0)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            //设置
            commands = ":CHAN" + ChanID.ToString() + ":IMP ONEM"; //小示波器
            //commands = ":CHAN" + ChanID.ToString() + ":INPut DC"; //小示波器
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置通道阻抗" + ChanID.ToString() + "时间坐标轴失败!";
                return error;
            }

            return error;

        }


        /*************************************************
        函数原型：ViInt32 __stdcall  ScopeModule_AjustTimeAxis(HANDLE nInstrumentHandle,int nSimulateFlag,int ChanID,double dTimeS, char* pErrMsg)   
        函数功能：调整指定通道的时间轴 
        输入参数： dTimeS，数值，单位s
        输出参数：
        返 回 值：
        */
        public static int AjustTimeAxis(int nInstrumentHandle, int ChanID, double dTimeS, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;
            if (nInstrumentHandle == 0)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            //设置
            //	commands = ":CHAN" + ChanID.ToString() + ":DISPLAY:SCALE "+ dTimeS.ToString();  //小示波器
            commands = "TIMebase:SCALe " + dTimeS.ToString();    //msos804a 
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置通道" + ChanID.ToString() + "时间坐标轴失败!";
                return error;
            }

            return error;

        }
        /*************************************************
        函数原型：__stdcall  ScopeModule_AjustVoltAxis(HANDLE nInstrumentHandle,int nSimulateFlag,int ChanID,double dVoltvalV, char* pErrMsg)
        函数功能：调整指定通道的电压值
        输入参数： dVoltvalV，数值，单位V
        输出参数：
        返 回 值：
        */
        public static int AjustVoltAxis(int nInstrumentHandle, int ChanID, double dVoltvalV, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            //设置
            commands = ":CHAN" + ChanID.ToString() + ":SCALE " + dVoltvalV.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置通道" + ChanID.ToString() + "电压坐标轴失败!";
                return error;
            }

            return error;

        }

        /*************************************************
        函数原型： ScopeModule_SetBiasVolt(HANDLE nInstrumentHandle,double dVoltvalV, char* pErrMsg)
        函数功能：设置通道1偏置电压
        输入参数： dVoltvalV，数值，单位V
        输出参数：
        返 回 值：
        */
        public static int SetBiasVolt(int nInstrumentHandle,  int ChanID, double dVoltvalV, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            //设置
            commands = ":CHAN" + ChanID.ToString() + ":OFFSet " + dVoltvalV.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            {
                pErrMsg = "设置通道" + ChanID.ToString() + "偏置电压失败!";
                return error;
            }
        }
        /*************************************************
        函数原型：ScopeModule_SetTHResholds(int nInstrumentHandle, int nSimulateFlag, int ChanID, ref string pErrMsg)
        函数功能：设置测量阈值
        输入参数： 
        输出参数：
         * 创建日期：2016.8.16 add by msq
        返 回 值：:MEASure:DEFine THResholds, STANdard
        */
        public static int SetTHResholds(int nInstrumentHandle, int ChanID, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            commands = ":MEASure:DEFine THResholds, STANdard";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置触发通道失败!";
                return error;
            }
            return error;
        }
        /*************************************************
        函数原型：ScopeModule_SetTrigeMode(HANDLE nInstrumentHandle,int nFlag,char* pErrMsg)    
        函数功能：触发模式选择   
        输入参数：nFlag:0,single模式；1，触发模式
        输出参数：
        返 回 值：
        */
        public static int SetTrigeMode(int nInstrumentHandle,  int nFlag, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            if (nFlag == 1)  //触发模式
            {
                commands = ":TRIGGER:SWEEP TRIGGERED";
            }
            else  //AUTO
            {
                commands = ":TRIGGER:SWEEP AUTO";
            }
            error = error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置触发模式失败!";
                return error;
            }

            return error;

        }

        /*************************************************
        函数原型：ScopeModule_SetAverageNums(HANDLE nInstrumentHandle,int nFlag,char* pErrMsg)    
        函数功能：设置平均次数   
        输入参数：nNums，必须大于2；
        输出参数：
        返 回 值：
        */
        public static int SetAverageNums(int nInstrumentHandle, int nNums, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;
            
            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            if (nNums < 2)  //触发模式
            {
                nNums = 2;
            }
            commands = ":ACQuire:AVERage:COUNT " + nNums.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置平均次数失败!";
                return error;
            }

            return error;

        }


        /*************************************************
        函数原型：ScopeModule_isEnableAverage(HANDLE nInstrumentHandle,int nFlag,char* pErrMsg)    
        函数功能：是否启动平均测量 
        输入参数：nFlag：0，禁止；1，使能；
        输出参数：
        返 回 值：
        */
        public static int isEnableAverage(int nInstrumentHandle,  int nFlag, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;
            
            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            if (nFlag == 1)
            {
                commands = ":ACQuire:AVERage ON";
            }
            else
            {
                commands = ":ACQuire:AVERage OFF";
            }
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置是否使能平均失败!";
                return error;
            }

            return error;

        }

        /*************************************************
        函数原型：ScopeModule_isEnableWork(HANDLE nInstrumentHandle,int nFlag,ViString pErrMsg)    
        函数功能：是否启动示波器开始工作
        输入参数： nFlag：0，不启动；1，启动；
        输出参数：
        返 回 值：
        */
        public static int isEnableWork(int nInstrumentHandle, int displayOn, int ChanID, int nFlag, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;


            //检测是否处于模拟状态


            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }


            //设置
            if (nFlag == 1)
            {
                commands = ":RUN";
            }
            else
            {
                commands = ":STOP";
            }
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置示波器是否开始工作失败!";
                return error;
            }
            Thread.Sleep(100);

            commands = ":MARKer:MODE MEASurement";      //open marker
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (displayOn == 1)
                commands = ":CHANNEL" + ChanID.ToString() + ":DISPLAY ON";          //display channel 1
            else
                commands = ":CHANNEL" + ChanID.ToString() + ":DISPLAY OFF";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //commands = ":CHANNEL2:DISPLAY ON";          //display channel 2
            //error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //Thread.Sleep(100);
            //error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //commands = ":CHANNEL3:DISPLAY ON";          //display channel 1
            //error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //commands = ":CHANNEL4:DISPLAY ON";          //display channel 2
            //error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            Thread.Sleep(100);
            return error;
        }

        /*************************************************
        函数原型： ScopeModule_SetTrigeVolt(HANDLE nInstrumentHandle,double dVoltvalV, char* pErrMsg)
        函数功能：设置触发电平
        输入参数： nChannelID，通道编号，取值范围：1-4；
                   dVoltvalV，触发电平值
        输出参数：
        返 回 值：
        */
        public static int SetTrigeVolt(int nInstrumentHandle,  int ChanID, double dVoltvalV, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            //设置
            commands = ":TRIGGER:LEVEL CHAN" + ChanID.ToString() + "," + dVoltvalV.ToString();
            //commands = string.Format(":TRIG:EDGE:LEV %f", dVoltvalV);  
            //commands = ":TRIG:EDGE:LEV " + dVoltvalV.ToString(); //小示波器
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置通道%d触发电平失败!", ChanID);
                return error;
            }
            return error;
        }

        /*************************************************
        函数原型： ScopeModule_SetAuxTrigeVolt(HANDLE nInstrumentHandle,double dVoltvalV, char* pErrMsg)
        函数功能：设置AUX触发电平
        输入参数： dVoltvalV，触发电平值
        输出参数：
        返 回 值：
        */
        public static int SetAuxTrigeVolt(int nInstrumentHandle,  double dVoltvalV, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            //	commands = string.Format(":TRIGGER:LEVEL AUX,%lf",dVoltvalV); 
            commands = ":TRIGGER:LEVEL AUX," + dVoltvalV.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置通道AUX触发失败!";
                return error;
            }
            return error;
        }
        /*************************************************
        函数原型：ScopeModule_SetTrigerEdge(int nInstrumentHandle, int nSimulateFlag, int ChanID, ref string pErrMsg)
        函数功能：选择触发通道
        输入参数： 
        输出参数：
         * 创建日期：2016.8.1
        返 回 值：/*************************************************
        函数原型：ScopeModule_SetTrigerEdge(int nInstrumentHandle, int nSimulateFlag, int ChanID, ref string pErrMsg)
        函数功能：选择触发通道
        输入参数： 
        输出参数：
         * 创建日期：2016.8.1
        返 回 值：
        */
        public static int SetTrigerEdge(int nInstrumentHandle, int ChanID, int rise_fall_triger, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }
            if (rise_fall_triger == 1)
                commands = ":TRIG:EDGE:SLOP NEG";
            else
                commands = ":TRIG:EDGE:SLOP POS";

            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置触发通道失败!";
                return error;
            }
            return error;
        }


        /*************************************************
        函数原型：ScopeModule_SetTrigerSource(int nInstrumentHandle, int nSimulateFlag, int ChanID, ref string pErrMsg)
        函数功能：选择触发通道
        输入参数： 
        输出参数：
         * 创建日期：2016.8.1
        返 回 值：
        */
        public static int SetTrigerSource(int nInstrumentHandle, int ChanID, ref string pErrMsg)
        {
            int error = 0;
            string commands;
            int retCnt;
            
            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            //设置
            //   commands = string.Format(":MTESt:TRIGger:SOURce CHANnel{0}", ChanID);  //小示波器
            commands = ":TRIGger:EDGE1:SOURce CHAN" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = "设置触发通道失败!";
                return error;
            }
            return error;
        }

        /*************************************************
        函数原型：ScopeModule_GetCh1VoltVal(HANDLE nInstrumentHandle,double *pdVoltvalV, char* pErrMsg)
        函数功能：启动通道1信号幅度测量并获得测量的结果
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int GetChanVoltVal(int nInstrumentHandle,  int ChanID, int isSINK, ref double pdVoltvalV, string pErrMsg)
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
            //CmdBuf = string.Format(":MEASure:VRMS? DISPLAY,DC,CHAN%d",ChanID);  
            // CmdBuf = ":MEASure:VRMS? DISPLAY,CHAN" + ChanID.ToString();
            if (isSINK == 1)
                CmdBuf = "MEASure:VBASe? CHAN" + ChanID.ToString();
            else if (isSINK == 0)
                CmdBuf = "MEASure:VTOP? CHAN" + ChanID.ToString();
            else if (isSINK == 2)
                CmdBuf = "MEASure:VAV? CHAN" + ChanID.ToString();
            else
                CmdBuf = "MEASure:VTOP? CHAN" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置通道%dVRMS测试失败!", ChanID);
                return error;
            }

            Thread.Sleep(500);
            error = ReadDoubleValFromInstrument(nInstrumentHandle, out pdVoltvalV);
            return error;
        }

        /*************************************************
        函数原型：ScopeModule_GetRisingTime(HANDLE nInstrumentHandle,double *pVal, char* pErrMsg) 
        函数功能：获取通道2的上升沿时间
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int GetRisingTime(int nInstrumentHandle, int ChanID, ref double pVal, ref string pErrMsg)
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
            //CmdBuf = string.Format(":MEASure:EDGE? CHAN%d,RIS",ChanID);
            CmdBuf = ":MEASure:CLEAR" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":MEASure:RISETIME" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":MEASure:RISETIME? CHAN" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置查询通道%d上升时间指令失败!", ChanID);
                return error;
            }

            Thread.Sleep(500);
            //读数
            error = ReadDoubleValFromInstrument(nInstrumentHandle, out pVal);

            return error;
        }
        /*************************************************
        函数原型：ScopeModule_GetFallingTime(HANDLE nInstrumentHandle,double *pVal, char* pErrMsg) 
        函数功能：获取通道的下降沿沿时间
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int GetFallingTime(int nInstrumentHandle, int ChanID, ref double pVal, string pErrMsg)
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
            //CmdBuf = string.Format(":MEASure:EDGE? CHAN{0},FALL",ChanID); 
            CmdBuf = ":MEASure:CLEAR" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":MEASure:FALLTIME" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":MEASure:FALLTIME? CHAN" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置查询通道%d下降时间指令失败!", ChanID);
                return error;
            }

            Thread.Sleep(500);

            //读数
            error = ReadDoubleValFromInstrument(nInstrumentHandle, out pVal);

            return error;
        }
        /*************************************************
        函数原型：GetPPulseWidth(HANDLE nInstrumentHandle,double *pVal, char* pErrMsg) 
        函数功能：获取通道的下降沿沿时间
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int GetPPulseWidth(int nInstrumentHandle, int ChanID, ref double pVal, string pErrMsg)
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
            //CmdBuf = string.Format(":MEASure:EDGE? CHAN{0},FALL",ChanID);  
            CmdBuf = ":MEASure:CLEAR" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":MEASure:PWIDTH" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":MEASure:PWIDTH? CHAN" + ChanID.ToString();
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(100);
            //读数
            error = ReadDoubleValFromInstrument(nInstrumentHandle, out pVal);

            return error;
        }
        /*************************************************
        函数原型：ScopeModule_GetDelayTimeVal(HANDLE nInstrumentHandle,double *pVal, char* pErrMsg)
        函数功能：启动时延差测量并获取时间差值, 
        输入参数： 
        输出参数：
        返 回 值：
        */
        public static int GetDelayTimeVal(int nInstrumentHandle, int cmd_sel, int Chan1ID, int Chan2ID, ref double pVal, string pErrMsg)
        {
            int error = 0;
            string CmdBuf;
            int retCnt;


            //检测是否处于模拟状态
            // if (nSimulateFlag == 1)
            //     return 0;

            if (0 == nInstrumentHandle)
            {
                pErrMsg = "已经关闭或初始化失败!";
                return -1;
            }

            CmdBuf = ":MEAS:CLEAR";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            if (cmd_sel == 0)
                CmdBuf = ":MEAS:DEF -1,+2";//dso6054a
            else if (cmd_sel == 1)
                CmdBuf = ":MEAS:DEF -1,-2";//dso6054a
            else if (cmd_sel == 2)
                CmdBuf = ":MEAS:DEF +1,+2";//dso6054a
            else if (cmd_sel == 3)
                CmdBuf = ":MEAS:DEF +1,-2";//dso6054a
            // CmdBuf = ":MEASURE: DELTATIME: DEFINE RISING,1,MIDDLE,FALLING,1,MIDDLE";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);

            //设置
            //	CmdBuf = string.Format(":MEASure:DELTatime? CHAN{0},CHAN{1}",Chan1ID,Chan2ID);  
            //   CmdBuf = ":MEASure:DELTatime? CHAN" + Chan1ID.ToString() + ",CHAN" + Chan2ID.ToString();
            //   error =visa32.viWrite (nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            CmdBuf = ":MEAS:DEL CHAN1,CHAN2";   //dso6054a
            //CmdBuf = ":MEAS:DELT CHAN1,CHAN2";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);

            CmdBuf = ":MEAS:TDEL?";   //dso6054a
            //CmdBuf = "MEASure:DELTatime? CHAN1,CHAN2";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            if (error < 0)
            {
                pErrMsg = string.Format("设置查询时间差指令失败!");
                return error;
            }

            Thread.Sleep(500);

            //读数
            error = ReadDoubleValFromInstrument(nInstrumentHandle, out pVal);


            return error;
        }

        public static int SavePicture(int nInstrumentHandle, string file_name, string pErrMsg)
        {
            int error = 0;
            string CmdBuf;
            int retCnt;
            /*
             * :SAVE:IMAGE:FACTors 1
                :SAVE:IMAGE:FORmat BMP24bit
                :SAVE:IMAGE:IGColors 0
                :SAVE:imAGE:START "20NS_PULSE"
             * wait 10s 
             * */
            //设置
            //            CmdBuf = ":SAVE:IMAGE:FACTors 1";
            //            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":SAVE:IMAGE:FORmat PNG";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":SAVE:IMAGE:IGColors 0";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(500);
            CmdBuf = ":SAVE:imAGE:START " + "\"" + file_name + "\"";
            error = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(CmdBuf), CmdBuf.Length, out retCnt);
            Thread.Sleep(10000);

            return error;
        }

        #region 
           //触发设置
        public static int SetTick(int nInstrumentHandle, string strErrMsg)
        {
            int status = 0;
            int retCnt;
           // :SYST:PRES FACT
            string commands = ":SYST:PRES DEF";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
            commands = ":TRIG:EDGE:COUP DC";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
           
            commands = ":TRIG:EDGE:SOUR CHAN4";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
          
            commands = ":TRIG:EDGE:SLOP POS";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
           
            commands = ":TRIG:LEV CHAN4,0";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }            
            return status;
        }

        //通道做减法 修改纵坐标和偏置
        public static int UpdateAxis(int nInstrumentHandle, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            //string commands = ":CHAN1:OFFS 0";
            //status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //if (status < 0)
            //{
            //    strErrMsg = "函数错误!";
            //    return status;
            //}

            string commands = ":CHAN1:RANG 5";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
         
            commands = ":CHAN1:DIFF ON";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
       
            commands = ":CHAN1:DISP ON";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
           
            //commands = ":CHAN2:DISP OFF";
            //status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //if (status < 0)
            //{
            //    strErrMsg = "函数错误!";
            //    return status;
            //}
          
            //commands = ":CHAN3:DISP OFF";
            //status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //if (status < 0)
            //{
            //    strErrMsg = "函数错误!";
            //    return status;
            //}
           
            //commands = ":CHAN4:DISP OFF";
            //status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            //if (status < 0)
            //{
            //    strErrMsg = "函数错误!";
            //    return status;
            //}



            //Thread.Sleep(2000);
            return status;
        }

        //修改时间基数
        public static int UpdateTimeBase(int nInstrumentHandle, double RangeValue ,string strErrMsg)
        {
            int status = 0;
            int retCnt;
            string commands = string.Format(":TIM:POS 0");
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
            
            commands = string.Format(":TIM:RANG {0}", RangeValue);
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            } 
           // Thread.Sleep(2000);
            return status;
        }

        //启动余晖
        
        public static int StartYH(int nInstrumentHandle, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            string commands = ":DISP:CGR ON";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
            return status;
        }
       

        //最大显示
        public static int SetMaxShow(int nInstrumentHandle, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            string commands = ":DISP:WIND:MAXAREA1";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
            return status;
        }
        //启动测量
        public static int SetStartRun(int nInstrumentHandle, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            string commands = ":RUN";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
            

            commands = ":AUT:VERT CHAN1";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }

            Thread.Sleep(3000);

            //启动余晖
            commands = ":DISP:CGR ON";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }
            //清空显示
            commands = ":CDIS";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }

            Thread.Sleep(3000);//等待三秒后停止测试   
            return status;
        }
        public static int SetStopRun(int nInstrumentHandle, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            string commands = ":STOP";
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }

            //停止运行后截屏

            return status;
        }

        public static int SavePictureUrl(int nInstrumentHandle,string fileName, string strErrMsg)
        {
            int status = 0;
            int retCnt;
            string commands = string.Format(@":DISK:SAVE:IMAG '\New folder\{0}',JPEG,SCR,ON,NORM,ON",fileName);
            status = visa32.viWrite(nInstrumentHandle, System.Text.Encoding.Default.GetBytes(commands), commands.Length, out retCnt);
            if (status < 0)
            {
                strErrMsg = "函数错误!";
                return status;
            }

            //停止运行后截屏

            return status;
        }


        #endregion

    }
}
