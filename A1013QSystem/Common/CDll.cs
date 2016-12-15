using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using DigitalCircuitSystem.DriverDAL;
using A1013QSystem.DriverCommon;
using System.Data.OleDb;
using System.Data;
using A1013QSystem.Model;
using System.Threading;

namespace A1013QSystem.Common
{
    class CDll
    {
        private static string strError;


        [DllImport("kernel32")]
        //section：要读取的段落名
        // key: 要读取的键
        //defVal: 读取异常的情况下的缺省值
        //retVal: key所对应的值，如果该key不存在则返回空值
        //size: 值允许的大小
        //filePath: INI文件的完整路径和文件名
        private static extern int GetPrivateProfileString(string section, string key, string defVal, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32")]
        //section: 要写入的段落名
        //key: 要写入的键，如果该key存在则覆盖写入
        //val: key所对应的值
        //filePath: INI文件的完整路径和文件名
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);


        /*************************************************
      * 函数原型：private string GetContentValueFromFile(string strFilePath,string Section, string key)
      * 函数功能：从ini文件的指定段中指定的key中获取字符串数值信息
      * 输入参数：strFilePath，ini文件的路径信息；Section，段信息；key，key信息。
      * 输出参数：返回获得字符串类型的数值信息
      * 创 建 者：yzx
      * 创建日期：2016.7.23
      * 修改说明：
     */
        public static string GetValueFromIniFile(string strFilePath, string Section, string key)
        {
            int nRntCount;
            StringBuilder temp = new StringBuilder(1024);
            nRntCount = GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            if (nRntCount == 0)
            {
                return null;
            }
            else
            {
                return temp.ToString();
            }

        }

        /*************************************************
         * 函数原型：private long WriteValueToIniFile(string strFilePath,string Section, string key,string strValue)
         * 函数功能：向ini文件的指定段中指定的key中写入字符串数值信息
         * 输入参数：strFilePath，ini文件的路径信息；Section，段信息；key，key信息。
         * 输出参数：
         * 创 建 者：yzx
         * 创建日期：2016.7.23
         * 修改说明：
        */
        public static long WriteValueToIniFile(string strFilePath, string Section, string key, string strValue)
        {

            long error = 0;

            error = WritePrivateProfileString(Section, key, strValue, strFilePath);
            if (error < 0)
            {
                MessageBox.Show("ini文件写入出错！");
            }
            return error;
        }

        /*************************************************
         * 函数原型：private void GetAlltheInstumentsParasFromImiFile()
         * 函数功能：从ini文件中获取所有的仪器的参数信息.
         * 输入参数：
         * 输出参数：
         * 创 建 者：yzx
         * 创建日期：2016.7.23
         * 修改说明：
        */
        public static void GetAlltheInstumentsParasFromIniFile()
        {
            string strPort;
            string strFilePath;
            string strBusType;

            //获取ini文件的相对路径
            strFilePath = System.Windows.Forms.Application.StartupPath + "\\APP_INFOS.ini";
            if (File.Exists(strFilePath))//首先判读INI文件是否存在
            {

                CGloabal.g_InstrPowerModule.ipAdress = GetValueFromIniFile(strFilePath, "电源", "IP地址");
                strPort = GetValueFromIniFile(strFilePath, "电源", "端口号");
                CGloabal.g_InstrPowerModule.port = int.Parse(strPort);

                CGloabal.g_InstrScopeModule.ipAdress = GetValueFromIniFile(strFilePath, "示波器", "IP地址");
                strPort = GetValueFromIniFile(strFilePath, "示波器", "端口号");
                CGloabal.g_InstrScopeModule.port = int.Parse(strPort);

                CGloabal.g_InstrMultimeterModule.ipAdress = GetValueFromIniFile(strFilePath, "万用表", "GPIB地址");
                strPort = GetValueFromIniFile(strFilePath, "万用表", "GPIB地址");
                CGloabal.g_InstrMultimeterModule.port = int.Parse(strPort);
            }
            else
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.error, "APP_INFOS.ini文件不存在！");
            }
        }
        /*************************************************
         * 函数原型：private void StoreAlltheInstumentsParas2IniFile()
         * 函数功能：将仪器的参数信息保存在ini文件中
         * 输入参数：
         * 输出参数：
         * 创 建 者：yzx
         * 创建日期：2016.7.23
         * 修改说明：
        */
        private static void StoreAlltheInstumentsParas2IniFile()
        {
            string strFileName;
            string strFilePath;

            //获取ini文件的相对路径
            strFilePath = System.Windows.Forms.Application.StartupPath + "\\APP_INFOS.ini";
            if (File.Exists(strFilePath))//先判断INI文件是否存在
            {
                strFileName = Path.GetFileNameWithoutExtension(strFilePath);//获得ini文件名

                WriteValueToIniFile(strFilePath, "电源", "IP地址", CGloabal.g_InstrPowerModule.ipAdress);
                WriteValueToIniFile(strFilePath, "电源", "端口号", CGloabal.g_InstrPowerModule.port.ToString());

                WriteValueToIniFile(strFilePath, "示波器", "IP地址", CGloabal.g_InstrScopeModule.ipAdress);
                WriteValueToIniFile(strFilePath, "示波器", "端口号", CGloabal.g_InstrScopeModule.port.ToString());

                WriteValueToIniFile(strFilePath, "万用表", "IP地址", CGloabal.g_InstrMultimeterModule.ipAdress);
                WriteValueToIniFile(strFilePath, "万用表", "端口号", CGloabal.g_InstrMultimeterModule.port.ToString());

            }
            else
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.error, "APP_INFOS.ini文件不存在！");
            }
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
        public static int ConnectSpecificInstrument(string strInstruName, string resourceName)
        {
            int error = 0;
            if (strInstruName == "电源")
            {
                if (CGloabal.g_InstrPowerModule.nHandle == 0)
                {
                    error = Power_Driver.Init(resourceName, ref CGloabal.g_InstrPowerModule.nHandle, strError);
                    if (error < 0)
                    {
                        CGloabal.g_InstrPowerModule.bInternet = false;
                        CCommonFuncs.ShowHintInfor(eHintInfoType.error, CGloabal.g_InstrPowerModule.strInstruName + "连接失败");
                    }
                    else
                    {

                        CGloabal.g_InstrPowerModule.bInternet = true;
                    }
                }
                else
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.hint, CGloabal.g_InstrPowerModule.strInstruName + "已经处于连接状态");
                }

            }
            else if (strInstruName == "示波器")
            {
                if (CGloabal.g_InstrScopeModule.nHandle == 0)
                {
                    error = Oscilloscope_Driver.Init(resourceName, ref CGloabal.g_InstrScopeModule.nHandle, ref strError);
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

            }
            else if (strInstruName == "万用表")
            {
                if (CGloabal.g_InstrMultimeterModule.nHandle == 0)
                {
                    error = Multimeter_Driver.Connect_Multimeter(resourceName);
                    if (error < 0)
                    {
                        CGloabal.g_InstrPowerModule.bInternet = false;
                        CCommonFuncs.ShowHintInfor(eHintInfoType.error, CGloabal.g_InstrMultimeterModule.strInstruName + "连接失败");
                    }
                    else
                    {

                        CGloabal.g_InstrPowerModule.bInternet = true;
                    }
                }
                else
                {
                    CCommonFuncs.ShowHintInfor(eHintInfoType.hint, CGloabal.g_InstrPowerModule.strInstruName + "已经处于连接状态");
                }
            }
            else
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.error, "错误的仪器名");
                return -1;
            }

            return error;

        }




        /******************************************************************************************
        * 函数原型：CloseSpecificInstrument(string strInstruName)
        * 函数功能：断开指定的仪器的网络连接
        * 输入参数：strInstruName，仪器名字
        * 输出参数：
        * 创 建 者：yzx
        * 创建日期：2016.7.27
        * 修改说明：
        * */
        public static int CloseSpecificInstrument(string strInstruName)
        {
            int error = 0;
            if (strInstruName == "电源")
            {
                if (CGloabal.g_InstrPowerModule.nHandle > 0)
                {
                    error = Power_Driver.Close(CGloabal.g_InstrPowerModule.nHandle, 0, strError);
                    if (error < 0)
                    {
                        CCommonFuncs.ShowHintInfor(eHintInfoType.error, "电源断开失败");
                    }
                    else//断开成功，要将此时的连接状态更新到仪器参数中
                    {
                        CGloabal.g_InstrPowerModule.bInternet = false;
                    }
                }

            }
            else if (strInstruName == "示波器")
            {
                if (CGloabal.g_InstrScopeModule.nHandle > 0)
                {
                    error = Oscilloscope_Driver.Close(CGloabal.g_InstrScopeModule.nHandle, 0, ref strError);
                    if (error < 0)
                    {
                        CCommonFuncs.ShowHintInfor(eHintInfoType.error, "示波器断开失败");

                    }
                    else//断开成功，要将此时的连接状态更新到仪器参数中
                    {
                        CGloabal.g_InstrScopeModule.bInternet = false;
                    }
                }
            }
            else if (strInstruName == "万用表")
            {
                if (CGloabal.g_InstrPowerModule.nHandle > 0)
                {
                    error = Multimeter_Driver.Close();
                    if (error < 0)
                    {

                        CCommonFuncs.ShowHintInfor(eHintInfoType.error, "信号发生器断开失败");
                    }
                    else//断开成功，要将此时的连接状态更新到仪器参数中
                    {
                        CGloabal.g_InstrPowerModule.bInternet = false;
                    }
                }
            }
            else
            {
                CCommonFuncs.ShowHintInfor(eHintInfoType.error, "错误的仪器名");
                return -1;
            }
            return error;
        }

        /******************************************************************************************
         * 函数原型：int SaveInputNetInforsToIniFile(string strInstruName,string strIP,UInt32 port )
         * 函数功能：当用户在界面上正确连接设备后，要将该此时的IP地址和端口号保存到ini文件中并更新仪器的参数信息
         *            这个函数只有在仪器连接成功后，才能调用。
         * 输入参数：strInstruName，仪器名字。
         * 输出参数：
         * 返 回 值：
         * 创 建 者：yzx
         * 创建日期：2016.7.27
         * 修改说明：
         * */
        public static int SaveInputNetInforsToIniFile(string strInstruName, string strIP, UInt32 port)
        {
            string strFilePath;
            //获取ini文件的相对路径
            strFilePath = System.Windows.Forms.Application.StartupPath + "\\APP_INFOS.ini";
            if (File.Exists(strFilePath))//先判断INI文件是否存在
            {
                if (strInstruName == "电源")
                {
                    //保存到ini文件
                    CDll.WriteValueToIniFile(strFilePath, "电源", "IP地址", strIP);
                    CDll.WriteValueToIniFile(strFilePath, "电源", "端口号", port.ToString());
                    //更新当前仪器的参数信息
                    CGloabal.g_InstrPowerModule.ipAdress = strIP;
                    CGloabal.g_InstrPowerModule.port = (int)port;
                    CGloabal.g_InstrPowerModule.bInternet = true;
                }
                else if (strInstruName == "示波器")
                {
                    //保存到ini文件
                    CDll.WriteValueToIniFile(strFilePath, "示波器", "IP地址", strIP);
                    CDll.WriteValueToIniFile(strFilePath, "示波器", "端口号", port.ToString());
                    //更新当前仪器的参数信息
                    CGloabal.g_InstrScopeModule.ipAdress = strIP;
                    CGloabal.g_InstrScopeModule.port = (int)port;
                    CGloabal.g_InstrScopeModule.bInternet = true;
                }
                else if (strInstruName == "万用表")
                {
                    //保存到ini文件
                    CDll.WriteValueToIniFile(strFilePath, "万用表", "IP地址", strIP);
                   
                    //更新当前仪器的参数信息
                    CGloabal.g_InstrMultimeterModule.ipAdress = strIP;
                    CGloabal.g_InstrMultimeterModule.port = (int)port;
                    CGloabal.g_InstrMultimeterModule.bInternet = true;
                }               
                else
                {
                    MessageBox.Show("错误的仪器名", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return -1;
                }
            }
            return 0;
        }


        const int BUFF_SIZE = 512;

        static int ScopeDefaultRM;
        static int PowerDefaultRM;
        static int MultimeterDefaultRM;





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

        /*************************************************
    函数原型：MultimeterModule_Init (ViChar resourceName[],HANDLE *pnHandle,ViString pErrMsg)  
    函数功能：万用表模块初始化
    输入参数：
    输出参数：
    返 回 值：
    */
        public static int MultimeterModule_Init(string resourceName, int nSimulateFlag, ref int pnHandle, ref string pErrMsg)
        {
            int error = 0;

            //检测是否处于模拟状态
            if (nSimulateFlag == 1)
                return 0;


            if ((error = visa32.viOpenDefaultRM(out MultimeterDefaultRM)) < 0)
                return error;

            if ((error = visa32.viOpen(MultimeterDefaultRM, resourceName, 0, 0, out pnHandle)) < 0)
            {
                visa32.viClose(ScopeDefaultRM);
                pnHandle = 0;
                pErrMsg = "万用表模块初始化失败！";
                return error;
            }

            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_TERMCHAR_EN, visa32.VI_TRUE);//终止符使能
            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_SEND_END_EN, visa32.VI_TRUE);//终止符使能	
            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_TERMCHAR, 0xA);//终止符设置0xA

            visa32.viSetBuf(pnHandle, visa32.VI_READ_BUF, 500);//RECVMAXLEN+4

            visa32.viSetAttribute(pnHandle, visa32.VI_ATTR_TMO_VALUE, 2000); //超时2000ms

            return 0;
        }

        /// <summary>
        /// 数据保存到Excel中
        /// </summary>
        /// <param name="LRModel"></param>
        /// <returns></returns>
        public static int DataSaveExcel(List<RecordModel> LRModel)
        {
            var fileNa = System.DateTime.Now.Hour;
            string fileNa1 = "";
            if (fileNa >= 00 && fileNa <= 8)
            {
                fileNa1 = System.DateTime.Now.ToString("yyyyMMdd") + "08";
            }
            else if(fileNa > 8 && fileNa <= 16)
            {
                fileNa1 = System.DateTime.Now.ToString("yyyyMMdd") + "16";
            }
            else if (fileNa > 16 && fileNa < 24)
            {
                fileNa1 = System.DateTime.Now.ToString("yyyyMMdd") + "24";
            }

           

            string filename = Application.StartupPath + @"\Report\Result"+ fileNa1 + ".xls";
            string connstr = @"Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + filename + ";Extended Properties='Excel 8.0;HDR=Yes'";//这个链接字符串是excel2003的
            OleDbConnection oleConn = new OleDbConnection(connstr);
            try
            {
                oleConn.Open();

                string sqlStr;
                DataTable dt = oleConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables_Info, null);
                bool existTable = false;

                foreach (DataRow dr in dt.Rows)//检查是否有信息表
                {
                    if (dr["TABLE_NAME"].ToString() == "Sheet1$")//要加个$号
                        existTable = true;
                }
                if (!existTable)
                {
                    sqlStr = @"create table Sheet1( 时间 varchar(30),  电压1 varchar(30), 电流1 varchar(30), 电压2 varchar(30), 电流2 varchar(30),  速率 varchar(30), 通道11发  varchar(30),  通道11收 varchar(30), 通道11错误数  varchar(30),"
                   + @" 通道21发 varchar(30),通道21收 varchar(30),  通道21错误数 varchar(30), 通道12发 varchar(30),  通道12收 varchar(30), 通道12错误数  varchar(30),通道22发 varchar(30), 通道22收 varchar(30),  通道22错误数 varchar(30),"
                   + @"通道13发 varchar(30),  通道13收 varchar(30),通道13错误数 varchar(30),通道23发 varchar(30),通道23收  varchar(30), 通道23错误数 varchar(30),通道14发  varchar(30), 通道14收 varchar(30),"
                   + @"通道14错误数 varchar(30),通道24发 varchar(30),通道24收  varchar(30), 通道24错误数 varchar(30))";

                   OleDbCommand oleCmd = new OleDbCommand(sqlStr, oleConn);
                    oleCmd.ExecuteNonQuery();
                }

                //下面的代码用OleDbCommand的parameter添加参数
                int aa = 0;
                for (int i = 0; i < LRModel.Count; i++)
                {
                    sqlStr = "insert into Sheet1 values('" + LRModel[i].curDate + "','" + LRModel[i].volVal1 + "','" + LRModel[i].eleVal1 + "','" + LRModel[i].volVal2 + "','" + LRModel[i].eleVal2 + "','" + LRModel[i].rate + "','" + LRModel[i].sendVal11 + "','" + LRModel[i].receivVal11 + "','" + LRModel[i].errorVal11
                        + @"','" + LRModel[i].sendVal21 + "','" + LRModel[i].receivVal21 + "','" + LRModel[i].errorVal21 + "','" + LRModel[i].sendVal12 + "','" + LRModel[i].receivVal12 + "','" + LRModel[i].errorVal12 + "','" + LRModel[i].sendVal22
                        + @"','" + LRModel[i].receivVal22 + "','" + LRModel[i].errorVal22 + "','" + LRModel[i].sendVal13 + "','" + LRModel[i].receivVal13 + "','" + LRModel[i].errorVal13 + "','" + LRModel[i].sendVal23 + "','" + LRModel[i].receivVal23 
                        + @"','" + LRModel[i].errorVal23 + "','" + LRModel[i].sendVal14 + "','" + LRModel[i].receivVal14 + "','" + LRModel[i].errorVal14 + "','" + LRModel[i].sendVal24 + "','" + LRModel[i].receivVal24 + "','"+ LRModel[i].errorVal24 + "')";
                    OleDbCommand Cmd = new OleDbCommand(sqlStr, oleConn);
                   aa=  Cmd.ExecuteNonQuery();
                    if (aa <= 0) {
                        break;
                    }
                }
                if (aa <= 0)
                {
                    MessageBox.Show("添加数据失败！");
                    return -1;
                }
                else {
                    return aa;
                }               
            }
            catch (Exception te)
            {
                MessageBox.Show(te.Message);
                return -1;
            }
            finally
            {
                oleConn.Close();
            }

        }


        /// <summary>
        /// 基本功能测试发送数据
        /// </summary>
        /// <param name="chipNum">芯片</param>
        /// <param name="pathNum">通道</param>
        /// <param name="typeName">类型</param>
        public static void BaseTestSendData(int chipNum, string basePath, int sendNum)
        {
            byte[] cmdByte = new byte[10] { 0xAA, 0, 0, 0, 0, 0, 0, 0, 0, 0xBB };
            cmdByte[1] = (byte)4;
            cmdByte[2] = (byte)chipNum;
            cmdByte[4] = (byte)sendNum;
            if (basePath == "通道1")
            {
                cmdByte[3] = (byte)0;
            }
            else if (basePath == "通道2")
            {
                cmdByte[3] = (byte)1;
            }
            else if (basePath == "通道3")
            {
                cmdByte[3] = (byte)2;
            }
            else if (basePath == "通道4")
            {
                cmdByte[3] = (byte)3;
            }

            int error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);
            if (error < 0)
            {
                MessageBox.Show("芯片"+chipNum+""+basePath+"发送失败");
                return;
            }
        }

        /// <summary>
        /// 基本功能测试读取数据
        /// </summary>
        /// <param name="chipNum">芯片</param>
        /// <param name="pathNum">通道</param>
        /// <param name="typeName">类型</param>
        public static byte[] BaseTestReadData(int chipNum, string basePath, string typeName)
        {
            byte[] cmdByte = new byte[10] { 0xAA, 0, 0, 0, 0, 0, 0, 0, 0, 0xBB };
            cmdByte[1] = (byte)6;
            cmdByte[2] = (byte)chipNum;

            if (basePath == "通道1")
            {
                cmdByte[3] = (byte)0;
            }
            else if (basePath == "通道2")
            {
                cmdByte[3] = (byte)1;
            }
            else if (basePath == "通道3")
            {
                cmdByte[3] = (byte)2;
            }
            else if (basePath == "通道4")
            {
                cmdByte[3] = (byte)3;
            }

            int error = 0;
            switch (typeName)
            {
                case "BASE":
                    cmdByte[1] = (byte)5;
                    error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);                   
                    break;
                case "LSR":
                    cmdByte[4] = (byte)5;
                    error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);                   
                    break;
                case "IIR":
                    cmdByte[4] = (byte)2;
                    error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);                   
                    break;
                case "ARM":
                    cmdByte[1] = (byte)7;
                    error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);                  
                    break;
            }
            if (error < 0)
            {
                MessageBox.Show("芯片" + chipNum + "" + basePath + "读取失败");
            }

            Thread.Sleep(500);
           
            List<byte> buffer = new List<byte>(4096);
            int length;
            byte[] ReceiveBytes = new byte[24];

            CGloabal.g_bIsComRecvedDataFlag = true;  //串口是否收到数据
            length = CGloabal.g_serialPorForUUT.BytesToRead;
            byte[] Receivebuf = new byte[24];
            CGloabal.ReadCom(CGloabal.g_serialPorForUUT, Receivebuf, 24);
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
                    return ReceiveBytes;
                }
                else //帧头不正确时，记得清除
                {
                    buffer.RemoveAt(0);
                }
            }

            return ReceiveBytes;
        }


        /// <summary>
        /// 芯片设置
        /// </summary>
        /// <param name="CHIPMODEL"></param>
        /// <returns></returns>
        public static int ChipSet(ChipModel CHIPMODEL)
        {
            int error = 0;
            Byte[] cmdByte = new Byte[10] { 0XAA, 0X00 , 0X00 , 0X00 , 0X00 , 0X00 , 0X00 , 0X00 , 0X00 , 0XBB };
    
            cmdByte[2] = (byte)CHIPMODEL.chipSelect;
            cmdByte[3] = (byte)CHIPMODEL.pathSelect;

            //设置DUT波特率
            cmdByte[1] = 0x03;
            cmdByte[4] = 96;
            error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);
            Thread.Sleep(100);


            cmdByte[1] = 0x08;
            //发送中断、接收中断、接收缓存中断
            cmdByte[4] = 01;
            cmdByte[5] = 0x00;

            cmdByte[5] += ValueReturn(CHIPMODEL.receiveInterrupt, "接收中断");
            cmdByte[5] += ValueReturn(CHIPMODEL.sendInterrupt, "发送中断");
            cmdByte[5] += ValueReturn(CHIPMODEL.receiveCache, "接收缓存中断");

            error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);
            if (error < 0)
            {
                return -3;
            }
            Thread.Sleep(100);

            cmdByte[4] = 02;
            //FIFO使能，DMA模式，接收FIFO触发器，发送触发器
            cmdByte[5] = 0x00;
            cmdByte[5] += ValueReturn(CHIPMODEL.FIFOSelect, "FIFO使能");
            cmdByte[5] += ValueReturn(CHIPMODEL.DMAPattern, "DMA模式");
            cmdByte[5] += ValueReturn(CHIPMODEL.receiveFIFO, "接收触发器");
            cmdByte[5] += ValueReturn(CHIPMODEL.sendTarget, "发送触发器");

            error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);
            if (error < 0)
            {
                return -2;
            }

            //设置奇偶校验，停止位，字长
            cmdByte[4]  = 3;         
            cmdByte[5] = 0x00;
            cmdByte[5] += ValueReturn(CHIPMODEL.parityCheck, "奇偶校验");
            cmdByte[5] += ValueReturn(CHIPMODEL.stopBit, "停止位");
            cmdByte[5] += ValueReturn(CHIPMODEL.byteLength, "字长");
            error = CGloabal.WriteToCom(CGloabal.g_serialPorForUUT, cmdByte, 10);
            if (error < 0) {
                return -1;
            }
            Thread.Sleep(100);



            Thread.Sleep(100);



            return error;
        }


        private static byte ValueReturn(string  data, string typeName)
        {
            byte byVal=0x00;

            switch (typeName) {
                case "奇偶校验":
                    if (data == "无")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "奇校验")
                    {
                        byVal = 0x08;
                    }
                    else if (data == "偶校验")
                    {
                        byVal = 0x18;
                    }
                    break;
                case "停止位":
                    if (data == "1")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "2")
                    {
                        byVal = 0x04;
                    }
                    break;
                case "字长":

                    if (data == "5")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "6")
                    {
                        byVal = 0x01;
                    }
                    else if (data == "7")
                    {
                        byVal = 0x02;
                    }
                    else if (data == "8")
                    {
                        byVal = 0x03;
                    }
                    break;
                case "DMA模式":
                    if (data == "0")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "1")
                    {
                        byVal = 0x08;
                    }
                    break;
                case "接收触发器":
                    if (data == "1")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "2")
                    {
                        byVal = 0x40;
                    }
                    else if (data == "3")
                    {
                        byVal = 0x80;
                    }
                    else if (data == "4")
                    {
                        byVal = 0xC0;
                    }
                    break;
                case "发送触发器":
                    if (data == "1")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "2")
                    {
                        byVal = 0x10;
                    }
                    else if (data == "3")
                    {
                        byVal = 0x20;
                    }
                    else if (data == "4")
                    {
                        byVal = 0x30;
                    }
                    break;
                case "FIFO使能":
                    if (data == "不使能")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "使能")
                    {
                        byVal = 0x01;
                    }
                    break;
                case "接收中断":
                    if (data == "不使能")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "使能")
                    {
                        byVal = 0x04;
                    }
                    break;
                case "发送中断":
                    if (data == "不使能")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "使能")
                    {
                        byVal = 0x02;
                    }
                    break;
                case "接收缓存中断":
                    if (data == "不使能")
                    {
                        byVal = 0x00;
                    }
                    else if (data == "使能")
                    {
                        byVal = 0x01;
                    }
                    break;
                default:
                    byVal = 0x00;
                    break;
            }

            return byVal;
        } 
    }
}
