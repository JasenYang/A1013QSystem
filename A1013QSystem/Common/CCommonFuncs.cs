using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Forms.VisualStyles;

namespace A1013QSystem.Common
{
    //定义提示信息枚举体
    public enum eHintInfoType
    {
        hint=0,
        waring=1,
        error=2,
    }
    
    class CCommonFuncs
    {
        /*************************************************
       * 函数原型：void RangeCheck(ref byte nInputVal,byte nMinVal,byte nMaxVal)
       * 函数功能：对输入的数据进行范围限定。
       * 输入参数：
       * 输出参数：
       * 创 建 者：yzx
       * 创建日期：2016.7.23
       * 修改说明：
      */
        public static void RangeCheck(ref byte nInputVal, byte nMinVal, byte nMaxVal)
        {
            if (nInputVal < nMinVal)
            {
                nInputVal = nMinVal;
            }
            else if (nInputVal > nMaxVal)
            {
                nInputVal = nMaxVal;
            }
            else
            {
                ;
            }
        }

        /******************************************************************
         * 函数原型：void ShowHintInfor(eHintInfoType eInfoType, string strInfor)
         * 函数功能：提示信息
         * 输入参数：InfoType，提示的类型；strInfor，提示的内容
         * 输出参数：
         * 返 回 值：
         * 创 建 者：yzx
         * 创建日期：
         * 修改说明：
         * */
        public static void ShowHintInfor(eHintInfoType eInfoType, string strInfor)
        {
            switch (eInfoType)
            {
                case eHintInfoType.hint://提示类型的信息
                    MessageBox.Show(strInfor, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                    break;
                case eHintInfoType.waring://警告类型的信息
                    MessageBox.Show(strInfor, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                    break;
                case eHintInfoType.error://错误类型的信息
                    MessageBox.Show(strInfor, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                default:
                    break;      
            }
            
        }

    









    }
}
