using System;
using System.Data;
using System.Configuration;
using System.Text;

namespace CIMSClass
{
    public abstract class StringUtil
    {
        public StringUtil() { }

        /// <summary>
        /// 輸入的數據轉換成正確的日期



        /// </summary>
        /// <param name="Str">yyyy/MM/dd格式的字符串</param>
        /// <returns>變換后的日期</returns>
        public static String ReplaceDate(String Str)
        {
            DateTime dt = new DateTime();
            //(輸入的日期形式是yyyy/MM/dd HH:MM)
            String[] Strs = Str.Split('/');

            Int32 Year = Int32.Parse(Strs[0]);
            Int32 Month = Int32.Parse(Strs[1]);
            Int32 Day = Int32.Parse(Strs[2]);

            //月份為12以上的等同於12
            if (12 < Month)
            {
                Month = 12;
            }
            //輸入某月天數如果大於該月的固有天數，則天數為該月的最後一天



            if (DateTime.DaysInMonth(Year, Month) < Day)
            {
                dt = new DateTime(Year, Month, DateTime.DaysInMonth(Year, Month));
            }

            return dt.ToString("yyyy/MM/dd");
        }

        //************************************************************************
        /// <summary>
        /// 是否為空
        /// </summary>
        /// <param name="strText">判斷的字符</param>
        /// <returns>true:為空|false:不為空</returns>
        //************************************************************************
        public static bool IsEmpty(string strText)
        {
            if (strText == null || strText.Trim().Length == 0)
                return true;
            return false;
        }

        /// <summary>
        /// 刪除字符串中的單引號
        /// </summary>
        /// <param name="Str">字符串對象</param>
        /// <returns>刪除單引號后的字符串</returns>
        public static String RemoveSct(String Str)
        {
            return Str.Replace("'", "");
        }

        /// <summary>
        /// 將數值類型字符串中的千分位刪除



        /// </summary>
        /// <param name="Str">數值類型字符串對象</param>
        /// <returns>刪除千分位后的字符串</returns>
        public static String SpecHDelComma(String Str)
        {
            if ((Str == null) || ("".Equals(Str)))
            {
                return Str;
            }

            return Str.Replace(",", "");
        }

        /// <summary>
        /// 為整數增加千分位
        /// </summary>
        /// <param name="Str"></param>
        /// <returns></returns>
        public static String SpecAddComma(String Str)
        {
            // 12345789->12,345,789

            if ((Str == null) || ("".Equals(Str)))
            {
                return Str;
            }

            String value;
            String minus;
            if ("-".Equals(Str.Substring(0, 1)))
            {
                value = Str.Substring(1, Str.Length - 1);
                minus = "-";
            }
            else
            {
                value = Str;
                minus = "";
            }

            if ("".Equals(value))
            {
                return minus + "0";
            }

            value = value.Replace(",", "");
            int pos = 0;
            String newValue = "";
            for (int i = value.Length; i > 0; i--)
            {
                if (pos == 3)
                {
                    newValue = "," + newValue;
                    pos = 1;
                }
                else
                {
                    pos = pos + 1;
                }
                newValue = value.Substring(i - 1, 1) + newValue;
            }

            return minus + newValue;
        }

        /// <summary>
        /// 為浮點數增加千分位



        /// </summary>
        /// <param name="Str"></param>
        /// <returns></returns>
        public static String SpecDecimalAddComma(String Str)
        {
            if ((Str == null) || ("".Equals(Str)))
            {
                return Str;
            }
            String sDecimal;
            String sInteger;
            int indexDot = Str.IndexOf('.');
            if (indexDot != -1)
            {
                sDecimal = Str.Substring(indexDot);
                sInteger = Str.Substring(0, indexDot);
            }
            else
            {
                sDecimal = "";
                sInteger = Str;
            }
            String sValue = SpecAddComma(sInteger);
            if (sValue.Equals(""))
            {
                sValue = "0";
            }

            return sValue + sDecimal;
        }

        /// <summary>
        /// 根據默認的編碼取得字符串的字節長度
        /// </summary>
        /// <param name="text">字符串對象</param>
        public static int GetByteLength(string text)
        {
            return System.Text.Encoding.Default.GetBytes(text).Length;
        }

        /// <summary>
        /// 根據字節長度截取字符串



        /// </summary>
        /// <param name="srcString">字符串對象</param>
        /// <param name="bytes">要截取的字節長度</param>
        /// <returns></returns>
        public static string GetSubstring(string srcString, Int32 bytes)
        {
            if (string.IsNullOrEmpty(srcString) ||
                bytes <= 0)
            {
                return srcString;
            }

            Int32 srcLength = GetByteLength(srcString);

            if (bytes > srcLength)
            {
                bytes = srcLength;
            }

            StringBuilder returnString = new StringBuilder();
            int count = 0;
            for (int j = 0; j < srcString.Length; j++)
            {
                count += GetByteLength(srcString.Substring(j, 1));

                if (count > bytes)
                {
                    break;
                }
                else
                {
                    returnString.Append(srcString[j]);
                }
            }

            return returnString.ToString();
        }

        /// <summary>
        /// 將對象轉換成Bool類型
        /// </summary>
        /// <param name="objInput"></param>
        /// <returns></returns>
        public static bool ToBoolean(object objInput)
        {
            bool returnResult = false;

            if (objInput == null)
            {
                return false;
            }

            string tempValue = objInput.ToString().ToLower();

            if (tempValue.Equals("1") ||
                tempValue.Equals("true") ||
                tempValue.Equals("yes"))
            {
                returnResult = true;
            }
            else
            {
                returnResult = false;
            }

            return returnResult;
        }

        /// <summary>
        /// 日期格式轉換(yyyymmdd -> yyyy/mm/dd)
        /// </summary>
        /// <param name="targetString">要變換的日期字符串</param>
        /// <returns>轉換后的日期字符串</returns>
        public static string DateStringTransform(string targetString)
        {
            if (targetString.Length < 8)
            {
                return targetString;
            }

            return targetString.Substring(0, 4) + "/" + targetString.Substring(4, 2) + "/" + targetString.Substring(6, 2);
        }

        /// <summary>
        /// 日期格式轉換(yyyymmddhhmmss -> yyyy/mm/dd hh:mm:ss)
        /// </summary>
        /// <param name="targetString">要變換的日期字符串</param>
        /// <returns>轉換后的日期字符串</returns>
        public static string DateTimeStringTransform(string targetString)
        {
            string returnString = "";

            if (targetString.Length < 14)
            {
                return targetString;
            }

            //日期部分
            returnString = targetString.Substring(0, 4) + "/" + targetString.Substring(4, 2) + "/" + targetString.Substring(6, 2);

            //時刻部分
            returnString += " " + targetString.Substring(8, 2) + ":" + targetString.Substring(10, 2) + ":" + targetString.Substring(12, 2);

            return returnString;
        }

        /// <summary>
        /// 將Null轉換成空串



        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ConvertNullToString(object obj)
        {
            string strRtn = "";
            if (obj == null)
            {
                return strRtn;
            }
            else
            {
                return obj.ToString();
            }
        }

        /// <summary>
        /// 將Null轉換成0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ConvertNullToZero(object obj)
        {
            if (obj == null)
            {
                return "0";
            }
            else if (obj.ToString().Trim().Equals(""))
            {
                return "0";
            }
            else
            {
                return obj.ToString();
            }
        }
        public static string GetSubstringByByte(string strReadLine, int begin, int length, out int nextBegin)
        {
            string strTemp1 = strReadLine.Substring(begin, length + length - StringUtil.GetByteLength(strReadLine.Substring(begin, length)));
            nextBegin = begin + strTemp1.Length;
            return strTemp1;
        }

        /// </summary>
        /// <param name="Str">字符串對象</param>
        /// <returns>去除換行符后的字符串</returns>
        /// <remarks>Legend 2017/09/26 添加</remarks>
        public static String RemoveNewLineChar(String Str)
        {
            if ((Str == null) || ("".Equals(Str)))
            {
                return Str;
            }

            return Str.Replace("\n", "").Replace("\r", "").Replace("%3d", "").Replace("%0a", "");
        }
    }
}