using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Base64ToImage
{
    interface MyUnit
    {
        int Pos(string target, string str);
        string LeftStr(string str, int index);
        string RightStr(string str, int index);
        int Length(string str);
        int MyStrToInt(string str);
    }
    class Myutil1 : MyUnit
    {
        public int Pos(string target, string str)
        {
            return str.IndexOf(target);
        }

        public string LeftStr(string str, int index)
        {
            if (index == -1) return "";
            return str.Substring(0, index);  //获取从0，到index字符
        }
        public string RightStr(string str, int index)
        {

            return str.Substring(str.Length - index + 1); //获取右数 index -1 个字符
        }
        public int Length(string str)
        {
            return str.Length;
        }

        public int MyStrToInt(string str)
        {
            int Result = 0;
            try
            {
                // 引起异常的语句
                if (str != "")
                {
                    Result = int.Parse(str);
                }
            }
            catch
            {

            }
            return Result;
        }

    }

}
