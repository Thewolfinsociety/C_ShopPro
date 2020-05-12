using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections;

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

    class Myutil : MyUnit
    {
        static public string GetAttributeValue(XmlNode node, string attristring, string value, string defalut)
        {
            string result = "";
            XmlAttribute attri = node.Attributes[attristring];
            if (attri != null)
            {
                value = attri.Value;
                result = value;
                //int startrow = int.Parse(attri.Value);
            }
            if (value == "") result = defalut;
            return result;
        }

        static public int GetAttributeValue(XmlNode node, string attristring, int value, int defalut)
        {
            int result = 0;
            XmlAttribute attri = node.Attributes[attristring];
            if (attri != null)
            {
                value = int.Parse(attri.Value);
                result = value;
                //int startrow = int.Parse(attri.Value);
            }
            if (value == 0) result = defalut;
            return result;
        }

        static public float GetAttributeValue(XmlNode node, string attristring, float value, float defalut)
        {
            float result = 0;
            XmlAttribute attri = node.Attributes[attristring];
            if (attri != null)
            {
                value = float.Parse(attri.Value);
                result = value;
                //int startrow = int.Parse(attri.Value);
            }
            if (value == 0) result = defalut;
            return result;
        }
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

    class TExpress
    {
        struct myvariable
        {
            public string name, alias, value, opvalue, extra;
            public int mytype, remove;
            public bool change;
        }

        struct mysymbol
        {
            public string var1, var2, value, value2;
            public int flag, used;
        }
        private string mSubject;
        private int mDotNum;
        private ArrayList mSblList;
        private bool mIsSimple, mIsChangeVar;
        private string mResult;
        public ArrayList mVarList;
        public int mBHValue;

        public TExpress()
        {
            Console.WriteLine("TExpress对象已创建");
            mVarList = new ArrayList();
            mSblList = new ArrayList();

            mIsSimple = false;
            mBHValue = 18;
            mDotNum = -1;
            mIsChangeVar = true;
        }

        /*~TExpress() //析构函数
        {
            ClearVarList;
            ClearSblList;
            Console.WriteLine("对象已删除");
        }*/

        public void SetSubject(string str, bool recalc = false)
        {
            int mDotNum = -1;
        }

        
    }
}
