using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections;
using System.Text.RegularExpressions;

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
        public int Round(Single value)
        {

            return (int)(Math.Round(value));
        }
        public bool IsNumeric(char[] s)
        {
            bool result = false;
            int i = 0;
            char[] numArray = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.'};
            while (s[i] != '\0')
            {
                if (!numArray.Contains(s[i])) return result;
                i = i + 1;
            }
            result = true;
            return result;

        }
        public bool IsNumeric(string s)
        {
            Regex objNotNumberPattern = new Regex("[^0-9.-]");
            Regex objTwoDotPattern = new Regex("[0-9]*[.][0-9]*[.][0-9]*");
            Regex objTwoMinusPattern = new Regex("[0-9]*[-][0-9]*[-][0-9]*");
            String strValidRealPattern = "^([-]|[.]|[-.]|[0-9])[0-9]*[.]*[0-9]+$";
            String strValidIntegerPattern = "^([-]|[0-9])[0-9]*$";
            Regex objNumberPattern = new Regex("(" + strValidRealPattern + ")|(" + strValidIntegerPattern + ")");

            return !objNotNumberPattern.IsMatch(s) &&
                   !objTwoDotPattern.IsMatch(s) &&
                   !objTwoMinusPattern.IsMatch(s) &&
                   objNumberPattern.IsMatch(s);


        }
        public string IntToStr(int num)
        {
            return num.ToString();
        }
        public Single StrToFloat(string str)
        {
            Single result= 0;
            Single.TryParse(str, out result);
            return result;
        }
        
    }

    class TExpress : Myutil
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
   
        private bool mIsSimple, mIsChangeVar;
        private string mResult;
        //public ArrayList mVarList;
        List<myvariable> mVarList;
        List<mysymbol> mSblList;
        public int mBHValue;

        public TExpress()
        {
            Console.WriteLine("TExpress对象已创建");
            mVarList = new List<myvariable>();
            mSblList = new List<mysymbol>();

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
            int n = Pos(":", str);
            string s = "";
            if (n > 0)
            {
                s = RightStr(str, Length(str) - n);
                this.mDotNum = MyStrToInt(s);
                str = LeftStr(str, n - 1);
            }

            s = LeftStr(str, 1);
            if ((s == "+") || (s == "-")) str = "0" + str;

            if ((mSubject != str) || (recalc))
            {
                this.mSubject = str;
                mIsSimple = false;
            }
        }
        public void AddVariable(string name, string alias, string value, string opvalue, string extra, int mt)
        {
            bool mIsChangeVar = true;
            myvariable p; //
            for (int i = 0; i <= mVarList.Count; i++)
            {
                p = mVarList[i];
                if (p.name == name)
                {
                    p.value = value;
                    p.mytype = mt;
                    p.opvalue = opvalue;
                    p.extra = extra;
                    p.remove = 0;
                    return;
                }
                p = new myvariable();
                p.alias = alias;
                p.name = name;
                p.value = value;
                p.mytype = mt;
                p.opvalue = opvalue;
                p.extra = extra;
                p.remove = 0;
                mVarList.Add(p);
            }
        }

        public string ToValueString()
        {

            return (ToValueFloat().ToString());

        }
        public int ToValueInt()
        {
            return Round(ToValueFloat());
        }

        public Single ToValueFloat()
        {
            int Result = 0;
            if (! mIsSimple) mResult = ToSimpleExpress2(true);
            return 0;

        }
        private void ClearSblList()
        {
            mysymbol p;
            for (int i = 0; i <= mSblList.Count - 1; i++)
            {
                p = mSblList[i];
                p.var1 = "";
                p.var2 = "";
                p.value = "";
                p.value2 = "";
            }
            mSblList.Clear();
            mIsChangeVar = true;
        }
        private void Parse(string str)
        {
            ClearSblList();
            if (str == null) return;
            char[] buf = str.ToCharArray();
            int len = Length(str);
            int i = 0;
            int j = 0;
            char[] tmp = new char[64];
            char[] OpSymbolArray = { '+', '-', '*', '/', '(', ')'};
            mysymbol p;
            myvariable v;
            while (i < len)
            {
                if (buf[i] == ' ' || buf[i] =='\r' || buf[i] == '\n')
                {
                    i = i + 1;
                    continue;
                }
                if (OpSymbolArray.Contains(buf[i]))
                {
                    p = new mysymbol();
                    tmp[0] = buf[i];
                    p.value = new string(tmp);
                    p.flag = 0;
                    mSblList.Add(p);
                    i = i + 1;
                    continue;
                }

                j = i;
                while (j <= len)
                {
                    if ((OpSymbolArray.Contains(buf[i])) || (j == len))
                    {
                        p = new mysymbol();
                        p.value = new string(tmp);
                        p.flag = 1;
                        if (!IsNumeric(tmp)) p.flag = 2;
                        mSblList.Add(p);
                        i = j;
                        break;
                    }
                    tmp[j - i] = buf[j];
                    j = j + 1;
                }
            }
            for (i = 0; i <= mSblList.Count - 1; i++)
            {
               
                p = mSblList[i];
                for (j = 0; j <= mVarList.Count - 1; j++)
                {
                    
                    v = mVarList[j];
                    if (v.value == "")  continue;
                    if (p.value == v.name)
                    {
         
                        p.value = v.value;
                        if (v.value == "-10001")       //特殊变量
                        {
                            p.value = IntToStr(mBHValue);
                        }
                        break;
                    }
                }
            }         
        }
        private int RemovemBrackets(int start, string s, bool neg)
        {
            int i = 0, j = 0;
            mysymbol p, p0;
            string ss = "";
            bool tneg = false;
            int Result = start;

            p = mSblList[start];
            p.flag = -1;
            tneg = false;
            i = start + 1;
            while (i < mSblList.Count)
            {
                p = mSblList[i];
                if (p.flag < 0) 
                {
                    i = i + 1;
                    continue;
                }
                if ((p.flag == 0) && (p.value == "+") || (p.value == "-"))
                {
                    if ((neg) && (p.value == "+")) p.value = "-";
                    if ((neg) && (p.value == "-")) p.value = "+";
                    if (p.value == "+") tneg = false;
                    if (p.value == "-") tneg = true;
                }
                if ((p.flag == 0) && (p.value == "("))
                {
                    ss = "";
                    j = i - 1;
                    while (j >= 0)
                    {
                        p0 = mSblList[j];
                        if (p0.flag >= 0)
                        {
                            if ((p0.flag == 0) && (p0.value != "*") && (p0.value != "/")) break;
                            ss = p0.value + ss;
                            p0.flag = -2;
                        }
                        j= j - 1;
                    }
                    i = RemovemBrackets(i, s + ss, tneg);
                    continue;
                }
                if ((p.flag == 0) && (p.value == ")"))
                {
                    p.flag = -1;
                    ss = "";
                    j = i - 1;
                    while (j < mSblList.Count)
                    {
                        p0 = mSblList[j];
                        if (p0.flag >= 0)
                        {
                            if ((p0.flag == 0) && (p0.value != "*") && (p0.value != "/")) break;
                            ss = p0.value + ss;
                            p0.flag = -2;
                        }
                        j = j + 1;
                    }
                    Result = j - 1;
                    for (i = start + 1; i <= j - 1; i++)
                    {
                        p = mSblList[i];
                        if (p.flag <= 0) continue;
                        p.value = p.value + ss;
                        return Result;
                    }
                }
                if (p.flag > 0)
                {               
                    ss = "";
                    j = i + 1;
                    while (j < mSblList.Count)
                    {
                        p0 = mSblList[j];
                        if (p0.flag >= 0)
                        {
                            if ((p0.flag == 0) && (p0.value != "*") && (p0.value != "/")) break;
                            ss = p0.value + ss;
                            p0.flag = -2;
                        }
                        j = j + 1;
                    }
                    p.value = p.value + ss + s;
                    p.flag = 3;
                }
                i = i + 1;
            }
            return Result;
        }
        public Single MyStrToFloat(string str)
        {
            Single Result = 0;
            if (IsNumeric(str)) Result = StrToFloat(str);
            return Result;
        }
        public int MyStrToInt(string str, int v)
        {
            int Result = v;
            if (IsNumeric(str)) Result = (int)Math.Round(MyStrToFloat(str));
            return Result;
        }
        private void GetVariantValueSign(string str, ref mysymbol p)
        {
            char[] OpSymbolArray = { '.', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            int i = 0, j = 0, len = 0, flag = 0;
            string lastsign = "";
            char[] tmp = new char[64];
            char[] buf = new char[512];
            p.value = "";
            p.var1 = "";
            p.var2 = "";
            p.value2 = "";
            lastsign = "*";
            buf = str.ToCharArray();
            len = Length(str);
            i = 0;
            while (i < len)
            {
                if (buf[i] == '*' || buf[i] == '/')
                {

                    tmp[0] = buf[i];
                    lastsign = new string(tmp);
                    i = i + 1;
                    continue;
                }
                j = i;
                flag = 1;
                while (j <= len)
                {
                    if ((buf[i] == '*' || buf[i] == '/') || (j == len))
                    {

                        if (flag == 1)                 //数值
                        {
                            if (lastsign == "*")
                            {
                                if (p.value == "") p.value = new string(tmp);
                                else
                                    p.value = IntToStr(MyStrToInt(p.value, 1) * MyStrToInt(new string(tmp), 1));
                            }
                            if (lastsign == "/")
                            {
                                if (p.value2 == "") p.value2 = new string(tmp);
                                else
                                    p.value2 = IntToStr(MyStrToInt(p.value2, 1) * MyStrToInt(new string(tmp), 1));
                            }
                        }

                        if (flag == 2)                 //数值
                        {
                            if (lastsign == "*")
                            {
                                if (p.var1 == "") p.var1 = new string(tmp);
                                else
                                    p.var1 = p.var1 + '*' + new string(tmp);
                            }
                            if (lastsign == "/")
                            {
                                if (p.var2 == "") p.var2 = new string(tmp);
                                else
                                    p.var2 = p.var2 + '*' + new string(tmp);
                            }
                        }
                        i = j;
                        break;
                    }
                    if (!(OpSymbolArray.Contains(buf[j]))) flag = 2;
                    tmp[j - i] = buf[j];
                    j = j + 1;
                }
            }
            if ((p.value != "") && (p.value2 != ""))
            {
                i = MyStrToInt(p.value, 1);
                j = MyStrToInt(p.value2, 1);
                if ((i % j) == 0)
                {
                    p.value = IntToStr(i / j);
                    p.value2 = "";
                }
            }
            if ((p.var1 != "1") && (p.var2 != "")) p.value = "";
            if (p.value == "1") p.value2 = "";
            if ((p.var1 == "") && (p.var2 == "") && (p.value == "") && (p.value2 == ""))  p.value = str;
        }
        public void SimpleParse(string str)
        {
        
            ClearSblList();
            if (str == null) return;
            char[] buf = str.ToCharArray();
            int len = Length(str);
            int i = 0;
            int j = 0;
            int flag = 0;
            char[] tmp = new char[64];
            char[] OpSymbolArray = { '.', '*', '/', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
            mysymbol p;
            myvariable v;
            while (i < len)
            {
                if (buf[i] == '+' || buf[i] == '-')
                { 
              
                    p = new mysymbol();
                    tmp[0] = buf[i];
                    p.value = new string(tmp);
                    p.flag = 0;
                    mSblList.Add(p);
                    i = i + 1;
                    continue;
                }

                j = i;
                while (j <= len)
                {
                    if ((buf[i] == '+' || buf[i] == '-') || (j == len))
                    {
                        p = new mysymbol();
                        p.value = new string(tmp);
                        GetVariantValueSign(p.value, ref p);
                        p.flag = flag;
                        
                        mSblList.Add(p);
                        i = j;
                        break;
                    }
                    if (! OpSymbolArray.Contains(buf[i])) p.flag = 2;
                    
                    tmp[j - i] = buf[j];
                    j = j + 1;
                }
            }
        }
        private void ComSymbolList()
        {
            int i, j, a1, a2, b1, b2;
            mysymbol p1, p2, ps1, ps2;
            mysymbol symbol;
            ps1 = new mysymbol();
            ps1.flag = 0;
            ps1.value = "+";
            symbol = ps1;
            for (i = 0; i <= mSblList.Count - 1; i++)
            {
                p1 = mSblList[i];
                p1.used = 0;
            }
            for (i = 0; i <= mSblList.Count - 1; i++)
            {
                p1 = mSblList[i];
                if (p1.used == 1) continue;
                if (p1.flag == 0)
                {
                    ps1 = p1;
                    continue;
                }
                if (p1.flag != 2) continue;
            }


        }
        private string ToSimple(bool isfloat)
        {
            int i = 0, t = 0;
            mysymbol p1, p2;
            string s1, s2, sign;
            Single v1, v2;
            string Result = "";
            s1 = "";
            s2 = "";
            v1 = 0;
            v2 = 0;
            i = 1;
            ComSymbolList;                      //合并相同变量的表达式
            return Result;
        }
        private string ToSimpleExpress2(bool isfloat)
        {
            int i = 0, j = 0;
            mysymbol p, p0;
            string ss = "";
            bool tneg = false;
            string Result = mResult;
            if (mIsSimple) return Result;
            Parse(mSubject);
            tneg = false;
            Result = "";
            i = 0;
            while (i < mSblList.Count)
            {
                p = mSblList[i];
                if (p.flag < 0)
                {
                    i = i + 1;
                    continue;
                }
                if ((p.flag == 0) && (p.value == "+")) tneg = false;
                else if ((p.flag == 0) && (p.value == "-")) tneg = true;

                if ((p.flag == 0) && (p.value == "("))
                {
                    ss = "";
                    j = i - 1;
                    while (j >= 0)
                    {

                        p0 = mSblList[j];
                        if (p0.flag >= 0)
                        {
                            if ((p0.flag == 0) && (p0.value != "*") && (p0.value != "/")) break;
                            ss = p0.value + ss;
                            p0.flag = -2;
                        }
           
                        j = j - 1;
                    }

                    i = RemovemBrackets(i, ss, tneg);
                    continue;
                }
                i = i + 1;
            }

            j = 0;
            if (mSblList.Count > 2)
            {
                p0 = mSblList[0];
                p = mSblList[1];
                if ((p0.value == "0") && (p.value == "+")) j = 2;
            }


            for (i = j; i <= mSblList.Count - 1; i++)
            {
                p = mSblList[i];
                if (p.flag >= 0) Result = Result + p.value;

            }
            SimpleParse(Result);
            Result = ToSimple(isfloat);
            mIsSimple = true;
            return Result;
        }
    }
}
