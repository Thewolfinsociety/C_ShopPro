using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows.Forms;
using ZXing;
using System.Xml;
using DataMatrix.net;
using System.Linq;

namespace Base64ToImage
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length >= 3)
            {
                string datafile = args[0];
                string savefile = args[1];
                string format = args[2];
                if (format == "datamatrix")
                {
                    DoConvert_DataMatrix(datafile);
                }
                else
                {
                    DoConvert(datafile, savefile, format);
                }
            }
            else
            {
                /*char[] OpSymbolArray = { '1', '2', ' '};
               
                char[] tmp = new char[64];
                tmp[0] = '1';
                tmp[1] = '\0';
                string str = new string(OpSymbolArray);
                //MessageBox.Show(str);
                string a = str.ToString() + "123";
                MessageBox.Show(a);
                //MessageBox.Show("456" + "123");
                char buf = '+';
                string str1 = str;
                int b = int.Parse(str1);
                MessageBox.Show(int.Parse(str1).ToString());
                if (OpSymbolArray.Contains(buf))
                {
                    //MessageBox.Show("hello");

                }
                if (buf.ToString() == "+")
                {
                   // MessageBox.Show("hello1");
                }
                string xmlfile = args[0];
                string xml = System.IO.File.ReadAllText(xmlfile, System.Text.Encoding.Default);
                char[] buf1 = xml.ToCharArray();

                char c3 = (char)0;
                for (int i = 0; i < xml.Length; i++)
                {
                    if (buf1[i] == 0)
                    {
                        MessageBox.Show("Hello2");
                    }
                }

                
                string xmlfile = args[0];
                string xml = System.IO.File.ReadAllText(xmlfile, System.Text.Encoding.Default);
                BomApi bom = new BomApi();
                bom.LoadXML2Bom(xml);*/
                TExpress exp = new TExpress();
                bool str = exp.IsNumeric("1223addd");

            }
        }

        static void DoConvert(string datafile, string savefile, string format)
        {
            if (!File.Exists(datafile)) return;

            string base64 = System.IO.File.ReadAllText(datafile, System.Text.Encoding.Default);       
            string[] sArray = Regex.Split(base64, ",", RegexOptions.IgnoreCase);    //分割字符串，分割","前后字符串成为数组对象初始序号值为0
            base64 = sArray[1];
            Console.Write(base64);
            byte[] arr = Convert.FromBase64String(base64);
            MemoryStream ms = new MemoryStream(arr);
            if (format == "pdf")
            {
                FileStream fs = new FileStream(savefile, FileMode.Create);
                //写入流文件
                ms.WriteTo(fs);
                return;
            }

            Bitmap bmp = new Bitmap(ms);
            if (format=="png")
            {
                bmp.Save(savefile, System.Drawing.Imaging.ImageFormat.Png);
            }
            if (format=="jpeg")
            {
                bmp.Save(savefile, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            if (format == "bmp")
            {
                bmp.Save(savefile, System.Drawing.Imaging.ImageFormat.Bmp);
            }
            if (format == "wmf")
            {
                bmp.Save(savefile, System.Drawing.Imaging.ImageFormat.Wmf);
            }
        }

        static void DoConvert_DataMatrix(string datafile){
            if (!File.Exists(datafile)) return;
            XmlDocument doc = new XmlDocument();
            doc.Load(datafile);
            XmlNode root = doc.DocumentElement;
            for (int i = 0; i < root.ChildNodes.Count; i++)
            {
                XmlNode item = root.ChildNodes[i];
                string content = item.Attributes[@"内容"].Value;
                int height = int.Parse(item.Attributes[@"高"].Value);
                int width = int.Parse(item.Attributes[@"宽"].Value);
                string filename = item.Attributes[@"内容"].Value;
                string savefile = item.Attributes[@"文件名"].Value;

                DmtxImageEncoder Die = new DmtxImageEncoder();
                DataMatrix.net.DmtxImageEncoderOptions option = new DmtxImageEncoderOptions();
                option.SizeIdx = DmtxSymbolSize.DmtxSymbolSquareAuto;//形状 
                option.MarginSize = 0;//边距  
                option.ModuleSize = 4;//点阵大小  

                Bitmap b = Die.EncodeImage(content, option);
                b.Save(savefile, System.Drawing.Imaging.ImageFormat.Bmp);
            }
        }
    }
}
