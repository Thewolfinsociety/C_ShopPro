using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
namespace Base64ToImage
{
    struct BomOrderItem
    {
        int cid, id, pid, seq, classseq, mark, vp;
        string code, name, mat, mat2, mat3, color, workflow;
        int pl, pd, ph, space_x, space_y, space_z, space_id, gcl, gcd, gch, gcl2, gcd2, gch2;  //gcl, gcd, gch图形成品修正值
        string tmp_soz;
        int lx, ly, lz, x, y, z, l, p, h, gl, gp, gh, holeflag, linemax, holetype;
        double ox, oy, oz;
        int childnum;
        string desc, bomdes, bomwjdes, bomstddes, childbom, myclass, nodename, linecalc, bomstd, bg;
        int direct, lgflag, holeid, kcid;
        int num;
        double lfb, llk, wfb, wlk, llfb, rrfb, ddfb, uufb, fb;
        string holestr, kcstr;
        string memo, gno, gdes, gcb, extra, fbstr, subspace, process, ls, myunit, bomtype, bdxmlid, user_fbstr;
        double bl, bp, bh;                 //物料尺寸
        string[] var_name;
        int[] var_args;
        //$左收口宽度， $右收口宽度， $柱切角宽度， $柱切角深度， $梁切角深度， $梁切角高度， $左侧趟门位， $右侧趟门位
        int value_lsk, value_rsk, value_zk, value_zs, value_ls, value_lg, value_ltm, value_rtm;
        string a_hole_info, b_hole_info, holeinfo;
        bool isoutput, is_outline;
        string outputtype, holeconfig_flag, kcconig_flag, bg_data, mBGParam;
        string bg_filename, mpr_filename, bpp_filename, devcode, bg_filestring, bdstring;
        //增加bg_filename 对应的字符串
        int zero_y, direct_calctype, youge_holecalc; //zero_y封边靠档
        int is_output_bgdata, is_output_mpr, is_output_bpp;

        //基础图形描述
        int bg_l_minx, bg_l_maxx, bg_r_minx, bg_r_maxx, bg_l_miny, bg_l_maxy, bg_r_miny, bg_r_maxy;
        int bg_d_minx, bg_d_maxx, bg_u_minx, bg_u_maxx, bg_d_miny, bg_d_maxy, bg_u_miny, bg_u_maxy;
        int bg_b_minx, bg_b_maxx, bg_f_minx, bg_f_maxx, bg_b_miny, bg_b_maxy, bg_f_miny, bg_f_maxy;

        int hole_back_cap, hole_2_dist; //第一个孔靠背的距离，两孔间距

        bool trans_ab;                  //AB面反转
        int[] ahole_index; int[] bhole_index;
        int[] akc_index;
        int[] bkc_index;
        int[] is_calc_holeconfig;
        //void* parent;
        //void* userdata;

        double basewj_price;
        ////// ERP数据
        string extend, group, packno, userdefine, erpunit, erpmatcode, blockmemo, number_text, price_calctype, table_type, worktype, munit;
    };

    struct TProductItem    //产品结构体
    {
        public string name, gno, des, gcb, color, mat, Extra, pricecalctype;   //增加板材单价
        public int id, l, d, h, bh;
        public string productlevel, productnum, spacename, seriesname, ordertype;
    }

    struct BomParam
    {
        public int productid, cid, boardheight;
        public List<BomOrderItem> blist;
        public string xml;
        public string gno, gdes, gcb, extra, pname, subspace, sozflag, textureclass, pmat, pcolor, group;
        public int pid, pl, pd, ph;
        public int px, py, pz, space_x, space_y, space_z, space_id;
        public string outputtype;
        public string blockmemo, number_text, pricecalctype;
        public int num, mark;
        public XmlNode rootnode;
        public XmlDocument xdoc;
    }

    class BomApi : Myutil
    {
        public double length;
        public double breadth;
        public double height;
        public int value_lsk, value_rsk, value_zk, value_zs, value_ls, value_lg, value_ltm, value_rtm;
        string[] mVName;
        string[] mVValue;
        int[] mC;
       
        List<BomOrderItem> bomlist;
        List<TProductItem> mProductList;
        public void UnloadBom(bool freedata = true)
        {
            if (bomlist == null)  return;

        }
        public void InitSysVariantValue()
        {
            value_lsk = 0;
            value_rsk = 0;
            value_zk = 0;
            value_zs = 0;
            value_ls = 0;
            value_lg = 0;
            value_ltm = 0;
            value_rtm = 0;
        }
        void MyVariant(string str, ref  string s1, ref string s2)
        {
            string ws = str;
            int n = 0;
    
            n = Pos(":", ws);
            s1 = LeftStr(ws, n);
            s2 = RightStr(ws, Length(ws) - n);

        }
        public int ImportXomItemForBom(ref BomParam param0, ref int id, ref int slino)
        {
            int Result = 0;
            string ls = "", childxml = "";
            XmlNode node = null;
            BomParam param = new BomParam();
            TExpress exp = new TExpress();

            param.blockmemo = "";
            XmlNode root = param0.rootnode;
            string str = GetAttributeValue(root, "模块备注", "", "");
            
            if (str != "")
            {
                param0.blockmemo = str;
                param0.blockmemo = param0.blockmemo.Replace("[宽]", param0.pl.ToString());
                param0.blockmemo = param0.blockmemo.Replace("[深]", param0.pd.ToString());
                param0.blockmemo = param0.blockmemo.Replace("[高]", param0.ph.ToString());

            }
            str = GetAttributeValue(root, "类别", "", "");
            if ((str == "趟门,趟门") || (str == "掩门,掩门"))
            {
                node = root.SelectSingleNode("模板");
                if (node != null)
                {
                    childxml = "";
                    if (node.ChildNodes.Count > 0) childxml = node.ChildNodes[0].OuterXml;
                    return 0;
                }

            }

            node = root.SelectSingleNode("我的模块");
            if (node != null)
            {

                for (int i = 0; i <= node.ChildNodes.Count - 1; i++)
                {
                    XmlNode cnode = node.ChildNodes[i];
                    if ((cnode.Name != "板件") && (cnode.Name != "五金") && (cnode.Name != "型材五金") && (cnode.Name != "模块") && (cnode.Name != "门板")) continue;
                    childxml = "";
                    if (cnode.ChildNodes.Count > 0) childxml = cnode.ChildNodes[0].OuterXml;
                    if (childxml != "")
                    {
                        param = param0;
                        if (cnode.ChildNodes.Count > 0)
                        {
                            param.rootnode = cnode.ChildNodes[0];
                            param.xdoc = param0.xdoc;
                            int childnum = ImportXomItemForBom( ref param, ref id, ref slino);
                        } 
                    }

                }
                      
                          

            }
            
            return 0;
        }
        public string LoadXML2Bom(string xml)    //加载订单数据
        {
            int id = 1;
            int slino = 1;
            int cid = 0;
            string name, mat, color, des, gcb, extra, pricecalctype, spaceflag;
            int l, d, h, bh;
            UnloadBom(false);
            if (bomlist == null) bomlist = new List<BomOrderItem>();

            //for (int i = 0; i < bomlist.Length; i++)
            //{
            //    BomOrderItem poi = bomlist[i];
            //}
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xml);
            XmlNode root = xmlDoc.DocumentElement;
            if ((root.Name != "产品") && (root.ChildNodes.Count <= 0)) return "";
            InitSysVariantValue();
            name = "";
            l = 0;
            d = 0;
            h = 0;
            bh = 18;
            mat = "";
            color = "";
            des = "";
            gcb = "";
            spaceflag = "";
            pricecalctype = "";
            extra = "";
            XmlAttribute attri = root.Attributes["Extra"];
            if (attri != null)
            {
                extra = attri.Value;
                //int startrow = int.Parse(attri.Value);
            }
            if (root.ChildNodes[0].Name == "产品")
            {
                attri = root.Attributes["板材厚度"];
                if (attri != null)
                {
                    bh = int.Parse(attri.Value);
                    //int startrow = int.Parse(attri.Value);
                }
            }
            name = Myutil.GetAttributeValue(root, "名称", "", "");
            des = Myutil.GetAttributeValue(root , "描述", "", "");
            gcb = Myutil.GetAttributeValue(root, "CB", "", "");
            l = Myutil.GetAttributeValue(root, "宽", 0, 0);
            d = Myutil.GetAttributeValue(root, "深", 0, 0);
            h = Myutil.GetAttributeValue(root, "高", 0, 0);
            gcb = Myutil.GetAttributeValue(root, "材料", "", "");
            gcb = Myutil.GetAttributeValue(root, "颜色", "", "");
            gcb = Myutil.GetAttributeValue(root, "基础图形", "", "");
            spaceflag = Myutil.GetAttributeValue(root, "SpaceFlag", "", "");
            pricecalctype = Myutil.GetAttributeValue(root, "板材单价", "", "");
            mVName = new string[16];
            mVValue = new string[16];
            mC = new int[16];
        
            for (int j = 0; j<=15; j++)
            {
                mVName[j] = "";
                mVValue[j] = "";
                mC[j] = 0;
                //Myutil.GetAttributeValue(root, "参数" + j.ToString(), "", "");
                MyVariant(Myutil.GetAttributeValue(root, "参数" + j.ToString(), "", ""), ref mVName[j], ref mVValue[j]);
                mC[j] = MyStrToInt(mVValue[j]);
            }
            TProductItem p = new TProductItem();
            mProductList = new List<TProductItem>();
            p.id = cid;
            p.name = name;
            p.gno = name;
            p.mat = mat;
            p.color = color;
            p.des = des;
            p.gcb = gcb;
            p.Extra = extra;
            p.l = l;
            p.d = d;
            p.h = h;
            p.bh = bh;
           
            mProductList.Add(p);
            BomParam param = new BomParam();

            param.productid = mProductList.Count - 1;
            param.cid = cid - 1;
            param.boardheight = bh;
            param.blist = bomlist;
            param.gno = name;
            param.gdes = des;
            param.gcb = gcb;
            param.extra = extra;
            param.pname = "";
            param.subspace = "";
            param.sozflag = "";
            param.xml = root.ChildNodes[0].OuterXml;
            param.textureclass = "";
            param.pmat = mat;
            param.pcolor = color;
            param.pid = -1;
            param.pl = l;
            param.pd = d;
            param.ph = h;
            param.px = 0;
            param.py = 0;
            param.pz = 0;
            param.space_x = 0;
            param.space_y = 0;
            param.space_z = 0;
            if (spaceflag == "1") param.space_id = 0;
            else param.space_id = -1;
            param.outputtype = "";
            param.num = 1;
            //param.parent = null;
            param.blockmemo = "";
            param.number_text = "";
            param.rootnode = root.ChildNodes[0];
            param.xdoc = xmlDoc;
            ImportXomItemForBom(ref param, ref id, ref slino);
            return "";
        }
    }

}
