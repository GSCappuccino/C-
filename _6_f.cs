using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace MyApp
{
    public partial class _6_f : Form,FormFather
    {
        /*
                            * 查询数据块-
                            */
        public Dictionary<string, string> Item_list = new Dictionary<string, string>() { { "a0", "内部管理编号" } };

        /* 
         * -查询数据块
         */
        private string DB_table_name = "f81wts"; 
        public static object text_a0;//传A0的值
        private string FormId = "6";
        //数据库链接
        public static string ConnString = "Provider="+main.Office_Engen+";Data Source=.\\" + main.companyName + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password="+main.appDB_password+"";
        public static OleDbConnection MyConn;

        public static OleDbDataAdapter MyAd;
        public static OleDbCommandBuilder objCommandBuilder;
        public DataSet MyData = new DataSet();//不能设置为静态，，fill一次增加一次数据，，，记录条数翻倍
        public static object MyData_delete;
        public static int i;
        public _6_f()
        {
            InitializeComponent();
            i = 0;
            text_a0 = this.a0;
            MyData_delete = MyData;
            Control.CheckForIllegalCrossThreadCalls = false;
            _6_f.MyConn = new OleDbConnection(_6_f.ConnString);
            _6_f.MyConn.Open();
            _6_f.MyAd = new OleDbDataAdapter("select *  from f81wts ", _6_f.MyConn);
            _6_f.MyAd.Fill(MyData);
            objCommandBuilder = new OleDbCommandBuilder(_6_f.MyAd);
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\6");

            panel1.Enabled = false;

            button8.Enabled = false;
            button9.Enabled = false;
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                textBox17.Text = (MyData.Tables[0].Rows.Count).ToString();
                textBox18.Text = (i + 1).ToString();

                a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                a9.Text = MyData.Tables[0].Rows[i][9].ToString();

                a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();
                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                a18.Text = MyData.Tables[0].Rows[i][18].ToString();
           

            }
            else
            {
                textBox17.Text = "0";
                textBox18.Text = "0";

            }
        }

        private void _6_f_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0)
            {
                i = 0;
                a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                a9.Text = MyData.Tables[0].Rows[i][9].ToString();
                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a10_1.Text =Convert.ToDateTime( MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                a10_2.Text =Convert.ToDateTime( MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                a10_3.Text =Convert.ToDateTime( MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();

                a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();

                a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                a18.Text = MyData.Tables[0].Rows[i][18].ToString();
               
                textBox18.Text = (i + 1).ToString();
            }
            else
            {
                i = 0;
                a0.Text = null;
                a1.Text = null;
                a2.Text = null;
                a3.Text = null;
                a4.Text = null;
                a5.Text = null;
                a6.Text = null;
                a7.Text = null;
                a8.Text = null;
                a9.Text = null;
                a10_1.Text = null;
                a10_2.Text = null;
                a10_3.Text = null;
                a11.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;
                a15.Text = null;
                a16.Text = null;
                a17.Text = null;
                a18.Text = null;
      
                textBox18.Text = null;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0 && i > 0)
            {
                i = i - 1;
                a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                a9.Text = MyData.Tables[0].Rows[i][9].ToString();

                a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();

                a11.Text = MyData.Tables[0].Rows[i][11].ToString();

                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();

                a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                a18.Text = MyData.Tables[0].Rows[i][18].ToString();
           

                textBox18.Text = (i + 1).ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0 && MyData.Tables[0].Rows.Count - 1 > i)
            {
                i++;


                a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                a9.Text = MyData.Tables[0].Rows[i][9].ToString();

                a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();


                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                a18.Text = MyData.Tables[0].Rows[i][18].ToString();
           
                textBox18.Text = (i + 1).ToString();
            }
            else if (MyData.Tables[0].Rows.Count == 0)
            {
                i = -1;
                a0.Text = null;
                a1.Text = null;
                a2.Text = null;
                a3.Text = null;
                a4.Text = null;
                a5.Text = null;
                a6.Text = null;
                a7.Text = null;
                a8.Text = null;
                a9.Text = null;
                a11.Text = null;
                a10_1.Text = null;
                a10_2.Text = null;
                a10_3.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;

                a15.Text = null;
                a16.Text = null;
                a17.Text = null;
                a18.Text = null;
            
                textBox18.Text = null;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0)
            {
                i = MyData.Tables[0].Rows.Count - 1;
                a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                a9.Text = MyData.Tables[0].Rows[i][9].ToString();
                a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();



                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                a18.Text = MyData.Tables[0].Rows[i][18].ToString();
        
                textBox18.Text = (i + 1).ToString();
            }
            else
            {
                i = 0;
                a0.Text = null;
                a1.Text = null;
                a2.Text = null;
                a3.Text = null;
                a4.Text = null;
                a5.Text = null;
                a6.Text = null;
                a7.Text = null;
                a8.Text = null;
                a9.Text = null;
                a11.Text = null;
                a10_1.Text = null;
                a10_2.Text = null;
                a10_3.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;

                a15.Text = null;
                a16.Text = null;
                a17.Text = null;
                a18.Text = null;
         
                textBox18.Text = null;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel2.Enabled = false;



            /*  
              MyConn.Open();*/

            if (MyData.Tables[0].Rows.Count == 0)
                a0.Text = (60100001).ToString();
            else
            {
                OleDbDataAdapter MyAd00 = new OleDbDataAdapter("select a0  from f81wts order by a0 desc ", MyConn);
                DataSet MyData00 = new DataSet();
                MyAd00.Fill(MyData00);
                a0.Text = (int.Parse(MyData00.Tables[0].Rows[0][0].ToString()) + 1).ToString();

                // MessageBox.Show(a0.Text);
            }
            a1.Text = null;
            a2.Text = null;
            a3.Text = null;
            a4.Text = null;
            a5.Text = null;
            a6.Text = null;
            a8.Text = null;
            a7.Text = null;
            a8.Text = null;
            a9.Text = null;
            a11.Text = null;
            a10_1.Text =main.currentTime.Year.ToString();
            a10_2.Text =main.currentTime.Month.ToString();
            a10_3.Text = main.currentTime.Day.ToString() ;
            a12.Text = a0.Text;

            a13.Text = null;
            a14.Text = null;
            a15.Text = main.UserId;
            a16.Text = null;


            a17.Text = null;
            a18.Text =null;
       
            panel1.Enabled = true;


            button6.Enabled = false;
            button7.Enabled = false;

            button8.Enabled = true;
            button9.Enabled = true;

            button10.Enabled = false;
            button11.Enabled = false;
            button12.Enabled = false;
     
     
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                panel1.Enabled = true;


                button6.Enabled = false;
                button7.Enabled = false;

                button8.Enabled = true;
                button9.Enabled = true;

                button10.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
      
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            for (int v = 0; v < MyData.Tables[0].Rows.Count; v++)
            {

                //     MessageBox.Show("run");
                if (this.a0.Text == MyData.Tables[0].Rows[v][0].ToString())
                {

                    string SQL = "delete from f81wts where a0= '" + this.a0.Text + "'";
                    OleDbCommand MyCom = new OleDbCommand(SQL, MyConn);
                    MyCom.ExecuteNonQuery();
                    MyData.Tables[0].Rows.RemoveAt(v);
                    break;
                }
            }

            panel2.Enabled = true;


            /*
            *word  存储开始
            *
            */
            //代码域

            Thread MySaveS = new Thread(new ThreadStart(save));
            MySaveS.Start();

            /*
             * word存储结束
             */
            /*
             * 数据库存储开始
             */
            DataRow NewRow = MyData.Tables[0].NewRow();
            NewRow["a0"] = a0.Text;
            NewRow["a1"] = a1.Text;
            NewRow["a2"] = a2.Text;
            NewRow["a3"] = a3.Text;
            NewRow["a4"] = a4.Text;
            NewRow["a5"] = a5.Text;
            NewRow["a6"] = a6.Text;
            NewRow["a7"] = a7.Text;
            NewRow["a8"] = a8.Text;
            NewRow["a9"] = a9.Text;
            NewRow["a10"] = a10_1.Text + '/' + a10_2.Text + '/' + a10_3.Text;
            NewRow["a11"] = a11.Text;
            NewRow["a12"] = a12.Text;
            NewRow["a13"] = a13.Text;
            NewRow["a14"] = a14.Text;
            NewRow["a15"] = a15.Text;
            NewRow["a16"] = a16.Text;
            NewRow["a17"] = a17.Text;
            NewRow["a18"] = a18.Text;
         

            NewRow["time_e"] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();


            MyData.Tables[0].Rows.Add(NewRow);

            _6_f.MyAd.Update(MyData);



            //MessageBox.Show(a13_1.CheckState);



            /*
             * 数据库的存储结束
             *
             * *
             */






            MessageBox.Show("保存成功！！");

            i = MyData.Tables[0].Rows.Count - 1;
            textBox18.Text = (i + 1).ToString();


            panel1.Enabled = false;

            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = true;
         
        
            this.textBox17.Text = MyData.Tables[0].Rows.Count.ToString();
        }

        private void save()
        {
            this.Enabled = true;
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\6\\" + a0.Text + "");

            string temp = System.IO.Directory.GetCurrentDirectory();// d当前运行路径

            //    MessageBox.Show(temp);
            string OrignFile;//word模板路径
            OrignFile = "\\baseDB\\商标书式\\6 商标代理委托书\\商标代理委托书（示范文本）.dot";

            //开始写入数据

            string parFilePath = temp + OrignFile;//文件路径
            object FilePath = parFilePath;
            Microsoft.Office.Interop.Word._Application AppliApp = new Microsoft.Office.Interop.Word.Application();
            AppliApp.Visible = false;
            Microsoft.Office.Interop.Word._Document doc = AppliApp.Documents.Add(ref FilePath);
            object missing = System.Reflection.Missing.Value;
            object isReadOnly = false;




            doc.Activate();

            //数据写入代码段


            object aa = temp + "\\" + main.companyName + "\\wjgl\\6\\" + a0.Text + "\\" + a0.Text + ".docx";

            object[] MyBM = new object[57];//创建一个书签数组

            for (int i = 0; i < 4; i++)//给书签数组赋值
                MyBM[i] = "a" + (i + 1).ToString();
            for (int i = 4; i < 42;i++ )
                MyBM[i] = "a5_" + (i-3).ToString();
            for (int i = 42; i < 55; i++)
                MyBM[i] = "a" + (i - 36).ToString();
            MyBM[46] = "a10_1";
            MyBM[55] = "a10_2";
            MyBM[56] = "a10_3";
                //给对应的书签位置写入数据
                doc.Bookmarks.get_Item(ref MyBM[0]).Range.Text = a1.Text;
            doc.Bookmarks.get_Item(ref MyBM[1]).Range.Text = a2.Text;
            doc.Bookmarks.get_Item(ref MyBM[2]).Range.Text = a3.Text;
            doc.Bookmarks.get_Item(ref MyBM[3]).Range.Text = a4.Text;

            if (a5.Text =="商标注册申请")
                doc.Bookmarks.get_Item(ref MyBM[4]).Range.Text = "✔";
           else if (a5.Text == "商标异议申请")
                doc.Bookmarks.get_Item(ref MyBM[5]).Range.Text = "✔";
           else if (a5.Text == "商标异议答辩")
                doc.Bookmarks.get_Item(ref MyBM[6]).Range.Text = "✔";
           else if (a5.Text == "更正商标申请/注意事项申请")
                doc.Bookmarks.get_Item(ref MyBM[7]).Range.Text = "✔";
           else if (a5.Text == "变更商标申请人/注册人名义/地址 变更集体商标/证明商标管理规则/集体成员名单申请")
                doc.Bookmarks.get_Item(ref MyBM[8]).Range.Text = "✔";
           else if (a5.Text == "变更商标代理人/文件接收人申请")
                doc.Bookmarks.get_Item(ref MyBM[9]).Range.Text = "✔";
           else if (a5.Text == "删减商品/服务项目申请")
                doc.Bookmarks.get_Item(ref MyBM[10]).Range.Text = "✔";
           else if (a5.Text == "商标续展注册申请")
                doc.Bookmarks.get_Item(ref MyBM[11]).Range.Text = "✔";
           else if (a5.Text == "商标注册申请")
                doc.Bookmarks.get_Item(ref MyBM[12]).Range.Text = "✔";
          else  if (a5.Text == "转让/移转申请/注册商标申请书")
             
                doc.Bookmarks.get_Item(ref MyBM[13]).Range.Text = "✔";
             else if (a5.Text == "商标使用许可备案")
                doc.Bookmarks.get_Item(ref MyBM[14]).Range.Text = "✔";
             else if (a5.Text == "变更许可人/被许可人名称备案")
                doc.Bookmarks.get_Item(ref MyBM[15]).Range.Text = "✔";
            else if (a5.Text == "商标专用权质权登记申请")
                doc.Bookmarks.get_Item(ref MyBM[16]).Range.Text = "✔";
            else if (a5.Text == "商标专用权质权登记事项变更申请")
                doc.Bookmarks.get_Item(ref MyBM[17]).Range.Text = "✔";
            else if (a5.Text == "商标专用权质权登记期限延期申请")
                doc.Bookmarks.get_Item(ref MyBM[18]).Range.Text = "✔";
            else if (a5.Text == "商标专用权质权登记证补发申请")
                doc.Bookmarks.get_Item(ref MyBM[19]).Range.Text = "✔";
            else if (a5.Text == "商标专用权质权登记注销申请")
                doc.Bookmarks.get_Item(ref MyBM[20]).Range.Text = "✔";
            else if (a5.Text == "商标注销申请")
                doc.Bookmarks.get_Item(ref MyBM[21]).Range.Text = "✔";
            else if (a5.Text == "撤销连续三年不使用注册商标申请")
                doc.Bookmarks.get_Item(ref MyBM[22]).Range.Text = "✔";
            else if (a5.Text == "撤销成为商品/服务通用名称注册商标申请")
                doc.Bookmarks.get_Item(ref MyBM[23]).Range.Text = "✔";
            else if (a5.Text == "撤销连续三年不使用注册商标提供证据")
                doc.Bookmarks.get_Item(ref MyBM[24]).Range.Text = "✔";
             else if (a5.Text == "撤销成为商品/服务通用名称注册商标答辩")
                doc.Bookmarks.get_Item(ref MyBM[25]).Range.Text = "✔";
            else if (a5.Text == "补发变更/转让/续展证明申请")
                doc.Bookmarks.get_Item(ref MyBM[26]).Range.Text = "✔";
            else if (a5.Text == "补发商标注册证申请")
                doc.Bookmarks.get_Item(ref MyBM[27]).Range.Text = "✔";
            else if (a5.Text == "出具商标注册证明申请")
                doc.Bookmarks.get_Item(ref MyBM[28]).Range.Text = "✔";
            else if (a5.Text == "出具优先权证明文件申请")
                doc.Bookmarks.get_Item(ref MyBM[29]).Range.Text = "✔";
            else if (a5.Text == "撤回商标注册申请")
                doc.Bookmarks.get_Item(ref MyBM[30]).Range.Text = "✔";
           else if (a5.Text == "撤回商标异议申请")
                doc.Bookmarks.get_Item(ref MyBM[31]).Range.Text = "✔";
           else if (a5.Text == "撤回变更商标申请人/注册人名义/地址 变更集体商标/证明商标管理规则/集体成员名单申请")
                doc.Bookmarks.get_Item(ref MyBM[32]).Range.Text = "✔";
            else if (a5.Text == "撤回变更商标代理人/文件接收人申请")
                doc.Bookmarks.get_Item(ref MyBM[33]).Range.Text = "✔";
           else if (a5.Text == "撤回删减商品/服务项目申请")
                doc.Bookmarks.get_Item(ref MyBM[34]).Range.Text = "✔";
           else if (a5.Text == "撤回商标续展注册申请")
                doc.Bookmarks.get_Item(ref MyBM[35]).Range.Text = "✔";
           else if (a5.Text == "撤回转让/移转申请/注册商标申请")
                doc.Bookmarks.get_Item(ref MyBM[36]).Range.Text = "✔";
           else if (a5.Text == "撤回商标使用许可备案")
                doc.Bookmarks.get_Item(ref MyBM[37]).Range.Text = "✔";
           else if (a5.Text == "撤回商标注销申请")
                doc.Bookmarks.get_Item(ref MyBM[38]).Range.Text = "✔";
           else if (a5.Text == "撤回撤销连续三年不使用注册商标申请")
                doc.Bookmarks.get_Item(ref MyBM[39]).Range.Text = "✔";
           else if (a5.Text == "撤回撤销成为商品/服务通用名称注册商标申请")
                doc.Bookmarks.get_Item(ref MyBM[40]).Range.Text = "✔";
            else
                doc.Bookmarks.get_Item(ref MyBM[41]).Range.Text = a5.Text;
           
            doc.Bookmarks.get_Item(ref MyBM[42]).Range.Text = a6.Text;
            doc.Bookmarks.get_Item(ref MyBM[43]).Range.Text = a7.Text;
            doc.Bookmarks.get_Item(ref MyBM[44]).Range.Text = a8.Text;
            doc.Bookmarks.get_Item(ref MyBM[45]).Range.Text = a9.Text;
            doc.Bookmarks.get_Item(ref MyBM[46]).Range.Text = a10_1.Text;
            doc.Bookmarks.get_Item(ref MyBM[55]).Range.Text = a10_2.Text;
            doc.Bookmarks.get_Item(ref MyBM[56]).Range.Text = a10_3.Text;
           /* doc.Bookmarks.get_Item(ref MyBM[47]).Range.Text = a11.Text;
            doc.Bookmarks.get_Item(ref MyBM[48]).Range.Text = a12.Text;
            doc.Bookmarks.get_Item(ref MyBM[49]).Range.Text = a13.Text;
            doc.Bookmarks.get_Item(ref MyBM[50]).Range.Text = a14.Text;
            doc.Bookmarks.get_Item(ref MyBM[51]).Range.Text = a15.Text;
            doc.Bookmarks.get_Item(ref MyBM[52]).Range.Text = a16.Text;
            doc.Bookmarks.get_Item(ref MyBM[53]).Range.Text = a17.Text;
            doc.Bookmarks.get_Item(ref MyBM[54]).Range.Text = a18.Text;*/
       
            doc.SaveAs(ref aa);
            doc.Close();
            this.Enabled = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            panel2.Enabled = true;
            panel1.Enabled = false;

            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = true;
          

            if (MyData.Tables[0].Rows.Count >= 1)
            {
                a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                a9.Text = MyData.Tables[0].Rows[i][9].ToString();
                a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();


                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                a18.Text = MyData.Tables[0].Rows[i][18].ToString();
          
            }
            else
            {
                a0.Text = null;
                a1.Text = null;
                a2.Text = null;
                a3.Text = null;
                a4.Text = null;
                a5.Text = null;
                a6.Text = null;
                a7.Text = null;
                a8.Text = null;
                a9.Text = null;
                a10_1.Text = null;
                a10_2.Text = null;
                a10_3.Text = null;
                a11.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;

                a15.Text = null;
                a16.Text = null;
                a17.Text = null;
                a18.Text = null;
      
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                panel2.Enabled = false;





                OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select top 1 * from f81wts order by a0 desc", MyConn);
                DataSet MyData0 = new DataSet();
                MyAd0.Fill(MyData0);

                a0.Text = (int.Parse((MyData0.Tables[0].Rows[0]["a0"].ToString())) + 1).ToString();

                panel1.Enabled = true;


                button6.Enabled = false;
                button7.Enabled = false;

                button8.Enabled = true;
                button9.Enabled = true;

                button10.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
             
       
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count != 0)
            {
                (new YNcleanALL_f()).ShowDialog();
                if (YNcleanALL_f.YNdel == true)
                {
                    CleanApp();
                    textBox17.Text = (int.Parse(textBox17.Text) - 1).ToString();

                    if (MyData.Tables[0].Rows.Count == 0)
                    {
                        textBox18.Text = "0";
                        a1.Text = null;
                        a2.Text = null;
                        a3.Text = null;
                        a4.Text = null;
                        a5.Text = null;
                        a6.Text = null;
                        a8.Text = null;
                        a7.Text = null;
                        a8.Text = null;
                        a9.Text = null;
                        a10_1.Text = null;
                        a10_2.Text = null;
                        a10_3.Text = null;
                        a11.Text = null;
                        a13.Text = null;
                        a14.Text = null;
                        a15.Text = null;
                        a12.Text = null;
                        a16.Text = null;
                        a17.Text = null;
                        a18.Text = null;
                      
                    }
                    else if (i == MyData.Tables[0].Rows.Count)
                    {
                        i = i - 1;
                        a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                        a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                        a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                        a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                        a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                        a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                        a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                        a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                        a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                        a9.Text = MyData.Tables[0].Rows[i][9].ToString();
                        a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                        a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                        a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();



                        a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                        a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                        a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                        a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                        a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                        a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                        a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                        a18.Text = MyData.Tables[0].Rows[i][18].ToString();
                    
                        textBox18.Text = (int.Parse(textBox18.Text) - 1).ToString();
                    }
                    else
                    {

                        a0.Text = MyData.Tables[0].Rows[i][0].ToString();
                        a1.Text = MyData.Tables[0].Rows[i][1].ToString();
                        a2.Text = MyData.Tables[0].Rows[i][2].ToString();
                        a3.Text = MyData.Tables[0].Rows[i][3].ToString();
                        a4.Text = MyData.Tables[0].Rows[i][4].ToString();
                        a5.Text = MyData.Tables[0].Rows[i][5].ToString();
                        a6.Text = MyData.Tables[0].Rows[i][6].ToString();
                        a7.Text = MyData.Tables[0].Rows[i][7].ToString();
                        a8.Text = MyData.Tables[0].Rows[i][8].ToString();
                        a9.Text = MyData.Tables[0].Rows[i][9].ToString();
                        a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
                        a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
                        a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();



                        a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                        a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                        a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                        a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                        a15.Text = MyData.Tables[0].Rows[i][15].ToString();
                        a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                        a17.Text = MyData.Tables[0].Rows[i][17].ToString();
                        a18.Text = MyData.Tables[0].Rows[i][18].ToString();
                      
                    }
                  
                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(".\\" + main.companyName + "\\wjgl\\6");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button14_Click(object sender, EventArgs e)
        {

        }

        public void CleanApp()
        {
            //内存数据删除
            ((DataSet)_6_f.MyData_delete).Tables[0].Rows.RemoveAt(_6_f.i);


            // int a =;// 数据库删除
            string SQL = "delete * from f81wts where a0='" + ((TextBox)_6_f.text_a0).Text + "'";
            OleDbCommand NewCom = new OleDbCommand(SQL, _6_f.MyConn);
            NewCom.ExecuteNonQuery();//执行SQL语句
            //  MessageBox.Show(_101_f.i.ToString());


            //文件删除&& 


            if (File.Exists(@".\\" + main.companyName + "\\wjgl\\6\\" + ((TextBox)_6_f.text_a0).Text + "\\" + ((TextBox)_6_f.text_a0).Text + ".docx"))
            {
                // MessageBox.Show("存在文件");
                File.Delete(@".\\" + main.companyName + "\\wjgl\\6\\" + ((TextBox)_6_f.text_a0).Text + "\\" + ((TextBox)_6_f.text_a0).Text + ".docx");

            }

            if (Directory.Exists(@".\\" + main.companyName + "\\wjgl\\6\\" + ((TextBox)_6_f.text_a0).Text + ""))
            {
                // MessageBox.Show("存在文件夹");
                Directory.Delete(@".\\" + main.companyName + "\\wjgl\\6\\" + ((TextBox)_6_f.text_a0).Text + "");
            }
            MessageBox.Show("删除成功");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            main.Print(MyData, a0.Text, FormId, int.Parse(null));
        }

        unsafe private void button1_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0)
            {
                fixed (int* pp = &i)
                {
                    Each_f_statistics a = new Each_f_statistics(DB_table_name, MyConn, pp, MyData.Tables[0], Item_list);
                    a.ShowDialog();
                }

                Record_show();
                textBox18.Text = (i + 1).ToString();

            }
        }
        private void Record_show()
        {
            a0.Text = MyData.Tables[0].Rows[i][0].ToString();
            a1.Text = MyData.Tables[0].Rows[i][1].ToString();
            a2.Text = MyData.Tables[0].Rows[i][2].ToString();
            a3.Text = MyData.Tables[0].Rows[i][3].ToString();
            a4.Text = MyData.Tables[0].Rows[i][4].ToString();
            a5.Text = MyData.Tables[0].Rows[i][5].ToString();
            a6.Text = MyData.Tables[0].Rows[i][6].ToString();
            a7.Text = MyData.Tables[0].Rows[i][7].ToString();
            a8.Text = MyData.Tables[0].Rows[i][8].ToString();
            a9.Text = MyData.Tables[0].Rows[i][9].ToString();

            a10_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Year.ToString();
            a10_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Month.ToString();
            a10_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][10].ToString()).Day.ToString();
            a11.Text = MyData.Tables[0].Rows[i][11].ToString();
            a12.Text = MyData.Tables[0].Rows[i][12].ToString();
            a13.Text = MyData.Tables[0].Rows[i][13].ToString();
            a14.Text = MyData.Tables[0].Rows[i][14].ToString();
            a15.Text = MyData.Tables[0].Rows[i][15].ToString();
            a16.Text = MyData.Tables[0].Rows[i][16].ToString();
            a17.Text = MyData.Tables[0].Rows[i][17].ToString();
            a18.Text = MyData.Tables[0].Rows[i][18].ToString();
        }
    }
}
