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
    public partial class _7_f : Form,FormFather
    {
        /*
                               * 查询数据块-
                               */
        public Dictionary<string, string> Item_list = new Dictionary<string, string>() { { "a0", "内部管理编号" } };

        /* 
         * -查询数据块
         */
        private string DB_table_name = "g8fswt"; 
        public static object text_a0;//传A0的值
        private string FormId = "7";
        //数据库链接
        public static string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + main.companyName + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
        public static OleDbConnection MyConn;

        public static OleDbDataAdapter MyAd;
        public static OleDbCommandBuilder objCommandBuilder;
        public DataSet MyData = new DataSet();//不能设置为静态，，fill一次增加一次数据，，，记录条数翻倍
        public static object MyData_delete;
        public static int i;
        public _7_f()
        {
            InitializeComponent();
           
            i = 0;
            text_a0 = this.a0;
            MyData_delete = MyData;
            Control.CheckForIllegalCrossThreadCalls = false;
            _7_f.MyConn = new OleDbConnection(_7_f.ConnString);
            _7_f.MyConn.Open();
            _7_f.MyAd = new OleDbDataAdapter("select *  from g8fswt ", _7_f.MyConn);
            _7_f.MyAd.Fill(MyData);
            objCommandBuilder = new OleDbCommandBuilder(_7_f.MyAd);
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\7");

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
                a10.Text = MyData.Tables[0].Rows[i][10].ToString();
               
                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
              

            }
            else
            {
                textBox17.Text = "0";
                textBox18.Text = "0";

            }
        }

        private void _7_f_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
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
                a10.Text = MyData.Tables[0].Rows[i][10].ToString();
                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();   
                a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();                
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
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
                a10.Text = MyData.Tables[0].Rows[i][10].ToString();
                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
               
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
                a10.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;
                a15_1.Text = null;
                a15_2.Text = null;
                a15_3.Text = null;
                a16.Text = null;
       
                textBox18.Text = null;
            }
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
                a10.Text = MyData.Tables[0].Rows[i][10].ToString();
               
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();
              
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
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
                a10.Text = null;
                a15_1.Text = null;
                a15_2.Text = null;
                a15_3.Text = null;
                a11.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;
         
                a16.Text = null;
           
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
                a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();



                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a10.Text = MyData.Tables[0].Rows[i][10].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
             
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
                a10.Text = null;
               
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;

                a15_1.Text = null;
                a15_2.Text = null;
                a15_3.Text = null;
                a16.Text = null;

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
                OleDbDataAdapter MyAd00 = new OleDbDataAdapter("select a0  from g8fswt order by a0 desc ", MyConn);
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
            a10.Text = null;
            a11.Text = null;
        
            a12.Text = null;

            a13.Text = null;
            a14.Text = null;
             
            a15_1.Text = main.currentTime.Year.ToString();
            a15_2.Text = main.currentTime.Month.ToString();
            a15_3.Text = main.currentTime.Day.ToString();
            a16.Text =a0.Text;


  
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

                    string SQL = "delete from g8fswt where a0= '" + this.a0.Text + "'";
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
           
            NewRow["a11"] = a11.Text;
            NewRow["a12"] = a12.Text;
            NewRow["a13"] = a13.Text;
            NewRow["a14"] = a14.Text;
            NewRow["a10"] = a10.Text;
            NewRow["a16"] = a16.Text;
            NewRow["a15"] = a15_1.Text + '/' + a15_2.Text + '/' + a15_3.Text;

            NewRow["time_e"] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();


            MyData.Tables[0].Rows.Add(NewRow);

            _7_f.MyAd.Update(MyData);



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
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\7\\" + a0.Text + "");

            string temp = System.IO.Directory.GetCurrentDirectory();// d当前运行路径

            //    MessageBox.Show(temp);
            string OrignFile;//word模板路径
            OrignFile = "\\baseDB\\商标书式\\7 评审复审\\8商标评审代理委托书\\商标评审代理委托书（样式）.dot";

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


            object aa = temp + "\\" + main.companyName + "\\wjgl\\7\\" + a0.Text + "\\" + a0.Text + ".docx";

            object[] MyBM = new object[21];//创建一个书签数组

            for (int i = 0; i < 12; i++)//给书签数组赋值
                MyBM[i] = "a" + (i + 1).ToString();
            MyBM[12] = "a13_1";
            MyBM[13] = "a13_2";
            MyBM[14] = "a13_3";
            MyBM[15] = "a13_4";
            MyBM[16] = "a13_5";
            MyBM[17] = "a14";
            MyBM[18] = "a15_1";
            MyBM[19] = "a15_2";
            MyBM[20] = "a15_3";
          //  MyBM[21] = "a16";
     /*      驳回商标注册申请复审案
商标不予注册复审案
撤销注册商标复审案
注册商标无效宣告案
注册商标无效宣告复审案*/
           
            //给对应的书签位置写入数据
            doc.Bookmarks.get_Item(ref MyBM[0]).Range.Text = a1.Text;
            doc.Bookmarks.get_Item(ref MyBM[1]).Range.Text = a2.Text;
            doc.Bookmarks.get_Item(ref MyBM[2]).Range.Text = a3.Text;
            doc.Bookmarks.get_Item(ref MyBM[3]).Range.Text = a4.Text;
            doc.Bookmarks.get_Item(ref MyBM[4]).Range.Text = a5.Text;
            doc.Bookmarks.get_Item(ref MyBM[5]).Range.Text = a6.Text;
            doc.Bookmarks.get_Item(ref MyBM[6]).Range.Text = a7.Text;
            doc.Bookmarks.get_Item(ref MyBM[7]).Range.Text = a8.Text;
            doc.Bookmarks.get_Item(ref MyBM[8]).Range.Text = a9.Text;
            doc.Bookmarks.get_Item(ref MyBM[9]).Range.Text = a10.Text;
            doc.Bookmarks.get_Item(ref MyBM[10]).Range.Text = a11.Text;
            doc.Bookmarks.get_Item(ref MyBM[11]).Range.Text = a12.Text;
            if (a13.Text == "驳回商标注册申请复审案")
                doc.Bookmarks.get_Item(ref MyBM[12]).Range.Text = "✔";
            if (a13.Text == "商标不予注册复审案")
                    doc.Bookmarks.get_Item(ref MyBM[13]).Range.Text = "✔";
            if (a13.Text == "撤销注册商标复审案")
                    doc.Bookmarks.get_Item(ref MyBM[14]).Range.Text = "✔";
            if (a13.Text == "注册商标无效宣告案")
                    doc.Bookmarks.get_Item(ref MyBM[15]).Range.Text = "✔";
            if (a13.Text == "注册商标无效宣告复审案")
                    doc.Bookmarks.get_Item(ref MyBM[16]).Range.Text = "✔";
            doc.Bookmarks.get_Item(ref MyBM[17]).Range.Text = a14.Text;
            doc.Bookmarks.get_Item(ref MyBM[18]).Range.Text = a15_1.Text;
            doc.Bookmarks.get_Item(ref MyBM[19]).Range.Text = a15_2.Text;
            doc.Bookmarks.get_Item(ref MyBM[20]).Range.Text = a15_3.Text;
          //  doc.Bookmarks.get_Item(ref MyBM[21]).Range.Text = a16.Text;
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
                a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();


                a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                a10.Text = MyData.Tables[0].Rows[i][10].ToString();
                a16.Text = MyData.Tables[0].Rows[i][16].ToString();
            
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
                a15_1.Text = null;
                a15_2.Text = null;
                a15_3.Text = null;
                a11.Text = null;
                a12.Text = null;
                a13.Text = null;
                a14.Text = null;

                a10.Text = null;
                a16.Text = null;
               

            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                panel2.Enabled = false;





                OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select top 1 * from g8fswt order by a0 desc", MyConn);
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
                        a15_1.Text = null;
                        a15_2.Text = null;
                        a15_3.Text = null;
                        a11.Text = null;
                        a13.Text = null;
                        a14.Text = null;
                        a10.Text = null;
                        a12.Text = null;
                        a16.Text = null;
                    
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
                        a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                        a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                        a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();



                        a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                        a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                        a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                        a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                        a10.Text = MyData.Tables[0].Rows[i][10].ToString();
                        a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                      
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
                        a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
                        a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
                        a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();



                        a11.Text = MyData.Tables[0].Rows[i][11].ToString();
                        a12.Text = MyData.Tables[0].Rows[i][12].ToString();
                        a13.Text = MyData.Tables[0].Rows[i][13].ToString();
                        a14.Text = MyData.Tables[0].Rows[i][14].ToString();
                        a10.Text = MyData.Tables[0].Rows[i][10].ToString();
                        a16.Text = MyData.Tables[0].Rows[i][16].ToString();
                    
                    }
                    
                }
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(".\\" + main.companyName + "\\wjgl\\7");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void CleanApp()
        {
            //内存数据删除
            ((DataSet)_7_f.MyData_delete).Tables[0].Rows.RemoveAt(_7_f.i);


            // int a =;// 数据库删除
            string SQL = "delete * from g8fswt where a0='" + ((TextBox)_7_f.text_a0).Text + "'";
            OleDbCommand NewCom = new OleDbCommand(SQL, _7_f.MyConn);
            NewCom.ExecuteNonQuery();//执行SQL语句
            //  MessageBox.Show(_101_f.i.ToString());


            //文件删除&& 


            if (File.Exists(@".\\" + main.companyName + "\\wjgl\\7\\" + ((TextBox)_7_f.text_a0).Text + "\\" + ((TextBox)_7_f.text_a0).Text + ".docx"))
            {
                // MessageBox.Show("存在文件");
                File.Delete(@".\\" + main.companyName + "\\wjgl\\7\\" + ((TextBox)_7_f.text_a0).Text + "\\" + ((TextBox)_7_f.text_a0).Text + ".docx");

            }

            if (Directory.Exists(@".\\" + main.companyName + "\\wjgl\\7\\" + ((TextBox)_7_f.text_a0).Text + ""))
            {
                // MessageBox.Show("存在文件夹");
                Directory.Delete(@".\\" + main.companyName + "\\wjgl\\7\\" + ((TextBox)_7_f.text_a0).Text + "");
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
            a10.Text = MyData.Tables[0].Rows[i][10].ToString();

            a11.Text = MyData.Tables[0].Rows[i][11].ToString();
            a12.Text = MyData.Tables[0].Rows[i][12].ToString();
            a13.Text = MyData.Tables[0].Rows[i][13].ToString();
            a14.Text = MyData.Tables[0].Rows[i][14].ToString();
            a15_1.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Year.ToString();
            a15_2.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Month.ToString();
            a15_3.Text = Convert.ToDateTime(MyData.Tables[0].Rows[i][15].ToString()).Day.ToString();
            a16.Text = MyData.Tables[0].Rows[i][16].ToString();
        }
    }
}
