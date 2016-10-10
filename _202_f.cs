﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace MyApp
{
    public partial class _202_f : Form,FormFather
    {
        /*
  * 查询数据块-
  */
        public Dictionary<string, string> Item_list = new Dictionary<string, string>() { { "a0", "内部管理编号" } };

        /* 
         * * -查询数据块
        */
        private string DB_table_name = "bsbyy2";
        private string FormId="202";
        public  object text_a0;//传A0的值

        //数据库链接
        public  string ConnString = "Provider="+main.Office_Engen+";Data Source=.\\" + main.companyName + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password="+main.appDB_password+"";
        public OleDbConnection MyConn;

        public  OleDbDataAdapter MyAd;
        public  OleDbCommandBuilder objCommandBuilder;
        public DataSet MyData = new DataSet();//不能设置为静态，，fill一次增加一次数据，，，记录条数翻倍
        public object MyData_delete;
        public int i = 0;
        public _202_f()
        {
            InitializeComponent();
          
            text_a0 = this.a0;
            MyData_delete = MyData;
            Control.CheckForIllegalCrossThreadCalls = false;
           MyConn = new OleDbConnection(ConnString);
            MyConn.Open();
            MyAd = new OleDbDataAdapter("select *  from bsbyy2 ", MyConn);
            MyAd.Fill(MyData);
            objCommandBuilder = new OleDbCommandBuilder(MyAd);
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\202");

            panel1.Enabled = false;

            button8.Enabled = false;
            button9.Enabled = false;
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                textBox17.Text = (MyData.Tables[0].Rows.Count).ToString();
                textBox18.Text = i.ToString();

                Record_show();
                textBox18.Text = (i + 1).ToString();

            }
            else
            {
                textBox17.Text = "0";
                textBox18.Text = "0";

            }

        }

        private void _202_f_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (i > 0)
            {
                i = i-1;
                Record_show();
                textBox18.Text = (i + 1).ToString();
            }
          
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if ((i < MyData.Tables[0].Rows.Count - 1) && (i != -1))
            {
                i++;
                Record_show();
                textBox18.Text = (i + 1).ToString();
            }
            else if (MyData.Tables[0].Rows.Count == 0)
            {
                i = -1;
                Record_null();
                textBox18.Text = null;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0)
            {
                i = 0;
                Record_show();
                textBox18.Text = (i + 1).ToString();
            }
            else
            {
                i = -1;
                Record_null();
                textBox18.Text = null;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count != 0)
            {
                i = MyData.Tables[0].Rows.Count - 1;
                Record_show();
                textBox18.Text = (i + 1).ToString();
            }
            else
            {
                i = -1;
                Record_null();
                textBox18.Text = null;
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel2.Enabled = false;



            /*  
              MyConn.Open();*/

            if (MyData.Tables[0].Rows.Count == 0)
                a0.Text = (20200001).ToString();
            else
            {
          
                OleDbDataAdapter MyAd00 = new OleDbDataAdapter("select a0  from bsbyy2 order by a0 desc ", MyConn);
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
            panel1.Enabled = true;


            button6.Enabled = false;
            button7.Enabled = false;

            button8.Enabled = true;
            button9.Enabled = true;

          
            button11.Enabled = false;
            button12.Enabled = false;
     
            button14.Enabled = false;
        
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

             
                button11.Enabled = false;
                button12.Enabled = false;
                
                button14.Enabled = false;
               

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            for (int v = 0; v < MyData.Tables[0].Rows.Count; v++)
            {

                //     MessageBox.Show("run");
                if (this.a0.Text == MyData.Tables[0].Rows[v][0].ToString())
                {

                    string SQL = "delete from bsbyy2 where a0= '" + this.a0.Text + "'";
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

            NewRow["time_e"] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();




            MyData.Tables[0].Rows.Add(NewRow);

            MyAd.Update(MyData);



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
            
            button11.Enabled = true;
            button12.Enabled = true;
       
            button14.Enabled = true;
   
            button16.Enabled = true;
            button17.Enabled = true;

            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select count(a0)  from bsbyy2 ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);

            this.textBox17.Text = MyData.Tables[0].Rows.Count.ToString();
            

        }

        private void save()
        {
            this.Enabled = false;
          

           // throw new NotImplementedException();
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\202\\" + a0.Text + "");

            string temp = System.IO.Directory.GetCurrentDirectory();// d当前运行路径

            //    MessageBox.Show(temp);
            string OrignFile;//word模板路径
            OrignFile = "\\baseDB\\商标书式\\2 异议\\02 撤回商标异议申请书.dot";

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


            object aa = temp + "\\" + main.companyName + "\\wjgl\\202\\" + a0.Text + "\\" + a0.Text + ".docx";

            object[] MyBM = new object[7];//创建一个书签数组

            for (int i = 0; i < 6; i++)//给书签数组赋值
                MyBM[i] = "a" + (i + 1).ToString();





            //给对应的书签位置写入数据
            doc.Bookmarks.get_Item(ref MyBM[0]).Range.Text = a1.Text;
            doc.Bookmarks.get_Item(ref MyBM[1]).Range.Text = a2.Text;
            doc.Bookmarks.get_Item(ref MyBM[2]).Range.Text = a3.Text;
            doc.Bookmarks.get_Item(ref MyBM[3]).Range.Text = a4.Text;
            doc.Bookmarks.get_Item(ref MyBM[4]).Range.Text = a5.Text;
            doc.Bookmarks.get_Item(ref MyBM[5]).Range.Text = a6.Text;
           



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
         
            button11.Enabled = true;
            button12.Enabled = true;
          
            button14.Enabled = true;
       button16.Enabled = true;
            button17.Enabled = true;
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                Record_show();
            }
            else
            {
                Record_null();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select *  from bsbyy2 ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);//获取记录条数 
            if (MyData0.Tables[0].Rows.Count != 0)
            {
                (new YNcleanALL_f()).ShowDialog();
                if (YNcleanALL_f.YNdel==true)
                {
                    CleanApp();
                textBox17.Text = (int.Parse(textBox17.Text) - 1).ToString();

                if (MyData.Tables[0].Rows.Count == 0)
                {
                    textBox18.Text = "0";
                    Record_null();
                }
                else if (i == MyData.Tables[0].Rows.Count)
                {
                    i = i - 1;
                    Record_show();
                    textBox18.Text = (int.Parse(textBox18.Text) - 1).ToString();
                }
                else
                {
Record_show();
                    //   textBox18.Text = (int.Parse(textBox18.Text) + 1).ToString();
                }
            }
        }
        }
        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(".\\" + main.companyName + "\\wjgl\\202");
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       public void CleanApp()
        {
            //内存数据删除
            ((DataSet)MyData_delete).Tables[0].Rows.RemoveAt(i);


            // int a =;// 数据库删除
            string SQL = "delete * from bsbyy2 where a0='" + ((TextBox)text_a0).Text + "'";
            OleDbCommand NewCom = new OleDbCommand(SQL,MyConn);
            NewCom.ExecuteNonQuery();//执行SQL语句
            //  MessageBox.Show(_101_f.i.ToString());


            //文件删除&& 


            if (File.Exists(@".\\" + main.companyName + "\\wjgl\\202\\" + ((TextBox)text_a0).Text + "\\" + ((TextBox)text_a0).Text + ".docx"))
            {
                // MessageBox.Show("存在文件");
                File.Delete(@".\\" + main.companyName + "\\wjgl\\202\\" + ((TextBox)text_a0).Text + "\\" + ((TextBox)text_a0).Text + ".docx");

            }

            if (Directory.Exists(@".\\" + main.companyName + "\\wjgl\\202\\" + ((TextBox)text_a0).Text + ""))
            {
                // MessageBox.Show("存在文件夹");
                Directory.Delete(@".\\" + main.companyName + "\\wjgl\\202\\" + ((TextBox)text_a0).Text + "");
            }
            MessageBox.Show("删除成功");
        }

       private void button12_Click(object sender, EventArgs e)
       {
           main.Print(MyData, a0.Text, FormId, 1);
       }

       private void button14_Click(object sender, EventArgs e)
       {
           main.Print(MyData, a0.Text, FormId, 2);
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
                
       }
       private void Record_null()
       {
           a0.Text = null;
           a1.Text = null;
           a2.Text = null;
           a3.Text = null;
           a4.Text = null;
           a5.Text = null;
           a6.Text = null;
       }
    }
}
