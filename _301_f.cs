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
    public partial class _301_f : Form,FormFather
    {
        /*
* 查询数据块-
*/
        public Dictionary<string, string> Item_list = new Dictionary<string, string>() { { "a0", "内部管理编号" } };

        /* 
         * * -查询数据块
        */
        private string DB_table_name = "c1bfsq";
        private string FormId="301";
        public  object text_a0;//传A0的值

        //数据库链接
        public string ConnString = "Provider="+main.Office_Engen+";Data Source=.\\" + main.companyName + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password="+main.appDB_password+"";
        public  OleDbConnection MyConn;

        public  OleDbDataAdapter MyAd;
        public OleDbCommandBuilder objCommandBuilder;
        public DataSet MyData = new DataSet();//不能设置为静态，，fill一次增加一次数据，，，记录条数翻倍
        public object MyData_delete;
        public int i = 0;
        public _301_f()
        {
            InitializeComponent();
            text_a0 = this.a0;
            MyData_delete = MyData;
            Control.CheckForIllegalCrossThreadCalls = false;
            MyConn = new OleDbConnection(ConnString);
            MyConn.Open();
            MyAd = new OleDbDataAdapter("select *  from c1bfsq ", MyConn);
            MyAd.Fill(MyData);
            objCommandBuilder = new OleDbCommandBuilder(MyAd);
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\301");

            /*
             *数据返回器赋给当前form地址 *
             */
           


            panel1.Enabled = false;

            button8.Enabled = false;
            button9.Enabled = false;

        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(".\\" + main.companyName + "\\wjgl\\301");
        }

        private void _301_f_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);

          


            if (MyData.Tables[0].Rows.Count >= 1)
            {
                textBox17.Text = (MyData.Tables[0].Rows.Count).ToString();
                //textBox18.Text = i.ToString();

                Record_show();
                textBox18.Text = (i + 1).ToString();

            }
            else
            {
                textBox17.Text = "0";
                textBox18.Text = "0";

            }
          

        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel2.Enabled = false;



            /*  
              MyConn.Open();*/

            if (MyData.Tables[0].Rows.Count == 0)
                a0.Text = (30100001).ToString();
            else
            {
               
                OleDbDataAdapter MyAd00 = new OleDbDataAdapter("select a0  from c1bfsq order by a0 desc ", MyConn);
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
            a7.Text = null;
                
                    a8_1.CheckState = CheckState.Unchecked;
               
              
                    a8_2.CheckState = CheckState.Unchecked;
                
               
                    a8_3.CheckState = CheckState.Unchecked;

                   
                //   MessageBox.Show(MyData.Tables[0].Rows.Count.ToString());
 
            a9.Enabled=false;
                    a10.Enabled=false;
                    a11.Enabled=false;
                    a12.Enabled=false ;
                    a13.Enabled=false;

                panel1.Enabled = true;


                button6.Enabled = false;
                button7.Enabled = false;

                button8.Enabled = true;
                button9.Enabled = true;

                button10.Enabled = false;
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

                button10.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
          
                button14.Enabled = false;
             

            }
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

        private void button8_Click(object sender, EventArgs e)
        {
            for (int v = 0; v < MyData.Tables[0].Rows.Count; v++)
            {

                //     MessageBox.Show("run");
                if (this.a0.Text == MyData.Tables[0].Rows[v][0].ToString())
                {

                    string SQL = "delete from c1bfsq where a0= '" + this.a0.Text + "'";
                    OleDbCommand MyCom = new OleDbCommand(SQL, MyConn);
                    MyCom.ExecuteNonQuery();
                    MyData.Tables[0].Rows.RemoveAt(v);
                    break;
                }
            }

            panel2.Enabled = true;
            int a8_1_a = 0, a8_2_a = 0, a8_3_a = 0;
            if (a8_1.CheckState == CheckState.Checked)
                a8_1_a = 1;
            if (a8_2.CheckState == CheckState.Checked)
                a8_2_a = 1;
            if (a8_3.CheckState == CheckState.Checked)
                a8_3_a = 1;
           
            /*
            *word  存储开始
            *
            */
            //代码域
            int a;
            Thread MySave = new Thread(new ThreadStart(save));
            MySave.Start();

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
            NewRow["a8"] = a8_1_a.ToString() + a8_2_a.ToString() + a8_3_a.ToString();
        //    NewRow["a8"] = a8.Text;
            NewRow["a9"] = a9.Text;
            NewRow["a10"] = a10.Text;
            NewRow["a11"] = a11.Text;
            NewRow["a12"] = a12.Text;
            NewRow["a13"] = a12.Text;
            NewRow["time_e"] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
          
            //MessageBox.Show(MyConn.State.ToString());

            MyData.Tables[0].Rows.Add(NewRow);
            MyAd.Update(MyData);
            MyAd.Dispose();


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
          
            button14.Enabled = true;
         
            button16.Enabled = true;
            button17.Enabled = true;

            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select count(a0)  from c1bfsq ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);

            this.textBox17.Text = MyData0.Tables[0].Rows[0][0].ToString();


          
                  
               
           
        }

        private void save()
        {
            this.Enabled = false;
          
           // thw new NotImplementedException();
            int a8_1_a = 0, a8_2_a = 0, a8_3_a = 0;
            if (a8_1.CheckState == CheckState.Checked)
                a8_1_a = 1;
            if (a8_2.CheckState == CheckState.Checked)
                a8_2_a = 1;
            if (a8_3.CheckState == CheckState.Checked)
                a8_3_a = 1;
           

            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\301\\" + a0.Text + "");

            string temp = System.IO.Directory.GetCurrentDirectory();// d当前运行路径

            //    MessageBox.Show(temp);
            string OrignFile;
            OrignFile = "\\baseDB\\商标书式\\5 补证撤三出具证明\\01 补发变更转让续展证明申请书.dot";

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


            object aa = temp + ".\\" + main.companyName + "\\wjgl\\301\\" + a0.Text + "\\" + a0.Text + ".docx";//命名更改

            object[] MyBM = new object[15];//创建一个书签数组

            for (int i = 0; i < 7; i++)//给书签数组赋值
                MyBM[i] = "a" + (i + 1).ToString();
            MyBM[7] = "a8_1";
            MyBM[8] = "a8_2";
            MyBM[9] = "a8_3";
            MyBM[10] = "a9";
            MyBM[11] = "a10";
            MyBM[12] = "a11";
            MyBM[13] = "a12";
            MyBM[14] = "a13";
          



            //给对应的书签位置写入数据
            doc.Bookmarks.get_Item(ref MyBM[0]).Range.Text = a1.Text;
            doc.Bookmarks.get_Item(ref MyBM[1]).Range.Text = a2.Text;
            doc.Bookmarks.get_Item(ref MyBM[2]).Range.Text = a3.Text;
            doc.Bookmarks.get_Item(ref MyBM[3]).Range.Text = a4.Text;
            doc.Bookmarks.get_Item(ref MyBM[4]).Range.Text = a5.Text;
            doc.Bookmarks.get_Item(ref MyBM[5]).Range.Text = a6.Text;
            doc.Bookmarks.get_Item(ref MyBM[6]).Range.Text = a7.Text;
         


 
            string xx;
            if (a8_1_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[7]).Range.Text = xx;
            if (a8_2_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[8]).Range.Text = xx;
            if (a8_3_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[9]).Range.Text = xx;
           

            
            doc.Bookmarks.get_Item(ref MyBM[10]).Range.Text = a9.Text;
            doc.Bookmarks.get_Item(ref MyBM[11]).Range.Text = a10.Text;
            doc.Bookmarks.get_Item(ref MyBM[12]).Range.Text = a11.Text;
            doc.Bookmarks.get_Item(ref MyBM[13]).Range.Text = a12.Text;
            doc.Bookmarks.get_Item(ref MyBM[14]).Range.Text = a13.Text;
          


           









            doc.SaveAs(ref aa);
            doc.Close();
  this.Enabled = true;

        }

        private void a8_1_CheckedChanged(object sender, EventArgs e)
        {
            if (a8_1.CheckState == CheckState.Checked)
            {
 
                a9.Enabled = true;
                a10.Enabled =true;
            }
            else
            {
                a9.Enabled = false;
                a10.Enabled =false;
            }
           
            /*a11.Enabled = true;
            a12.Enabled = true;
            a13.Enabled = true;*/

        }

        private void a8_2_CheckedChanged(object sender, EventArgs e)
        {
            if (a8_2.CheckState == CheckState.Checked)
            {

                a11.Enabled = true;
                a12.Enabled = true;
            }
            else
            {
                a11.Enabled = false;
                a12.Enabled = false;
            }
        }

        private void a8_3_CheckedChanged(object sender, EventArgs e)
        {
            if (a8_3.CheckState == CheckState.Checked)
            {

              
                a13.Enabled = true;
            }
            else
            {
               
                a13.Enabled = false;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select *  from c1bfsq ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);//获取记录条数 
            if (MyData0.Tables[0].Rows.Count != 0)
            {
                (new YNcleanALL_f()).ShowDialog();
                if (YNcleanALL_f.YNdel == true)
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
                    }
                }
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count >= 1)
            {
                panel2.Enabled = false;          
                OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select top 1 * from c1bfsq order by a0 desc", MyConn);
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
               
                button14.Enabled = false;
              
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (i > 0)
            {
                i = i - 1;
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

       public void CleanApp()
        {  //内存数据删除
            ((DataSet)MyData_delete).Tables[0].Rows.RemoveAt(i);


            // int a =;// 数据库删除
            string SQL = "delete * from c1bfsq where a0='" + ((TextBox)text_a0).Text + "'";
            OleDbCommand NewCom = new OleDbCommand(SQL,MyConn);
            NewCom.ExecuteNonQuery();//执行SQL语句
            //  MessageBox.Show(_101_f.i.ToString());


            //文件删除&& 


            if (File.Exists(@".\\" + main.companyName + "\\wjgl\\301\\" + ((TextBox)text_a0).Text + "\\" + ((TextBox)text_a0).Text + ".docx"))
            {
                //    MessageBox.Show("存在文件");
                File.Delete(@".\\" + main.companyName + "\\wjgl\\301\\" + ((TextBox)text_a0).Text + "\\" + ((TextBox)text_a0).Text + ".docx");

            }

            if (Directory.Exists(@".\\" + main.companyName + "\\wjgl\\301\\" + ((TextBox)text_a0).Text + ""))
            {
                //  MessageBox.Show("存在文件夹");
                Directory.Delete(@".\\" + main.companyName + "\\wjgl\\301\\" + ((TextBox)text_a0).Text + "");
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
           a7.Text = MyData.Tables[0].Rows[i][7].ToString();

           a9.Text = MyData.Tables[0].Rows[i][9].ToString();
           a10.Text = MyData.Tables[0].Rows[i][10].ToString();
           a11.Text = MyData.Tables[0].Rows[i][11].ToString();
           a12.Text = MyData.Tables[0].Rows[i][12].ToString();
           if (MyData.Tables[0].Rows[i][8].ToString()[0] == '1')
               a8_1.CheckState = CheckState.Checked;
           else
               a8_1.CheckState = CheckState.Unchecked;
           if (MyData.Tables[0].Rows[i][8].ToString()[1] == '1')
               a8_2.CheckState = CheckState.Checked;
           else
               a8_2.CheckState = CheckState.Unchecked;
           if (MyData.Tables[0].Rows[i][8].ToString()[2] == '1')
               a8_3.CheckState = CheckState.Checked;
           else
               a8_3.CheckState = CheckState.Unchecked;

           a13.Text = MyData.Tables[0].Rows[i][13].ToString();
       }
       private void Record_null()
       {
           a1.Text = null;
           a2.Text = null;
           a3.Text = null;
           a4.Text = null;
           a5.Text = null;
           a6.Text = null;
           a7.Text = null;
           a8_1.CheckState = CheckState.Unchecked;
           a8_2.CheckState = CheckState.Unchecked;
           a8_3.CheckState = CheckState.Unchecked;
           a9.Text = null;        
           a10.Text = null;
           a11.Text = null;
           a12.Text = null;
           a13.Text = null;
       }
    }
}
