using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Threading;
using System.Diagnostics;



namespace MyApp
{

    public partial class _101_f :Form,FormFather
    {

        /*
         * 查询数据块-
         */
       public  Dictionary<string, string> Item_list = new Dictionary<string, string>();
     
         /* 
          * * -查询数据块
         */


        public  int i;//记录当前显示的数据编号
        private object text_a0;
 
        private string FormId="101";
      

        private string   ConnString = "Provider="+main.Office_Engen+";Data Source=.\\"+main.companyName+"\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password="+main.appDB_password+"";
        private OleDbConnection MyConn;
        private  OleDbDataAdapter MyAd;
        private  DataSet MyData;
        private  OleDbCommandBuilder MyComm;
        private string DB_table_name = "1sbzcsq";
    
        public _101_f()
        {
            InitializeComponent();
            /*
             * 添加查询条目
             */
            Item_list.Add("a0","内部管理编号");
          
            /*
             * ——添加查询条目
             */
            main.text_19 = this.a19;
            main.text_20 = this.a20;
            main.text_21 = this.a21;
            main.text_22 = this.a22;
     
           text_a0 = this.a0;
          
            MyConn = new OleDbConnection(ConnString);
           MyConn.Open();
            MyAd = new OleDbDataAdapter("select * from "+DB_table_name+" ", MyConn);
           MyData = new DataSet();
           MyComm = new OleDbCommandBuilder(MyAd);
           MyAd.Fill(MyData);//数据表
            if (MyData.Tables[0].Rows.Count == 0)
                i = -1;
            else
                i = 0;
           
        }
    private void save()
        {
            
            this.Enabled = false;
            Directory.CreateDirectory(".\\"+main.companyName+"\\wjgl\\101\\" + a0.Text + "");

            string temp = System.IO.Directory.GetCurrentDirectory();// d当前运行路径

            //    MessageBox.Show(temp);
            string OrignFile;
            OrignFile = "\\baseDB\\商标书式\\1 注册申请\\01 商标注册申请书.dot";
  
            //开始写入数据

            int a13_1_a = 0, a13_2_a = 0, a13_3_a = 0, a13_4_a = 0, a13_5_a = 0, a13_6_a = 0;
            if (a13_1.CheckState == CheckState.Checked)
                a13_1_a = 1;
            if (a13_2.CheckState == CheckState.Checked)
                a13_2_a = 1;
            if (a13_3.CheckState == CheckState.Checked)
                a13_3_a = 1;
            if (a13_4.CheckState == CheckState.Checked)
                a13_4_a = 1;
            if (a13_5.CheckState == CheckState.Checked)
                a13_5_a = 1;
            if (a13_6.CheckState == CheckState.Checked)
                a13_6_a = 1;



      

            string parFilePath = temp + OrignFile;//文件路径
            object FilePath = parFilePath;
            Microsoft.Office.Interop.Word._Application AppliApp = new Microsoft.Office.Interop.Word.Application();
            AppliApp.Visible = false;
            Microsoft.Office.Interop.Word._Document doc = AppliApp.Documents.Add(ref FilePath);
            object missing = System.Reflection.Missing.Value;
            object isReadOnly = false;
          



            doc.Activate();

            //数据写入代码段


            object aa = temp + "\\"+main.companyName+"\\wjgl\\101\\" + a0.Text + "\\" + a0.Text + ".docx";//命名更改

            object[] MyBM = new object[30];//创建一个书签数组

            for (int i = 0; i < 12; i++)//给书签数组赋值
                MyBM[i] = "a" + (i + 1).ToString();
            MyBM[12] = "a13_1";
            MyBM[13] = "a13_2";
            MyBM[14] = "a13_3";
            MyBM[15] = "a13_4";
            MyBM[16] = "a13_5";
            MyBM[17] = "a13_6";
            MyBM[18] = "a14_1";
            MyBM[19] = "a14_2";
            MyBM[20] = "a14_3";

            MyBM[21] = "a15";
            MyBM[22] = "a16";
            MyBM[23] = "a17";
            MyBM[24] = "a18";
            MyBM[25] = "a19";
            MyBM[26] = "a20";
            MyBM[27] = "a21";
            MyBM[28] = "a22";
            MyBM[29] = "a25";



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
            string xx;
            if (a13_1_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[12]).Range.Text = xx;
            if (a13_2_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[13]).Range.Text = xx;
            if (a13_3_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[14]).Range.Text = xx;
            if (a13_4_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[15]).Range.Text = xx;
            if (a13_5_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[16]).Range.Text = xx;
            if (a13_6_a == 1)

                xx = "✔";
            else
                xx = " ";
            doc.Bookmarks.get_Item(ref MyBM[17]).Range.Text = xx;

            switch (a14.Text)
            {
                case "基于第一次申请的优先权":
                    doc.Bookmarks.get_Item(ref MyBM[18]).Range.Text = "✔";
                    break;
                case "基于展会的优先权":
                    doc.Bookmarks.get_Item(ref MyBM[19]).Range.Text = "✔";
                    break;
                case "优先权证明文件后补":
                    doc.Bookmarks.get_Item(ref MyBM[20]).Range.Text = "✔";
                    break;
                default:
                    break;
            }


            doc.Bookmarks.get_Item(ref MyBM[21]).Range.Text = a15.Text;
            doc.Bookmarks.get_Item(ref MyBM[22]).Range.Text = a16.Text;
            doc.Bookmarks.get_Item(ref MyBM[23]).Range.Text = a17.Text;
            doc.Bookmarks.get_Item(ref MyBM[24]).Range.Text = a18.Text;
            doc.Bookmarks.get_Item(ref MyBM[25]).Range.Text = a19.Text;
            doc.Bookmarks.get_Item(ref MyBM[26]).Range.Text = a20.Text;
            doc.Bookmarks.get_Item(ref MyBM[27]).Range.Text = a21.Text;
            doc.Bookmarks.get_Item(ref MyBM[28]).Range.Text = a22.Text;
            doc.Bookmarks.get_Item(ref MyBM[29]).Range.Text = a25.Text;









            doc.SaveAs(ref aa);
            doc.Close();
            this.Enabled = true;

        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void _501_f_Load(object sender, EventArgs e)
        {
            
            CheckForIllegalCrossThreadCalls = false;
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            Directory.CreateDirectory(".\\" + main.companyName + "\\wjgl\\101");
            a16.Text=DateTime.Now.ToShortDateString();

            panel1.Enabled=false;

            button8.Enabled = false;
            button9.Enabled = false;

            if (MyData.Tables[0].Rows.Count >= 1)
            {
                Record_show();

                textBox18.Text = (i + 1).ToString();

            }
            else
            {
                textBox18.Text = (0).ToString();
                
            }
        
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select count(a0)  from "+DB_table_name+" ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);

            this.textBox17.Text=MyData0.Tables[0].Rows[0][0].ToString();

          //  this.textBox18.Text = (i+1).ToString();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            main.Print(MyData, a0.Text,FormId, 2);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            panel2.Enabled = false;
           
         

        
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select count(a0)  from "+DB_table_name+"", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);
            if (int.Parse(MyData0.Tables[0].Rows[0][0].ToString()) == 0)
                a0.Text = (10100001).ToString();
            else 
            {
                OleDbDataAdapter MyAd00 = new OleDbDataAdapter("select a0  from "+DB_table_name+" order by a0 desc ", MyConn);
                DataSet MyData00 = new DataSet();
                MyAd00.Fill(MyData00);
                a0.Text=(int.Parse(MyData00.Tables[0].Rows[0][0].ToString())+1).ToString();
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
            a9.Text = null;
            a10.Text = null;
            a11.Text = null;
            a12.Text = null;
            a15.Text = null;
            a16.Text = DateTime.Now.ToShortDateString();
            a17.Text = null;
            a19.Text = null;
            a21.Text = null;
            a18.Text = null;
           a20.Text = null;
            a25.Text = null;
          a22.Text = null;
            panel1.Enabled = true;


            button6.Enabled = false;
            button7.Enabled = false;

            button8.Enabled = true;
            button9.Enabled = true;

            button10.Enabled = false;
            button11.Enabled = false;
            button12.Enabled = false;
            button13.Enabled = false;
            button14.Enabled = false;
            button15.Enabled = false;

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
            button13.Enabled = true;
            button14.Enabled = true;
            button15.Enabled = true;
            button16.Enabled = true;
            button17.Enabled = true;
            if (MyData.Tables[0].Rows.Count > 0)
            {
                Record_show();

                textBox18.Text = (i + 1).ToString();
            }
            else
            {
                Record_null();
                textBox18.Text = null;
            }


        }

        private void button8_Click(object sender, EventArgs e)
        {
             
            string SQL = "delete from "+DB_table_name+" where a0= '"+this.a0.Text+"'";
            OleDbCommand MyCom = new OleDbCommand(SQL,MyConn);
            MyCom.ExecuteNonQuery();
              //     MyData.Tables[0].Rows.Remove(MyData.Tables[0].Rows[0]);
            for (int v = 0; v < MyData.Tables[0].Rows.Count;v++ )
            {
               
               
                if(this.a0.Text==MyData.Tables[0].Rows[v][0].ToString())
                {  
                    MessageBox.Show("run");
                  
                   MyData.Tables[0].Rows.RemoveAt(v);
                   break;
                }
            }
            panel2.Enabled = true;
             int a13_1_a = 0, a13_2_a = 0, a13_3_a = 0, a13_4_a = 0, a13_5_a = 0, a13_6_a = 0;
            if (a13_1.CheckState == CheckState.Checked)
                a13_1_a = 1;
            if (a13_2.CheckState == CheckState.Checked)
                a13_2_a = 1;
            if (a13_3.CheckState == CheckState.Checked)
                a13_3_a = 1;
            if (a13_4.CheckState == CheckState.Checked)
                a13_4_a = 1;
            if (a13_5.CheckState == CheckState.Checked)
                a13_5_a = 1;
            if (a13_6.CheckState == CheckState.Checked)
                a13_6_a = 1;

            /*
            *word  存储开始
            *
            */
            //代码域
           
     
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
            NewRow["a1"]=a1.Text;
            NewRow["a2"] = a2.Text;
            NewRow["a3"] = a3.Text;
            NewRow["a4"] = a4.Text;
            NewRow["a5"] = a5.Text;
            NewRow["a6"] = a6.Text;
            NewRow["a7"] = a7.Text;
            NewRow["a8"] = a8.Text;
            NewRow["a9"] = a9.Text;
            NewRow["a10"] = a10.Text;
            NewRow["a11"] = a11.Text;
            NewRow["a12"] = a12.Text;
            NewRow["a13"] = a13_1_a.ToString()+a13_2_a.ToString()+a13_3_a.ToString()+a13_4_a.ToString()+a13_5_a.ToString()+a13_6_a.ToString();
            NewRow["a14"] = a14.Text;
            NewRow["a15"] = a15.Text;
            NewRow["a16"] = a16.Text;
            NewRow["a17"] = a17.Text;
            NewRow["a18"] = a18.Text;
            NewRow["a19"] = a19.Text;
            NewRow["a20"] = a20.Text;
            NewRow["a21"] = a21.Text;
            NewRow["a22"] = a22.Text;
            NewRow["a25"] = a25.Text;
            NewRow["time_e"] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
            //MessageBox.Show(MyConn.State.ToString());
          
            MyData.Tables[0].Rows.Add(NewRow);
            MyAd.Update(MyData);
           
           
      
            //MessageBox.Show(a13_1.CheckState);



            /*
             * 数据库的存储结束
             *
             * *
             */





         
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
            button13.Enabled = true;
            button14.Enabled = true;
            button15.Enabled = true;
            button16.Enabled = true;
            button17.Enabled = true;

            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select count(a0)  from "+DB_table_name+" ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);

            this.textBox17.Text = MyData0.Tables[0].Rows[0][0].ToString();       
             
               
            MessageBox.Show("保存成功！！");
                  
               
           
           
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if(MyData.Tables[0].Rows.Count>=1)
            {
                panel1.Enabled = true;


                button6.Enabled = false;
                button7.Enabled = false;

                button8.Enabled = true;
                button9.Enabled = true;

                button10.Enabled = false;
                button11.Enabled = false;
                button12.Enabled = false;
                button13.Enabled = false;
                button14.Enabled = false;
                button15.Enabled = false;

            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            uniquelistone.message=1;
           uniquelistone.cli =a19.Text;
            
                (new uniquelistone(this,this.GetType())).ShowDialog();
           
        }

        private void button19_Click(object sender, EventArgs e)
        {
            uniquelistone.message = 2;
           uniquelistone.cli = a21.Text;
                (new uniquelistone(this,this.GetType())).ShowDialog();
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
      
            OleDbDataAdapter MyAd = new OleDbDataAdapter("select * from "+DB_table_name+"", MyConn);
            DataSet MyData = new DataSet();
            MyAd.Fill(MyData);

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
        private void button3_Click(object sender, EventArgs e)
        {
            if (i>0)
            {
                i = i - 1;

                Record_show();

                textBox18.Text = (i + 1).ToString();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select *  from "+DB_table_name+" ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);//获取记录条数  
            if ((i < MyData0.Tables[0].Rows.Count-1)&&(i!=-1))
            {
                i++;


                Record_show();

                textBox18.Text = (i + 1).ToString();
            }
            else if(MyData0.Tables[0].Rows.Count==0)
            {
                i = -1;
                Record_null();

                textBox18.Text = null;
            }

        }
       
        private void button16_Click(object sender, EventArgs e)
        {
            Process.Start(".\\"+main.companyName+"\\wjgl\\101");
        }

        private void button5_Click(object sender, EventArgs e)
        {

            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select *  from "+DB_table_name+" ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);//获取记录条数 
            if (MyData0.Tables[0].Rows.Count != 0)
            {
                i = MyData0.Tables[0].Rows.Count - 1;
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
        public void CleanApp()
        {

            //内存数据删除
            MyData.Tables[0].Rows.RemoveAt(i);


            // int a =;// 数据库删除
            string SQL = "delete * from "+DB_table_name+" where a0='" + ((TextBox)text_a0).Text + "'";
            OleDbCommand NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();//执行SQL语句
            //  MessageBox.Show(_101_f.i.ToString());


            //文件删除&& 


            if (File.Exists(@".\\" + main.companyName + "\\wjgl\\101\\" + ((TextBox)text_a0).Text + "\\" + ((TextBox)text_a0).Text + ".docx"))
            {
                //    MessageBox.Show("存在文件");
                File.Delete(@".\\" + main.companyName + "\\wjgl\\101\\" + ((TextBox)text_a0).Text + "\\" + ((TextBox)text_a0).Text + ".docx");

            }

            if (Directory.Exists(@".\\" + main.companyName + "\\wjgl\\101\\" + ((TextBox)text_a0).Text + ""))
            {
                //  MessageBox.Show("存在文件夹");
                Directory.Delete(@".\\" + main.companyName + "\\wjgl\\101\\" + ((TextBox)text_a0).Text + "");
            }
            MessageBox.Show("删除成功");
        }
        private void button11_Click(object sender, EventArgs e)
        {
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select *  from "+DB_table_name+" ", MyConn);
            DataSet MyData0 = new DataSet();
            MyAd0.Fill(MyData0);//获取记录条数 
            if (MyData0.Tables[0].Rows.Count != 0)
            {
                (new YNcleanALL_f()).ShowDialog();
                if (YNcleanALL_f.YNdel = true)
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
                        a9.Text = null;
                        a10.Text = null;
                        a11.Text = null;
                        a12.Text = null;
                        a15.Text = null;
                        a16.Text = DateTime.Now.ToShortDateString();
                        a17.Text = null;
                        a19.Text = null;
                        a21.Text = null;
                        a18.Text = null;
                        a20.Text = null;
                        a25.Text = null;
                        a22.Text = null;
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
                        textBox18.Text = (int.Parse(textBox18.Text) + 1).ToString();
                    }
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
         if (MyData.Tables[0].Rows.Count>=1)
            {
                panel2.Enabled = false;



           
                OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select top 1 * from "+DB_table_name+" order by a0 desc", MyConn);
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
                button13.Enabled = false;
                button14.Enabled = false;
                button15.Enabled = false;
            }
        }

        private void a0_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void label29_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            main.Print(MyData,a0.Text,FormId,1);
             
        }

        private void SetDefaultPrinter(string p)
        {
            
        }

        unsafe private void button1_Click(object sender, EventArgs e)
        {
            if (MyData.Tables[0].Rows.Count > 0)
            {
                fixed (int* pp = &i)
                {
                    Each_f_statistics a = new Each_f_statistics(DB_table_name,MyConn,pp, MyData.Tables[0], Item_list);
                    a.ShowDialog();
                }
              
                    Record_show();
                    textBox18.Text = (i + 1).ToString();
 
            }
        }



        public string defaultPrinter { get; set; }

        private void button13_Click(object sender, EventArgs e)
        {
            main.Print(MyData, a0.Text, FormId, 3);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            main.Print(MyData, a0.Text, FormId, 4);
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
            if (MyData.Tables[0].Rows[i][13].ToString()[0] == '1')
                a13_1.CheckState = CheckState.Checked;
            else
                a13_1.CheckState = CheckState.Unchecked;
            if (MyData.Tables[0].Rows[i][13].ToString()[1] == '1')
                a13_2.CheckState = CheckState.Checked;
            else
                a13_2.CheckState = CheckState.Unchecked;
            if (MyData.Tables[0].Rows[i][13].ToString()[2] == '1')
                a13_3.CheckState = CheckState.Checked;
            else
                a13_3.CheckState = CheckState.Unchecked;
            if (MyData.Tables[0].Rows[i][13].ToString()[3] == '1')
                a13_4.CheckState = CheckState.Checked;
            else
                a13_4.CheckState = CheckState.Unchecked;
            if (MyData.Tables[0].Rows[i][13].ToString()[4] == '1')
                a13_5.CheckState = CheckState.Checked;
            else
                a13_5.CheckState = CheckState.Unchecked;
            if (MyData.Tables[0].Rows[i][13].ToString()[5] == '1')
                a13_6.CheckState = CheckState.Checked;
            else
                a13_6.CheckState = CheckState.Unchecked;

            a14.Text = MyData.Tables[0].Rows[i][14].ToString();

            a15.Text = MyData.Tables[0].Rows[i][15].ToString();
            a16.Text = MyData.Tables[0].Rows[i][16].ToString();
            a17.Text = MyData.Tables[0].Rows[i][17].ToString();
            a18.Text = MyData.Tables[0].Rows[i][18].ToString();
            a19.Text = MyData.Tables[0].Rows[i][19].ToString();
            a20.Text = MyData.Tables[0].Rows[i][20].ToString();
            a21.Text = MyData.Tables[0].Rows[i][21].ToString();
            a22.Text = MyData.Tables[0].Rows[i][22].ToString();

            a25.Text = MyData.Tables[0].Rows[i][23].ToString();
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
            a7.Text = null;
            a8.Text = null;
            a9.Text = null;
            a10.Text = null;
            a11.Text = null;
            a12.Text = null;

            a13_1.CheckState = CheckState.Unchecked;


            a13_2.CheckState = CheckState.Unchecked;

            a13_3.CheckState = CheckState.Unchecked;

            a13_4.CheckState = CheckState.Unchecked;

            a13_5.CheckState = CheckState.Unchecked;

            a13_6.CheckState = CheckState.Unchecked;

            a14.Text = null;

            a15.Text = null;
            a16.Text = null;
            a17.Text = null;
            a18.Text = null;
            a19.Text = null;
            a20.Text = null;
            a21.Text = null;
            a22.Text = null;

            a25.Text = null;
        }

    }
}
