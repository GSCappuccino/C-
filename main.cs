using System;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Data.OleDb;
using System.Threading;
using Microsoft.Win32;

namespace MyApp
{
    
    public partial class main : Form
    {
        /*
         * 数据库的连接信息
         */
        public static string Office_Engen = null;
        public static string Office_Version_Id = null;
            //"Microsoft.ACE.OLEDB.12.0";
        public static string sj_path=".\\baseDB\\sj.mdb";
        
        public static string appDB_path = ".\\baseDB\\appDB.mdb";
        public static string appDB_password = "fgf%ifdjAdgdlk";
        public static string sjDB_password = "123jiu@jjk,kk";
        public static string UserDB_password = sjDB_password;
        public static string User_path=sj_path;
        public static string companyName;//公司名字
        /*
        静态数据，，，传回数据接收器
         * 区分表//、数据
        */
       // public static _101_f RuningForm;

        public static object text_19;
        public static object text_20;
        public static object text_21;
        public static object text_22;

       //登陆者信息
        public static string UserId;//下方显示  使用者用户名
        public static string User_power;//使用者权限

        public static string User_password;//使用者密码


        public static System.DateTime currentTime = new DateTime();

      //  public static string Data;
        public main()
        {
            currentTime = System.DateTime.Now;
            InitializeComponent();
           
            
        }

        //数据清除
        public void AppSQLdelete()
        {
            this.Enabled = false;
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + main.companyName + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);
            MyConn.Open();

            string SQL = "delete from 1sbzcsq";
            OleDbCommand NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();//执行SQL语句

            SQL = "delete from 1sbzcsq2 ";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from achsq3";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from ayxzm4";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from bsbyy1";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from bsbyy2";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c1bfsq";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c2bfzc";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c3cjzc";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c4gzzc";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c5snby";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c6cxsn";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c7cxsp";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from c8chcx";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from dtsbz";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();


            SQL = "delete from e11sbxz";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e12chxz";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e21bgsb";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e22chsb";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e24chdl";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e23bgdl";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e31sjsp";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e32chsj";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e41zryz";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e42chrz";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e51xkba";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from e52bgmc";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e53tjzz";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e54chxk";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e61zrzy";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e62zqbg";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e63djyq";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e64bfsq";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e65zyzx";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e71sbzx";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from e72cxzx";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from f81wts";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            SQL = "delete from g1bhfs";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g2byfs";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g3cxfs";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g4wxfs";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g5wxsq";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g6yyfs";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g7chps";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from g8fswt";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();

            
            MyConn.Close();
            this.Enabled = true;
        }
        public  void CleanApp()
        {
            AppSQLdelete();
            
            if (Directory.Exists(@".\\" + main.companyName + "\\wjgl"))
            {
                System.IO.Directory.Delete(".\\" + main.companyName + "\\wjgl", true);
            }


        }
        //——数据清除
     
        private void main_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            


            
            textBox2.Text=UserId;
            textBox4.Text = User_power;
            textBox6.Text = "武汉市金中红商标代理有限公司";
            textBox8.Text=DateTime.Now.ToString();
            
            panelthree.Hide();
       /*     panelone.Show();
            paneltwo.Hide();*/

        
        }

        private void label4_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void label3_Click(object sender, EventArgs e)
        {
            this.Hide();
            (new login()).ShowDialog();
            this.Show();
           
        }

        private void label2_Click(object sender, EventArgs e)
        {
            (new changepassword()).ShowDialog(); 
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(button1.Text=="商标书式")
               Process.Start(".\\baseDB\\商标书式");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button1.Text = "商标书式";
         
        }

      

      

       
        private void _101_Click(object sender, EventArgs e)
        {
         
           _101_f f_501= new  _101_f();
          (f_501).ShowDialog();
        
        }

       

        private void _102_Click(object sender, EventArgs e)
        {
            _102_f f_102 = new _102_f();
            f_102.ShowDialog();
           
        }

        private void _103_Click(object sender, EventArgs e)
        {
           
            _103_f f_103 = new _103_f();
            f_103.ShowDialog();
           
        }

        private void _104_Click(object sender, EventArgs e)
        {
           
            _104_f f_104 = new _104_f();
            f_104.ShowDialog();
          
        }

        private void _201_Click(object sender, EventArgs e)
        {
          
            _201_f f_201 = new _201_f();
            f_201.ShowDialog();
          
        }

        private void _202_Click(object sender, EventArgs e)
        {
         
            _202_f f_202 = new _202_f();
            f_202.ShowDialog();
          
        }

        private void _4_Click(object sender, EventArgs e)
        {
         
            _4_f f_4 = new _4_f();
            f_4.ShowDialog();
        
        }

        private void _308_Click(object sender, EventArgs e)
        {
          
            _308_f f_308 = new _308_f();
            f_308.ShowDialog();
          
        }

        private void _307_Click(object sender, EventArgs e)
        {
          
            _307_f f_307 = new _307_f();
            f_307.ShowDialog();
        }

        private void _306_Click(object sender, EventArgs e)
        {
         
            _306_f f_306 = new _306_f();
            f_306.ShowDialog();
          
        }

        private void _305_Click(object sender, EventArgs e)
        {
         
            _305_f f_305 = new _305_f();
            f_305.ShowDialog();
         
        }

        private void _304_Click(object sender, EventArgs e)
        {
           
            _304_f f_304 = new _304_f();
            f_304.ShowDialog();
          
        }

        private void _303_Click(object sender, EventArgs e)
        {
         
            _303_f f_303 = new _303_f();
            f_303.ShowDialog();
          
        }

        private void _302_Click(object sender, EventArgs e)
        {
         

            _302_f f_302 = new _302_f();
            f_302.ShowDialog();
         
        }

        private void _301_Click(object sender, EventArgs e)
        {
        
            _301_f f_301 = new _301_f();
            f_301.ShowDialog();
       
        }

        private void _511_Click(object sender, EventArgs e)
        {
        
            _511_f f_511 = new _511_f();
            f_511.ShowDialog();
        
        }

        private void _512_Click(object sender, EventArgs e)
        {
      
            _512_f f_512 = new _512_f();
            f_512.ShowDialog();
      
        }

        private void _521_Click(object sender, EventArgs e)
        {
  
            _521_f f_521 = new _521_f();
            f_521.ShowDialog();
           
        }

        private void _522_Click(object sender, EventArgs e)
        {
           
            _522_f f_522 = new _522_f();
            f_522.ShowDialog();
           
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

        }

        private void dataclean_Click(object sender, EventArgs e)
        {
           
           
        }

        private void statistics_Click(object sender, EventArgs e)
        {
          
        }

        private void _523_Click(object sender, EventArgs e)
        {
            _523_f f_523 = new _523_f();
            f_523.ShowDialog();
           
        }

        private void _524_Click(object sender, EventArgs e)
        {
            _524_f f_524 = new _524_f();
            f_524.ShowDialog();
        }

        private void _531_Click(object sender, EventArgs e)
        {
            _531_f f_531 = new _531_f();
            f_531.ShowDialog();
        }

        private void _532_Click(object sender, EventArgs e)
        {
            _532_f f_532 = new _532_f();
            f_532.ShowDialog();
        }

        private void _541_Click(object sender, EventArgs e)
        {
            _541_f f_541 = new _541_f();
            f_541.ShowDialog();
        }

        private void _542_Click(object sender, EventArgs e)
        {
            _542_f f_542 = new _542_f();
            f_542.ShowDialog();
        }

        private void _551_Click(object sender, EventArgs e)
        {
            _551_f f_551 = new _551_f();
            f_551.ShowDialog();
        }

        private void _552_Click(object sender, EventArgs e)
        {
            _552_f f_552 = new _552_f();
            f_552.ShowDialog();
        }

        private void _553_Click(object sender, EventArgs e)
        {
            _553_f f_553 = new _553_f();
            f_553.ShowDialog();
        }

        private void _554_Click(object sender, EventArgs e)
        {
            _554_f f_554 = new _554_f();
            f_554.ShowDialog();
        }

        private void _561_Click(object sender, EventArgs e)
        {
            _561_f f_561 = new _561_f();
            f_561.ShowDialog();
        }

        private void _562_Click(object sender, EventArgs e)
        {
            _562_f f_562 = new _562_f();
            f_562.ShowDialog();
        }

        private void _563_Click(object sender, EventArgs e)
        {
            _563_f f_563 = new _563_f();
            f_563.ShowDialog();
        }

        private void _564_Click(object sender, EventArgs e)
        {
            _564_f f_564 = new _564_f();
            f_564.ShowDialog();
        }

        private void _565_Click(object sender, EventArgs e)
        {
            _565_f f_565 = new _565_f();
            f_565.ShowDialog();
        }

        private void _571_Click(object sender, EventArgs e)
        {
            _571_f f_571 = new _571_f();
            f_571.ShowDialog();
        }

        private void _572_Click(object sender, EventArgs e)
        {
            _572_f f_572 = new _572_f();
            f_572.ShowDialog();
        }

        private void weotuobook_Click(object sender, EventArgs e)
        {
           
        }

        private void usermanage_Click(object sender, EventArgs e)
        {
           
        }

        private void distinctlist_Click(object sender, EventArgs e)
        {
         
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _701_f f_701 = new _701_f();
            f_701.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _702_f f_702 = new _702_f();
            f_702.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            _703_f f_703 = new _703_f();
            f_703.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            _704_f f_704 = new _704_f();
            f_704.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            _705_f f_705 = new _705_f();
            f_705.ShowDialog();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            _706_f f_706 = new _706_f();
            f_706.ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            _707_f f_707 = new _707_f();
            f_707.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            _7_f f_7 = new _7_f();
            f_7.ShowDialog();
        }



        public static void Print(DataSet MyData,string a0_text,string Form_Id,int PrintPage)
        {
            if (MyData.Tables[0].Rows.Count > 0)
            {
               /* System.Diagnostics.Process p = new System.Diagnostics.Process();
                //不现实调用程序窗口,但是对于某些应用无效
                p.StartInfo.CreateNoWindow = true;
                p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                //采用操作系统自动识别的模式
                p.StartInfo.UseShellExecute = true;
                p.StartInfo.FileName = ".\\" + main.companyName + "\\wjgl\\"+Form_Id+"\\" + a0_text + "\\" + a0_text + ".docx";
                p.StartInfo.Verb = "print";
                //开始打印
                p.Start();*/
                Microsoft.Office.Interop.Word._Application AppliApp = new Microsoft.Office.Interop.Word.Application();
                AppliApp.Visible = false;
                object filepath = System.IO.Directory.GetCurrentDirectory() + @".\" + main.companyName + @"\wjgl\" + Form_Id + @"\" + a0_text + @"\" + a0_text + ".docx";
               
                Microsoft.Office.Interop.Word._Document doc = AppliApp.Documents.Add(ref filepath);
                object missing = System.Reflection.Missing.Value;
                object Page =""+PrintPage.ToString()+"";
                doc.PrintOut(
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    Page,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing,
                    missing
                    );

            }
        }

        private void panelone_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer1_Tick_1(object sender, EventArgs e)
        {
            textBox8.Text = DateTime.Now.ToString();
        }

        private void buttonItem1_Click(object sender, EventArgs e)
        {
            panelone.Show();
            paneltwo.Hide();
            panelthree.Hide();
        }

        private void buttonItem2_Click(object sender, EventArgs e)
        {
            panelone.Hide();
            paneltwo.Show();
            panelthree.Hide();
        }

        private void buttonItem3_Click(object sender, EventArgs e)
        {
            panelthree.Show();
            panelone.Hide();
            paneltwo.Hide();
        }

        private void buttonItem4_Click(object sender, EventArgs e)
        {
            _6_f f_6 = new _6_f();
            f_6.ShowDialog();
        }

        private void buttonItem5_Click(object sender, EventArgs e)
        {
            this.Hide();
            (new Statistics_first()).ShowDialog();
            this.Show();
        }

        private void buttonItem6_Click(object sender, EventArgs e)
        {
            if (main.User_power == "管理员")
            {
                User_manage f_usernamage = new User_manage();
                f_usernamage.ShowDialog();
            }
            else
                MessageBox.Show("该用户没有权限！");
        }

        private void buttonItem7_Click(object sender, EventArgs e)
        {

            list_change_import LCI = new list_change_import();
            LCI.ShowDialog();
        }

        private void buttonItem8_Click(object sender, EventArgs e)
        {
            YNcleanALL_f f_YNcleanALL = new YNcleanALL_f();
            f_YNcleanALL.ShowDialog();
            if (YNcleanALL_f.YNdel == true)
            {
                CleanApp();
            }
        }
       
   
    }
}
