using System;
using System.Data;
using System.Windows.Forms;
using System.Threading;
using System.Data.OleDb;
using DevComponents.DotNetBar.SuperGrid;
using BarChart;
using System.Drawing;
namespace MyApp
{
    
    public partial class Statistics_Chart : Form
    {


        private HBarChart barChart;
        public Statistics_Chart()
        {

           
            InitializeComponent();
            // Create, no need if you added the chart by visual editor
            barChart = new HBarChart();
            this.panel1.Controls.Add(barChart);
            barChart.Dock = DockStyle.Fill;
            barChart.Description.Text = "业务统计";
            
        }

        private void Statistics_Chart_Load(object sender, EventArgs e)
        {

            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            
        
      
            /*
             * supergrid制作
             */
            //统计数据
            GridRow[] nn = new GridRow[7];
            Thread one = new Thread(new ThreadStart(fun_1));
            one.Start();
            Thread two = new Thread(new ThreadStart(fun_2));
            two.Start();
            Thread four = new Thread(new ThreadStart(fun_4));
            four.Start();
            Thread five = new Thread(new ThreadStart(fun_5));
            five.Start();
            Thread six = new Thread(new ThreadStart(fun_6));
            six.Start();
            Thread serven = new Thread(new ThreadStart(fun_3));
            serven.Start();
            Thread three_1 = new Thread(new ThreadStart(fun_3_1));
            three_1.Start();
            Thread three_2 = new Thread(new ThreadStart(fun_3_2));
            three_2.Start();
            Thread three_3 = new Thread(new ThreadStart(fun_3_3));
            three_3.Start();


            fun_7();//线程同步
            //画表
            superGridControl1.PrimaryGrid.ReadOnly = true;
            nn[0] = new GridRow("1.注册申请类", Statistics_first.Statistics_Arry_num[0], Statistics_first.Statistics_Arry_num[0] * 2); 
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[0]);
            nn[1] = new GridRow("2.异议", Statistics_first.Statistics_Arry_num[1], Statistics_first.Statistics_Arry_num[1] * 2); 
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[1]);
            nn[2] = new GridRow("3.续展变更转让许可质押注销", Statistics_first.Statistics_Arry_num[2], Statistics_first.Statistics_Arry_num[2] * 2);
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[2]);    
            nn[3] = new GridRow("4.特殊标志申请", Statistics_first.Statistics_Arry_num[3], Statistics_first.Statistics_Arry_num[3] * 2);
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[3]);   
            nn[4] = new GridRow("5.补证撤三出具证明", Statistics_first.Statistics_Arry_num[4], Statistics_first.Statistics_Arry_num[4] * 2);
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[4]);
            nn[5] = new GridRow("6.商标代理委托书", Statistics_first.Statistics_Arry_num[5], Statistics_first.Statistics_Arry_num[5] * 2);
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[5]); 
            nn[6] = new GridRow("7.评审复审", Statistics_first.Statistics_Arry_num[6], Statistics_first.Statistics_Arry_num[6] * 2);
            superGridControl1.PrimaryGrid.Rows.Add((GridElement)nn[6]);
            /*
             *chart图制作
             * 
             */
        
            barChart.Add(Statistics_first.Statistics_Arry_num[0],"1.注册申请类", Color.FromArgb(255, 200, 255, 255));
            barChart.Add(Statistics_first.Statistics_Arry_num[1], "2.异议", Color.FromArgb(255, 150, 200, 255));
            barChart.Add(Statistics_first.Statistics_Arry_num[2], "3.续展变更转让许可质押注销", Color.FromArgb(255, 150, 200, 255));
            barChart.Add(Statistics_first.Statistics_Arry_num[3], "4.特殊标志申请", Color.FromArgb(255, 150, 200, 255));
            barChart.Add(Statistics_first.Statistics_Arry_num[4], "5.补证撤三出具证明", Color.FromArgb(255, 150, 200, 255));
            barChart.Add(Statistics_first.Statistics_Arry_num[5], "6.商标代理委托书", Color.FromArgb(255, 150, 200, 255));
            barChart.Add(Statistics_first.Statistics_Arry_num[6], "7.评审复审", Color.FromArgb(255, 150, 200, 255));
        }

        private void fun_3_3()
        {
             string ConnString = "Provider="+main.Office_Engen+";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password="+main.appDB_password+"";
          OleDbConnection  MyConn = new OleDbConnection(ConnString);

            OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from e11sbxz where time_e>=#"+Statistics_first.Time_String[0]+"# and time_e<=#"+Statistics_first.Time_String[1]+"#", MyConn);



            OleDbDataAdapter MyAd02 = new OleDbDataAdapter("select count(a0) from e12chxz where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd03 = new OleDbDataAdapter("select count(a0) from e21bgsb where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd04 = new OleDbDataAdapter("select count(a0) from e22chsb where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd05 = new OleDbDataAdapter("select count(a0) from e23bgdl where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);


            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);
            DataSet MyData02 = new DataSet();
            MyAd02.Fill(MyData02);
            DataSet MyData03 = new DataSet();
            MyAd03.Fill(MyData03);
            DataSet MyData04 = new DataSet();
            MyAd04.Fill(MyData04);
            DataSet MyData05 = new DataSet();
            MyAd05.Fill(MyData05);

            Statistics_first.Statistics_Arry_num[2]=Statistics_first.Statistics_Arry_num[2]+
                
                int.Parse(MyData01.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData02.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData03.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData04.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData05.Tables[0].Rows[0][0].ToString()) ;
          
        }

        private void fun_3_2()
        {
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
      OleDbConnection  MyConn = new OleDbConnection(ConnString);

      OleDbDataAdapter MyAd011 = new OleDbDataAdapter("select count(a0) from e51xkba  where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
      OleDbDataAdapter MyAd012 = new OleDbDataAdapter("select count(a0) from e52bgmc where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
      OleDbDataAdapter MyAd013 = new OleDbDataAdapter("select count(a0) from e53tjzz where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
      OleDbDataAdapter MyAd014 = new OleDbDataAdapter("select count(a0) from e54chxk where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
      OleDbDataAdapter MyAd015 = new OleDbDataAdapter("select count(a0) from e61zrzy where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            DataSet MyData011 = new DataSet();
            MyAd011.Fill(MyData011);
            DataSet MyData012 = new DataSet();
            MyAd012.Fill(MyData012);
            DataSet MyData013 = new DataSet();
            MyAd013.Fill(MyData013);
            DataSet MyData014 = new DataSet();
            MyAd014.Fill(MyData014);
            DataSet MyData015 = new DataSet();
            MyAd015.Fill(MyData015);
            Statistics_first.Statistics_Arry_num[2] =
              Statistics_first.Statistics_Arry_num[2] +
               int.Parse(MyData011.Tables[0].Rows[0][0].ToString()) +
               int.Parse(MyData012.Tables[0].Rows[0][0].ToString()) +
               int.Parse(MyData013.Tables[0].Rows[0][0].ToString()) +
               int.Parse(MyData014.Tables[0].Rows[0][0].ToString()) +
               int.Parse(MyData015.Tables[0].Rows[0][0].ToString());
        
        }

        private void fun_3_1()
        {

            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
         OleDbConnection  MyConn = new OleDbConnection(ConnString);
         OleDbDataAdapter MyAd06 = new OleDbDataAdapter("select count(a0) from e24chdl where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
         OleDbDataAdapter MyAd07 = new OleDbDataAdapter("select count(a0) from e31sjsp where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
         OleDbDataAdapter MyAd08 = new OleDbDataAdapter("select count(a0) from e32chsj where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
         OleDbDataAdapter MyAd09 = new OleDbDataAdapter("select count(a0) from e41zryz where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<= #" + Statistics_first.Time_String[1] + "#", MyConn);
         OleDbDataAdapter MyAd010 = new OleDbDataAdapter("select count(a0) from e42chrz where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);

            DataSet MyData06 = new DataSet();
            MyAd06.Fill(MyData06);
            DataSet MyData07 = new DataSet();
            MyAd07.Fill(MyData07);
            DataSet MyData08 = new DataSet();
            MyAd08.Fill(MyData08);
            DataSet MyData09 = new DataSet();
            MyAd09.Fill(MyData09);
            DataSet MyData010 = new DataSet();
            MyAd010.Fill(MyData010);

            Statistics_first.Statistics_Arry_num[2]=
                Statistics_first.Statistics_Arry_num[2]+       
                int.Parse(MyData06.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData07.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData08.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData09.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData010.Tables[0].Rows[0][0].ToString())  ;
        
        }
        private void fun_7()
        {

            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
       OleDbConnection  MyConn = new OleDbConnection(ConnString);

       OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from g1bhfs where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
       OleDbDataAdapter MyAd02 = new OleDbDataAdapter("select count(a0) from g2byfs where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
       OleDbDataAdapter MyAd03 = new OleDbDataAdapter("select count(a0) from g3cxfs where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
       OleDbDataAdapter MyAd04 = new OleDbDataAdapter("select count(a0) from g4wxfs where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
       OleDbDataAdapter MyAd05 = new OleDbDataAdapter("select count(a0) from g5wxsq where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "# ", MyConn);
       OleDbDataAdapter MyAd06 = new OleDbDataAdapter("select count(a0) from g6yyfs where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
       OleDbDataAdapter MyAd07 = new OleDbDataAdapter("select count(a0) from g7chps where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
       OleDbDataAdapter MyAd08 = new OleDbDataAdapter("select count(a0) from g8fswt where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
         








            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);
            DataSet MyData02 = new DataSet();
            MyAd02.Fill(MyData02);
            DataSet MyData03 = new DataSet();
            MyAd03.Fill(MyData03);
            DataSet MyData04 = new DataSet();
            MyAd04.Fill(MyData04);
            DataSet MyData05 = new DataSet();
            MyAd05.Fill(MyData05);
            DataSet MyData06 = new DataSet();
            MyAd06.Fill(MyData06);
            DataSet MyData07 = new DataSet();
            MyAd07.Fill(MyData07);
            DataSet MyData08 = new DataSet();
            MyAd08.Fill(MyData08);
 
            Statistics_first.Statistics_Arry_num[6] =
                int.Parse(MyData01.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData02.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData03.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData04.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData05.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData06.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData07.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData08.Tables[0].Rows[0][0].ToString());
         //   MessageBox.Show(Statistics_first.Statistics_Arry_num[6].ToString());
         
        }

        private void fun_6()
        {
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);


            OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from f81wts where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);


            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);



            Statistics_first.Statistics_Arry_num[5] = int.Parse(MyData01.Tables[0].Rows[0][0].ToString());
          
            
        }

        private void fun_5()
        {
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);


            OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from c1bfsq where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd02 = new OleDbDataAdapter("select count(a0) from c2bfzc where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd03 = new OleDbDataAdapter("select count(a0) from c3cjzc where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd04 = new OleDbDataAdapter("select count(a0) from c4gzzc where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd05 = new OleDbDataAdapter("select count(a0) from c5snby  where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd06 = new OleDbDataAdapter("select count(a0) from c6cxsn where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd07 = new OleDbDataAdapter("select count(a0) from c7cxsp where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd08 = new OleDbDataAdapter("select count(a0) from c8chcx where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);


            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);
            DataSet MyData02 = new DataSet();
            MyAd02.Fill(MyData02);
            DataSet MyData03 = new DataSet();
            MyAd03.Fill(MyData03);
            DataSet MyData04 = new DataSet();
            MyAd04.Fill(MyData04);
            DataSet MyData05 = new DataSet();
            MyAd05.Fill(MyData05);
            DataSet MyData06 = new DataSet();
            MyAd06.Fill(MyData06);
            DataSet MyData07 = new DataSet();
            MyAd07.Fill(MyData07);
            DataSet MyData08 = new DataSet();
            MyAd08.Fill(MyData08);
            Statistics_first.Statistics_Arry_num[4] =
                int.Parse(MyData01.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData02.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData03.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData04.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData05.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData06.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData07.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData08.Tables[0].Rows[0][0].ToString());
          
        }

        private void fun_4()
        {
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);

            OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from dtsbz where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
         

            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);
          


            Statistics_first.Statistics_Arry_num[3] = int.Parse(MyData01.Tables[0].Rows[0][0].ToString());
         
            
        }

        private void fun_2()
        {

            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);

            OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from bsbyy1 where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd02 = new OleDbDataAdapter("select count(a0) from bsbyy2 where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);

            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);
            DataSet MyData02 = new DataSet();
            MyAd02.Fill(MyData02);


            Statistics_first.Statistics_Arry_num[1] = 
                int.Parse(MyData01.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData02.Tables[0].Rows[0][0].ToString());
           
         
        }

        private void fun_3()
        {


            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);

            OleDbDataAdapter MyAd016 = new OleDbDataAdapter("select count(a0) from e62zqbg where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd017 = new OleDbDataAdapter("select count(a0) from e63djyq where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd018 = new OleDbDataAdapter("select count(a0) from e64bfsq where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd019 = new OleDbDataAdapter("select count(a0) from e65zyzx where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd020 = new OleDbDataAdapter("select count(a0) from e71sbzx where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd021 = new OleDbDataAdapter("select count(a0) from e72cxzx where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);

          
           
           
            DataSet MyData016 = new DataSet();
            MyAd016.Fill(MyData016);
            DataSet MyData017 = new DataSet();
            MyAd017.Fill(MyData017);
            DataSet MyData018 = new DataSet();
            MyAd018.Fill(MyData018);
            DataSet MyData019 = new DataSet();
            MyAd019.Fill(MyData019);
            DataSet MyData020 = new DataSet();
            MyAd020.Fill(MyData020);
            DataSet MyData021 = new DataSet();
            MyAd021.Fill(MyData021);
            Statistics_first.Statistics_Arry_num[2] = 
               Statistics_first.Statistics_Arry_num[2]+
                
                int.Parse(MyData016.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData017.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData018.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData019.Tables[0].Rows[0][0].ToString()) + 
                int.Parse(MyData020.Tables[0].Rows[0][0].ToString()) +
                int.Parse(MyData021.Tables[0].Rows[0][0].ToString());
          
        }

        private void fun_1()
        {
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + Statistics_first.Statistics_Company_name + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);

            OleDbDataAdapter MyAd01 = new OleDbDataAdapter("select count(a0) from 1sbzcsq where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd02 = new OleDbDataAdapter("select count(a0) from 1sbzcsq2 where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd03 = new OleDbDataAdapter("select count(a0) from achsq3 where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            OleDbDataAdapter MyAd04 = new OleDbDataAdapter("select count(a0) from ayxzm4 where time_e>=#" + Statistics_first.Time_String[0] + "# and time_e<=#" + Statistics_first.Time_String[1] + "#", MyConn);
            /*1sbzcsq2
achsq3
ayxzm4*/
            DataSet MyData01 = new DataSet();
            MyAd01.Fill(MyData01);
            DataSet MyData02 = new DataSet();
            MyAd02.Fill(MyData02);
            DataSet MyData03 = new DataSet();
            MyAd03.Fill(MyData03);
            DataSet MyData04 = new DataSet();
            MyAd04.Fill(MyData04);

            Statistics_first.Statistics_Arry_num[0] = int.Parse(MyData01.Tables[0].Rows[0][0].ToString()) + int.Parse(MyData02.Tables[0].Rows[0][0].ToString()) + int.Parse(MyData03.Tables[0].Rows[0][0].ToString()) + int.Parse(MyData04.Tables[0].Rows[0][0].ToString());
         
          
        }

    }
}
