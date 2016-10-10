using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace MyApp
{
    public partial class Statistics_first : Form
    {
        public static string[] Time_String = new string[2]{null,null};//查询时间段
        public static string Statistics_Company_name=main.companyName;
        public static int[] Statistics_Arry_num=new int[7]{0,0,0,0,0,0,0};
        public Statistics_first()
        {
            InitializeComponent();
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker2.Value <= dateTimePicker1.Value)
            {
                MessageBox.Show("请重新选择终止时间！");
            }
        }

        private void Statistics_f_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
           
        }

        

       

        private void buttonX4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       
        private void buttonX2_Click(object sender, EventArgs e)
        {
            string ConnString = "Provider="+main.Office_Engen+";Data Source="+main.User_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.UserDB_password+"";
            OleDbConnection MyConn = new OleDbConnection(ConnString);
            OleDbDataAdapter Myad = new OleDbDataAdapter("select * from company_list where company_name like '%" +textBox1.Text + "%'", MyConn);//模糊查询
            DataSet mydata = new DataSet();
            Myad.Fill(mydata);
            

            this.Hide();
            (new Statistics_Find_f(mydata.Tables[0])).ShowDialog();
            this.Show();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
           
                for (int i = 0; i < 7; i++)//统计数据清零
                {
                    Statistics_Arry_num[i] = 0;
                }
                if (radio2.Checked)
                {
                    Statistics_Company_name = textBoxX1.Text;
                }//设置查询公司名字-----默认链接公司名

                /*
                 * 设置查询时间段
                 */
                if (radio3.Checked)
                {


                    Time_String[0] = main.currentTime.AddMonths(-3).ToString("yyyy-MM-dd");
                    main.currentTime = DateTime.Now;
                    Time_String[1] = main.currentTime.ToString("yyyy-MM-dd");
                    MessageBox.Show(Time_String[0]+Time_String[1]);
                 
                }
                else if (radio4.Checked)
                {

                   
                    Time_String[0] = main.currentTime.AddMonths(-6).ToString("yyyy-MM-dd");
                    main.currentTime = DateTime.Now;
           
                    Time_String[1] = main.currentTime.ToString("yyyy-MM-dd");
                   
                }
                else if (radio5.Checked)
                {
                    Time_String[0] = (main.currentTime.Year - 1).ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
                 
                    Time_String[1] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
                   
                }
                else if (dateTimePicker1.Value < dateTimePicker2.Value)
                {
                    Time_String[0] = dateTimePicker1.Value.Year.ToString() + "-" + dateTimePicker1.Value.Month.ToString() + "-" + dateTimePicker1.Value.Day.ToString();
                   
                    Time_String[1] = dateTimePicker2.Value.Year.ToString() + "-" + dateTimePicker2.Value.Month.ToString() + "-" + dateTimePicker2.Value.Day.ToString();
                    
                }
                else
                {
                    MessageBox.Show("时间选择错误！");
                    return;
                }
             
          
   
            this.Hide();
            (new Statistics_Chart()).ShowDialog();
          
            this.Show();
        }

        private void radio2_CheckedChanged(object sender, EventArgs e)
        {
            if (radio1.Checked)
                textBoxX1.Enabled = false;
            else
                textBoxX1.Enabled = true;
        }

      
      
    }
}
