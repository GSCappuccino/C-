using System;
using System.Data;
using System.Windows.Forms;

namespace MyApp
{
  
    public partial class Statistics_Find_f : Form
    {
       
        DataTable mm;
        private  DataGridViewCheckBoxColumn c1 =new DataGridViewCheckBoxColumn();//按钮列
      
        public Statistics_Find_f(DataTable nn)
        {
            
            mm = nn;
            InitializeComponent();  
          
        }
     

        private void Statistics_Find_f_Load(object sender, EventArgs e)
        {
            
          
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
         
            this.dataGridViewX1.Columns.Add(c1);
       
            c1.Selected=true;   
            c1.HeaderText = "统计公司";
            c1.Width = 60;
            dataGridViewX1.DataSource = mm;//绑定数据源
            
            dataGridViewX1.Columns[0].ReadOnly = false;
            dataGridViewX1.Columns[1].ReadOnly =true;
            dataGridViewX1.Columns[2].ReadOnly =true;
            dataGridViewX1.Columns[3].ReadOnly = true;
            dataGridViewX1.Columns[4].ReadOnly = true;
            dataGridViewX1.Columns[5].ReadOnly = true;

        }
     
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker2.Value <= dateTimePicker1.Value)
            {
                MessageBox.Show("请重新选择终止时间！");
            }
        }

        private void buttonX4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Closing(object sender, FormClosingEventArgs e)//关闭form时发生

        {
            dataGridViewX1.Columns.Clear();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 7; i++)//统计数据清零
            {
               Statistics_first.Statistics_Arry_num[i] = 0;
            }
            int cc = Convert.ToInt32(dataGridViewX1.Rows.Count.ToString());//获取行数
            if (cc==0)
            {
                MessageBox.Show("当前没有选择被统计公司！");
                return;
            }
            for (int i = 0; i < cc;i++ )
            {
                DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridViewX1.Rows[i].Cells[0];//创建单元格子（单选）
                Boolean flag = Convert.ToBoolean(checkCell.Value);//判断是否check
                if (flag = true)
                {
                    Statistics_first.Statistics_Company_name =dataGridViewX1.Rows[i].Cells[2].Value.ToString();
                   // MessageBox.Show(Statistics_first.Statistics_Company_name);
                    break;
                }
            }
         
            /*
             * 设置查询时间段
             */
            if (radio1.Checked)
            {
                if (main.currentTime.Month > 3)
                {
                    Statistics_first.Time_String[0] = main.currentTime.Year.ToString() + "-" + (main.currentTime.Month - 3).ToString() + "-" + main.currentTime.Day.ToString();
                }
                else
                {
                    Statistics_first.Time_String[0] = (main.currentTime.Year - 1).ToString() + "-" + (main.currentTime.Month + 9).ToString() + "-" + main.currentTime.Day.ToString();
                }

                Statistics_first.Time_String[1] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
                // MessageBox.Show(DateTime.Parse("2015-3-3").ToString());
            }
            else if (radio2.Checked)
            {
                if (main.currentTime.Month > 6)
                {
                    Statistics_first.Time_String[0] = main.currentTime.Year.ToString() + "-" + (main.currentTime.Month - 6).ToString() + "-" + main.currentTime.Day.ToString();
                }
                else
                {
                    Statistics_first.Time_String[0] = (main.currentTime.Year - 1).ToString() + "-" + (main.currentTime.Month + 6).ToString() + "-" + main.currentTime.Day.ToString();
                }

                Statistics_first.Time_String[1] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
            }
            else if (radio3.Checked)
            {
                Statistics_first.Time_String[0] = (main.currentTime.Year - 1).ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
                Statistics_first.Time_String[1] = main.currentTime.Year.ToString() + "-" + main.currentTime.Month.ToString() + "-" + main.currentTime.Day.ToString();
            }
            else if (dateTimePicker1.Value < dateTimePicker2.Value)
            {
                Statistics_first.Time_String[0] = dateTimePicker1.Value.Year.ToString() + "-" + dateTimePicker1.Value.Month.ToString() + "-" + dateTimePicker1.Value.Day.ToString();
                //  MessageBox.Show(Time_String[0]);
                Statistics_first.Time_String[1] = dateTimePicker2.Value.Year.ToString() + "-" + dateTimePicker2.Value.Month.ToString() + "-" + dateTimePicker2.Value.Day.ToString();
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

      

        
       

    }
}
