using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MyApp
{
    public partial class uniquelistone : Form
    {
        public static string  ConnString = "Provider="+main.Office_Engen+";Data Source="+main.sj_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.sjDB_password+"";
        public static OleDbConnection    MyConn = new OleDbConnection(ConnString); 
          

        object motherForm;//基调用者
        Type motherType;
        public static int message;//两个 按钮  区分至  以及  填充完后关闭窗体
        public static string cli;
        //数据库链接
       
        OleDbDataAdapter MyAd;//第一个的数据源
      
      

        DataSet MyData;
       public static string a;
       public  static string text;
        public uniquelistone(object aa,Type bb)
        {
            
            MyConn.Open();
            motherType = bb;
            motherForm=aa;//把基调用者传进去
            InitializeComponent();
            //数据库链接，读取数据形成表
           
           // MessageBox.Show(cli.Length.ToString());
            if (cli.Length == 1)
                cli = '0' + cli;
            MyAd = new OleDbDataAdapter("select * from 0qfb1 where a0 like '%" + cli + "'", MyConn);
             MyData = new DataSet();
            
            MyAd.Fill(MyData);
        }

        private void uniquelistone_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            //绑定数据源
          
            dataGridView1.DataSource=MyData.Tables[0];
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
          /*  
           * 测试鼠标获取的值。。
           * 
           * String ID;
            if (e.RowIndex < 0)
            {
                ID = "caonima ";
            }
            else
            {
              ID =  this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();//获得本行第一个单元格的数据，以此类推
            }
            MessageBox.Show(ID);
           
           */

           //双击后datagirdview改变
            
           //text 文本

              text = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            //数据源的更改,传递点击的数据
              a = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
              this.Hide();
              (new list_second(motherForm,motherType)).ShowDialog();
              if (message == 3)
                  this.Close();
              this.Show();


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            //搜索
         
            OleDbDataAdapter MyAd = new OleDbDataAdapter("select * from 0qfb1 where a1 like '%" + this.textBox1.Text  + "%'", MyConn);//"
            DataSet MyData = new DataSet();

          

            MyAd.Fill(MyData);
            this.dataGridView1.Refresh();

            this.dataGridView1.DataSource=MyData.Tables[0];

           
        }

        private void CloseDB(object sender, FormClosingEventArgs e)
        {
            MyConn.Close();
        }
    }
}
