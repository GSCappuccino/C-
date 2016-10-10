using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MyApp
{
    public partial class list_second : Form
    {
        object motherForm;//基调用者   第一个表
        Type motherType;
        public static string text;

        public static string a;


      
        
        public list_second(object aa,Type bb)
        { 
           
            motherForm=aa;//赋值
            motherType = bb;
            InitializeComponent();
           
            
        }

        private void list_second_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
         richTextBox1.Text = uniquelistone.text;

        
         OleDbDataAdapter MyAd = new OleDbDataAdapter("select a0,a1 from 0qfb2 where a2 = '"+uniquelistone.a+"'",uniquelistone.MyConn);//"
         DataSet MyData = new DataSet();

      

         MyAd.Fill(MyData);
         
            this.dataGridView1.DataSource=MyData.Tables[0];
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
          
            OleDbDataAdapter MyAd = new OleDbDataAdapter("select a0,a1 from 0qfb2 where a2 = '" + uniquelistone.a + "'and a0 like '%" + textBox1.Text + "' and a1 like '%" + textBox2.Text + "%'",uniquelistone.MyConn);//群号查询有问题
            DataSet MyData = new DataSet();//

            // MessageBox.Show(uniquelistone.a);

            MyAd.Fill(MyData);

            this.dataGridView1.Refresh();

            this.dataGridView1.DataSource = MyData.Tables[0];
        }

        private void button2_Click(object sender, EventArgs e)
        {

           
            OleDbDataAdapter MyAd = new OleDbDataAdapter("select a0,a1 from 0qfb2 where a2 = '" + uniquelistone.a + "'",uniquelistone.MyConn);//群号查询有问题
            DataSet MyData = new DataSet();//

            // MessageBox.Show(uniquelistone.a);

            MyAd.Fill(MyData);

            this.dataGridView1.Refresh();

            this.dataGridView1.DataSource = MyData.Tables[0];

            this.textBox1.Text = null;
            this.textBox2.Text = null;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            //text 文本

            text = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString() + this.dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            //数据源的更改,传递点击的数据
            a = this.dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            this.Hide();
           (new list_third(motherForm,motherType)).ShowDialog();
           if (uniquelistone.message == 3)
               this.Close();
            this.Show();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
