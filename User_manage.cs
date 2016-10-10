using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace MyApp
{
    public partial class User_manage : Form
    {
        private string ConnString = "Provider="+main.Office_Engen+";Data Source="+main.sj_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.sjDB_password+"";
        private OleDbConnection MyConn;
        private OleDbDataAdapter MyAd;
        private  DataTable MyData;
        public User_manage()
        {
            InitializeComponent();

            MyConn = new OleDbConnection(ConnString);
            MyConn.Open();
            MyAd = new OleDbDataAdapter("select a0 as 编号,a1 as 用户名,a2 as 密码,a3 as 权限  from zzcc", MyConn);
            MyData = new DataTable();
            MyAd.Fill(MyData);


            dataGridView1.DataSource = MyData;  
          
      
        }

        private void User_manage_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);

        }

        private void button17_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
       /*     DataTable MyData0 = new DataTable();
            OleDbDataAdapter MyAd0 = new OleDbDataAdapter("select *  from zzcc where a0 like '%"+textBox1.Text+"%' or a3 like '%"+comboBox1.Text+"%'", MyConn);
            MyAd0.Fill(MyData0);
            dataGridView1.DataSource = MyData0;*/
            
        }

        private void tabStrip1_SelectedTabChanged(object sender, DevComponents.DotNetBar.TabStripTabChangedEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void click_fun(object sender, EventArgs e)
        {
            
        }

          
         

        private void button8_Click(object sender, EventArgs e)//baocun
        {
            OleDbCommandBuilder objCommandBuilder = new OleDbCommandBuilder(MyAd);         
            MyAd.Update(MyData);       
            MessageBox.Show("修改成功！");
       
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly = false;//将当前单元格设为可读
            this.dataGridView1.CurrentCell = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];//获取当前单元格
            this.dataGridView1.BeginEdit(true);//将单元格设为编辑状态
        }

     
    }
}
