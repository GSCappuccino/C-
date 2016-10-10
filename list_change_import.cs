using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace MyApp
{
    public partial class list_change_import : Form,FormFather
    {
        /*sjDB链接*/
        private static string ConnString = "Provider="+main.Office_Engen+";Data Source="+main.sj_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.sjDB_password+"";
        private OleDbConnection MyConn = new OleDbConnection(ConnString);
        private OleDbDataAdapter MyAd;
        private DataTable AccessTable = new DataTable();//access数据库表
        private DataTable ExcelTable = new DataTable();//excel数据表
        public list_change_import()
        {
            InitializeComponent();
            MyConn.Open();
        }

        //数据清除
        public void CleanApp()
        {
            this.Enabled = false;
            string ConnString = "Provider=" + main.Office_Engen + ";Data Source="+main.sj_path+";Persist Security Info=False;Jet OLEDB:Database Password=" +main.sjDB_password + "";
            OleDbConnection MyConn = new OleDbConnection(ConnString);
            MyConn.Open();

            string SQL = "delete from 0qfb1";
            OleDbCommand NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from 0qfb2 ";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            SQL = "delete from 0qfb3";
            NewCom = new OleDbCommand(SQL, MyConn);
            NewCom.ExecuteNonQuery();
            MyConn.Close();
            this.Enabled = true;
        }
        //__数据清除
        private void tabStrip1_SelectedTabChanged(object sender, DevComponents.DotNetBar.TabStripTabChangedEventArgs e)
        {

        }

        private void importSJ(object sender, EventArgs e)
        {
            MessageBox.Show("数据导入");
        }

        private void changeSJ(object sender, EventArgs e)
        {
            MessageBox.Show("数据修改");
        }

        private void list_change_import_Load(object sender, EventArgs e)
        {
            a4.Text = "列属性：类别号，类别说明\n=======================================================\n类别不可出现重复，否则以后面一列为准。类别请以00~99表示\n=======================================================";
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AccessTable.Clear();
            MyAd = new OleDbDataAdapter("select a0 as 编号,a1 as 名称,a2 as 所属类似群,a3 as 排序 from 0qfb3 where a0 like '%" + a5.Text + "%' or a2 like '%" + a6.Text + "%'", MyConn);
            MyAd.Fill(AccessTable);
            dataGridView1.DataSource = AccessTable;
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog2_HelpRequest(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

           if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    a2.Text = openFileDialog1.FileName;

                }
            }
            catch (Exception a)
            {
                MessageBox.Show(a.Message);
            }

        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            if (a1.Text == "类别")
                a4.Text = "列属性：类别号，类别说明\n=======================================================\n类别不可出现重复，否则以后面一列为准。类别请以00~99表示\n=======================================================";
            else if (a1.Text == "类似群")
                a4.Text = "列属性：类似群编号，说明，所属类别\n======================================================\n类似群不可出现重复，否则以后面一列为准。\n======================================================";
            else if (a1.Text == "商品服务名称")
                a4.Text = "列属性：商品服务编号，商品服务名称，所属类似群\n======================================================\n商品服务编号不可出现重复，否则以后面一列为准。\n======================================================";
            else
                a4.Text = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
          if (a1.Text=="类别")
          {
              string strConn = "Provider=" + main.Office_Engen + ";Data Source=" + a2.Text + ";Extended Properties='Excel 8.0;HDR=YES';";
              OleDbConnection ExcelConn = new OleDbConnection(strConn);
              ExcelConn.Open();
              OleDbDataAdapter ExcelAd;
              if (a3.Text == null)
                  MessageBox.Show("请输入表名！");
              else
              {
                 ExcelAd = new OleDbDataAdapter("select * from [" + a3.Text + "$]", ExcelConn);
                 ExcelAd.Fill(ExcelTable);
                 OleDbCommand DBcomm;
                 for (int a = 0; a < ExcelTable.Rows.Count; a++)
                 {
                     DBcomm = new OleDbCommand("insert into 0qfb1(a0,a1) values('" + ExcelTable.Rows[a][0].ToString() + "','" + ExcelTable.Rows[a][1].ToString() + "')", MyConn);
                     DBcomm.ExecuteNonQuery();
                 }
                 MessageBox.Show("导入成功！");
                 ExcelConn.Close();
             
              }
              
              
                 
          }
          else if(a1.Text=="类似群")
          {
              /*从excel往数据库导入*/
              string strConn = "Provider=" + main.Office_Engen + ";Data Source=" + a2.Text + ";Extended Properties='Excel 12.0;HDR=YES';";
              OleDbConnection ExcelConn = new OleDbConnection(strConn);
              ExcelConn.Open();

              OleDbDataAdapter ExcelAd = new OleDbDataAdapter("select * from [" + a3.Text + "$]", ExcelConn);
           
              ExcelAd.Fill(ExcelTable);
              OleDbCommand DBcomm;
              for (int a = 0; a < ExcelTable.Rows.Count; a++)
              {
                  DBcomm = new OleDbCommand("insert into 0qfb3(a0,a1,a2) values('" + ExcelTable.Rows[a][0].ToString() + "','" + ExcelTable.Rows[a][1].ToString() + "','"+ExcelTable.Rows[a][2].ToString()+"')", MyConn);
                  DBcomm.ExecuteNonQuery();
              }
              MessageBox.Show("导入成功！");
              ExcelConn.Close();
          }
            else if(a1.Text=="商品服务名称")
          {
              /*从excel往数据库导入*/
              string strConn = "Provider=" + main.Office_Engen + ";Data Source=" + a2.Text + ";Extended Properties='Excel 12.0;HDR=YES';";
              OleDbConnection ExcelConn = new OleDbConnection(strConn);
              ExcelConn.Open();

              OleDbDataAdapter ExcelAd = new OleDbDataAdapter("select * from [" + a3.Text + "$]", ExcelConn);
         
              ExcelAd.Fill(ExcelTable);
              OleDbCommand DBcomm;
              for (int a = 0; a < ExcelTable.Rows.Count; a++)
              {
                  DBcomm = new OleDbCommand("insert into 0qfb3(a0,a1,a2) values('" + ExcelTable.Rows[a][0].ToString() + "','" + ExcelTable.Rows[a][1].ToString() + "','"+ExcelTable.Rows[a][2].ToString()+"')", MyConn);
                  DBcomm.ExecuteNonQuery();
              }
              MessageBox.Show("导入成功！");
              ExcelConn.Close();
          }
            else
          {
              MessageBox.Show("请选择正确的导入表类型！");
          }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            YNcleanALL_f newYnDe = new YNcleanALL_f();
            newYnDe.ShowDialog();
            if (YNcleanALL_f.YNdel == true)
                CleanApp();
        }

        private void CellValueChange(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void list_change_import_FormClosing(object sender, FormClosingEventArgs e)
        {
            MyConn.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OleDbCommandBuilder objCommandBuilder = new OleDbCommandBuilder(MyAd);
            MyAd.Update(AccessTable);
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
