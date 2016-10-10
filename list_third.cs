using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;

namespace MyApp
{
    public partial class list_third : Form
    {
        public static DataTable finalltable = new DataTable();

        object motherForm;
        Type motherType;
        public list_third(object aa,Type bb)
        {
            motherForm = aa;
            motherType = bb;
            InitializeComponent();
        }

        private void list_third_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            //初始化最后的表
            finalltable.Reset();
            finalltable.Columns.Add("序号");
            finalltable.Columns.Add("类别");
            finalltable.Columns.Add("类似群");
            finalltable.Columns.Add("商品编号");
            finalltable.Columns.Add("商品名称");
         

            //显示上方数据
            richTextBox1.Text=uniquelistone.text+"\n\n"+list_second.text;
            //checkbox的建立
            DataGridViewCheckBoxColumn newColumn = new DataGridViewCheckBoxColumn();
            newColumn.HeaderText = "选择";
            dataGridView1.Columns.Add(newColumn);


           
            OleDbDataAdapter MyAd = new OleDbDataAdapter("select a0,a1 from 0qfb3 where a2 = '" + list_second.a + "'",uniquelistone.MyConn);//"
            DataSet MyData = new DataSet();



            MyAd.Fill(MyData);

            this.dataGridView1.DataSource = MyData.Tables[0];

            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
         
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
             
            finalltable.Clear();
           
         
            //将选中的列表写入到box中去
            string one = null;
            int cc = Convert.ToInt32(dataGridView1.Rows.Count.ToString());//获取行数
            int z=1;
            for (int i = 0; i < cc; i++)//循环遍历表行 获取选中的数据
            {

                DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[0];
                Boolean flag = Convert.ToBoolean(checkCell.Value);


                if (flag == true)     //查找被选择的数据行
                {
                    //从 DATAGRIDVIEW 中获取数据项
                    string z_text = dataGridView1.Rows[i].Cells[2].Value.ToString().Trim();
                    string z_num = dataGridView1.Rows[i].Cells[1].Value.ToString().Trim();

                    one = one + dataGridView1.Rows[i].Cells[1].Value.ToString().Trim() + dataGridView1.Rows[i].Cells[2].Value.ToString().Trim();
                    // MessageBo      x.Show(z_zcode);
                  
                    finalltable.Rows.Add(new object[]{z.ToString(),uniquelistone.a,list_second.a,dataGridView1.Rows[i].Cells[1].Value.ToString().Trim(), dataGridView1.Rows[i].Cells[2].Value.ToString().Trim()});
                    z++;

                }


            }
          

            this.Hide();
         //  
           ( new list_fourth(motherForm,motherType)).ShowDialog();
           if (uniquelistone.message == 3)
           {
               finalltable.Reset();

               this.Close();
               
           }
            this.Show();

            

           
        }
    }
}
