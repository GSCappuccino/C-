using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace MyApp
{
    unsafe public partial class Each_f_statistics : Form
    {
        private string key;
        private string value;
        private DataTable MyData;
        Dictionary<string, string> Item_list;
        private Dictionary<string, string>.Enumerator Enumera;
        private int* i;//保存form i值的地址
        private OleDbConnection MyConn;
        private String DB_table_name;
        public Each_f_statistics()
        {
            InitializeComponent();
        }

        unsafe public Each_f_statistics(string DB_table_name,OleDbConnection MyConn,int* i,DataTable MyData,Dictionary<string,string> Item_list)
        {
            // TODO: Complete member initialization
            InitializeComponent();       
            this.Item_list = Item_list;
            this.MyData = MyData;
            this.MyConn = MyConn;
            this.i = i;
            this.DB_table_name = DB_table_name;
            
            Enumera =Item_list.GetEnumerator();
            for (int a = 0; a <Item_list.Count; a++)
            {
                if (Enumera.MoveNext())
                {                   
                     value =Enumera.Current.Value;
                }
                comboBox1.Items.Insert(a, value);
            }
          
           
        }

        private void Each_f_statistics_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        unsafe private void button1_Click(object sender, EventArgs e)
        {
            if (Each_statistics(DB_table_name, MyConn, MyData, Item_list, comboBox1.Text, textBox1.Text) >= 0) //返回查询后的数值i 
                *i = Each_statistics(DB_table_name, MyConn, MyData, Item_list, comboBox1.Text, textBox1.Text);
            else if (Each_statistics(DB_table_name, MyConn, MyData, Item_list, comboBox1.Text, textBox1.Text) == -2)
                return;
            this.Close();
        }
        private int Each_statistics(string DB_table_name, OleDbConnection MyConn, DataTable MyData, Dictionary<string, string> Item_list, string Item_text, string Key_text)//获取 i
        {
            string key = null;
            Dictionary<string, string>.Enumerator Enumera = Item_list.GetEnumerator();
            for (int a = 0; a < Item_list.Count; a++)
            {

                if (Enumera.MoveNext())
                {

                    if (Enumera.Current.Value == Item_text)
                    {
                        key = Enumera.Current.Key;
                        break;
                    }

                }


            }
            if (key == null)
            {
                MessageBox.Show("请输入正确的查询条目！");
                return -2;
            }
            //模糊查询  从数据库得到数据
            OleDbDataAdapter MyAd = new OleDbDataAdapter("select " + key + " from " + DB_table_name + " where " + key + " like '%" + Key_text + "%'", MyConn);
            DataTable my = new DataTable();
            MyAd.Fill(my);
            if (my.Rows.Count <= 0)
            {
                MessageBox.Show("对不起没有查到你所需要的数据！");
                return -1;

            }
            else
            {
                //将模糊查询得到的结果反馈到form的datatable中去   返回相应的I值
                for (int j = 0; j < MyData.Rows.Count; j++)
                {
                    if (my.Rows[0][0].ToString() == MyData.Rows[j][key].ToString())
                    {
                        return j;
                    }

                }
                return -1;
            }

        }
       
    }
}
