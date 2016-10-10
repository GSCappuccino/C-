using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace MyApp
{
    public partial class changepassword : Form
    {
        public changepassword()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FileStream fss = new FileStream("n.ini", FileMode.OpenOrCreate);       
            StreamReader srr = new StreamReader(fss);
        
            //数据库连接
            string ConnString = "Provider="+main.Office_Engen+";Data Source="+main.sj_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.sjDB_password+"";
            OleDbConnection MyConn = new OleDbConnection(ConnString);
            MyConn.Open();
            string StWords = "select * from zzcc where a1='"+srr.ReadLine()+"'";//where a1='" + textBox1.Text + "'
            OleDbDataAdapter MyAd = new OleDbDataAdapter(StWords, MyConn);
            DataSet MyDataSet = new DataSet();
            MyAd.Fill(MyDataSet);

            fss.Close();
            srr.Close();

            if (textBox1.Text != main.User_password)
            {
                MessageBox.Show("密码错误！！！");
                textBox1.Text = null;

            }
            else
            {
                if (textBox2.Text == "" || textBox3.Text == "")
                    MessageBox.Show("请不要更改成空密码！！");
                else if (textBox2.Text != textBox3.Text)
                {
                    MessageBox.Show("两次密码输入不同！！");
                    textBox2.Text = null;
                    textBox3.Text = null;
                }
                else//密码更改
                {

                    MyDataSet.Tables[0].Rows[0]["a2"] = textBox2.Text;
                    OleDbCommandBuilder MyComm = new OleDbCommandBuilder(MyAd);
                    MyAd.Update(MyDataSet);
                
                    FileStream fs = new FileStream("n.ini", FileMode.OpenOrCreate);
                    fs.SetLength(0);
                    StreamWriter sr = new StreamWriter(fs);
                    sr.WriteLine(MyDataSet.Tables[0].Rows[0]["a1"]);
                    sr.WriteLine(MyDataSet.Tables[0].Rows[0]["a2"]);
                    sr.WriteLine(MyDataSet.Tables[0].Rows[0]["a3"]);
                    sr.Close();
                    fs.Close();
              
                    MessageBox.Show("修改成功！！");
                    this.Close();

                }
            }
        }

        private void changepassword_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }
    }
}
