using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace MyApp
{
    public partial class login : Form
    {

      
        public login()
        {

           
            InitializeComponent();
        }
       ~login()
        {
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string ConnString = "Provider="+main.Office_Engen+";Data Source="+main.sj_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.sjDB_password+"";
            OleDbConnection MyConn = new OleDbConnection(ConnString);
            MyConn.Open();
            string StWords="select * from zzcc where a1='"+textBox1.Text+"'";
            OleDbDataAdapter MyAd = new OleDbDataAdapter(StWords,MyConn);
            DataSet MyDataSet = new DataSet();
            MyAd.Fill(MyDataSet);
            string aa=null;
            if (MyDataSet.Tables[0].Rows.Count > 0)
            {
                aa = MyDataSet.Tables[0].Rows[0]["a3"].ToString();
               
                if (textBox2.Text != MyDataSet.Tables[0].Rows[0]["a2"].ToString())
                { 
                    MessageBox.Show("密码或账号输入错误！！！！");
                    FileChose.loginYN = false;
                }
                else
                {
                    FileChose.loginYN = true;
                    main.UserId = textBox1.Text;
                    main.User_password=MyDataSet.Tables[0].Rows[0][2].ToString();
                    main.User_power = MyDataSet.Tables[0].Rows[0][3].ToString();
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("密码或账号输入错误！！！！");
                FileChose.loginYN = false;
            }


            if (checkBox1.Checked == true)
            {
                FileStream fs = new FileStream("n.ini", FileMode.OpenOrCreate);
                fs.SetLength(0);
                StreamWriter sr = new StreamWriter(fs);
                sr.WriteLine(this.textBox1.Text);
                sr.WriteLine(this.textBox2.Text);
                sr.Close();
                fs.Close();
            }
            else
            {
                FileStream fs = new FileStream("n.ini", FileMode.OpenOrCreate);
                fs.SetLength(0);
                fs.Close();
            }
            if (checkBox2.Checked == true)
            {

                FileStream fs = new FileStream("n.ini", FileMode.OpenOrCreate);
                fs.SetLength(0);
                StreamWriter sr = new StreamWriter(fs);
                sr.WriteLine(this.textBox1.Text);
                sr.WriteLine(this.textBox2.Text);
                sr.WriteLine(aa);
                sr.Close();
                fs.Close();

                FileStream ffs = new FileStream("m.ini", FileMode.OpenOrCreate);
                ffs.SetLength(0);
                StreamWriter ssr = new StreamWriter(ffs);
                ssr.WriteLine("1");
                ssr.Close();
                ffs.Close();
            }
            else
            { 
                FileStream ffs = new FileStream("m.ini", FileMode.OpenOrCreate);
                ffs.SetLength(0);
                StreamWriter ssr = new StreamWriter(ffs);
                ssr.WriteLine("0");
                ssr.Close();
                ffs.Close();
            }

        }

        private void login_Load(object sender, EventArgs e)
        {
          
            checkBox2.CheckState =CheckState.Unchecked;
            checkBox1.CheckState = CheckState.Unchecked;
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            FileStream fr = new FileStream("n.ini", FileMode.OpenOrCreate);
            StreamReader sw = new StreamReader(fr);
            this.textBox1.Text = sw.ReadLine();
            this.textBox2.Text = sw.ReadLine();
            sw.Close();
            fr.Close(); 
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.CheckState == CheckState.Checked)
                checkBox1.CheckState = CheckState.Checked;
        }
    }
}
