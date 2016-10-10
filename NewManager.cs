using System;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace MyApp
{
    public partial class NewManager : Form
    {
        object mmm;
        public NewManager(object mm)
        {
            mmm = mm;
            InitializeComponent();
        }

        private void NewManager_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {


            try
            {
                string ConnString = "Provider=" + main.Office_Engen + ";Data Source=" + main.User_path + ";Persist Security Info=False;Jet OLEDB:Database Password=" + main.UserDB_password + "";
                OleDbConnection MyConn = new OleDbConnection(ConnString);
                MyConn.Open();
                string SQL = "insert into company_list(company_name,company_people,company_tel,create_time) values('" + this.textBoxX1.Text + "','" + this.textBoxX2.Text + "','" + this.textBoxX3.Text + "','" + DateTime.Now.ToShortDateString() + "')";
                OleDbCommand MyCom = new OleDbCommand(SQL, MyConn);
                MyCom.ExecuteNonQuery();
                Directory.CreateDirectory("" + this.textBoxX1.Text + "");
                File.Copy("" + main.appDB_path + "", "" + this.textBoxX1.Text + "\\appDb.mdb");


                main.companyName = this.textBoxX1.Text;

                this.Hide();

                ((Form)mmm).ShowDialog();//显示main窗体

                this.Close();
            }
            catch(Exception x)
            {
                MessageBox.Show(x.Message);
            }
               
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            textBoxX1.Text = null;
            textBoxX2.Text = null;
            textBoxX3.Text = null;
        }
    }
}
