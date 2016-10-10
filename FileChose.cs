using Microsoft.Win32;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace MyApp
{
    public partial class FileChose : Form
    {
        object mmm;//
        
        public static bool loginYN=false;//判断是否登入成功


        
        public  FileChose(object MynForm)
        {

           ExistsRegedit();
            //main.Office_version = "Microsoft.ACE.OLEDB.12.0";
            FileStream ffs = new FileStream("m.ini", FileMode.OpenOrCreate);
            StreamReader ssr = new StreamReader(ffs);
            string aa = ssr.ReadLine();
            ssr.Close();
            ffs.Close();
            if (aa == "0")
            {
                (new login()).ShowDialog();
            }
            else
            {
                loginYN = true;
                FileStream fr = new FileStream("n.ini", FileMode.OpenOrCreate);
                StreamReader sw = new StreamReader(fr);
                main.UserId = sw.ReadLine();
                main.User_password = sw.ReadLine();
                main.User_power = sw.ReadLine();
                sw.Close();
                fr.Close();
            }

            if (!loginYN)
            {
                System.Environment.Exit(0);
            }

            else
            {
                InitializeComponent();
                mmm = MynForm;
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        { 
           
        }

        private void bindingNavigatorCountItem_Click(object sender, EventArgs e)
        {

        }

        private void buttonItem6_Click(object sender, EventArgs e)
        {

        }

     

        private void tabControlPanel2_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

       
        private void labelX1_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_Click(object sender, EventArgs e)
        {

        }

       

        private void buttonX3_Click(object sender, EventArgs e)//查询按钮
        {
            
          
        }

        private void dataGridViewX1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
         
        }

        private void tabControlPanel1_Click(object sender, EventArgs e)
        {

        }

      
       

        private void buttonX2_Click(object sender, EventArgs e)
        {
            if (radio2.Checked)
            {
                if (textBoxX3.Text == "")
                {
                    MessageBox.Show("不可输入空公司目录名！\n请重新输入！");
                    return;
                }

                try
                {
                    main.companyName = this.textBoxX3.Text;
                    string ConnString = "Provider=" + main.Office_Engen + ";Data Source=.\\" + main.companyName + "\\appDb.mdb;Persist Security Info=False;Jet OLEDB:Database Password=" + main.appDB_password + "";
                    OleDbConnection MyConn = new OleDbConnection(ConnString);
                    MyConn.Open();

                    //连接数据库 ，，表数据



                    this.Hide();

                    ((Form)mmm).ShowDialog();//显示main窗体

                    this.Close();

                }
                catch(Exception x)
                {
                    MessageBox.Show("该客户未创建！");
                }
                
               
                   
            }
        }

        private void FileChose_Load_1(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            textBoxX3.Enabled = false;
         //   textBoxX3.Text = "DB";///调试
        }

        private void buttonX5_Click(object sender, EventArgs e)
        {
            string ConnString = "Provider="+main.Office_Engen+";Data Source="+main.User_path+";Persist Security Info=False;Jet OLEDB:Database Password="+main.UserDB_password+"";
            OleDbConnection MyConn = new OleDbConnection(ConnString);
            MyConn.Open();
            OleDbDataAdapter Myad = new OleDbDataAdapter("select * from company_list where company_name like '%" + maskedTextBoxAdv1.Text + "%'", MyConn);//模糊查询
            DataSet mydata = new DataSet();
            Myad.Fill(mydata);
           

            if (mydata.Tables[0].Rows.Count == 0)
                MessageBox.Show("该客户资料未在本系统中！");
            else
            {
              
                (new FindManager(mydata.Tables[0])).ShowDialog();
            }
        }

        private void textBoxX3_TextChanged(object sender, EventArgs e)
        {

        }

        private void radio1_CheckedChanged(object sender, EventArgs e)
        {
            if (radio1.Checked)
            {
                this.Hide();
                (new NewManager(mmm)).ShowDialog();
                this.Show();
            }
        }

        private void radio2_CheckedChanged(object sender, EventArgs e)
        {
            if (radio2.Checked)
            textBoxX3.Enabled = true;
            else
                textBoxX3.Enabled = false;
        }
        public void ExistsRegedit()
        {
            try  
            {  
                RegistryKey rk = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Excel.exe");  
                  
                if (rk != null)  
                {  
                    string path1 = rk.GetValue("Path").ToString();  
                    string path = path1.Substring(0, path1.Length - 1);  
                    string version = path.Substring(path.LastIndexOf("\\") + 1);   
                   
                    switch (version)  
                    {  
                        case "Office11"://检查本机是否安装Office2003  
                            //label1.Text = "您的Office版本是2003，Version11.0。";  
                           // label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.Jet.OLEDB.4.0";
                            main.Office_Version_Id = "11.0";
                            break;  
                        case "OFFICE11"://检查本机是否安装Office2003  
                            //label1.Text = "您的Office版本是2003，Version11.0。";  
                           // label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.Jet.OLEDB.4.0";
                            main.Office_Version_Id = "11.0";
                            break;  
                        case "Office12"://检查本机是否安装Office2007  
                           // label1.Text = "您的Office版本是2007，Version12.0。";  
                            //label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.ACE.OLEDB.12.0";
                            main.Office_Version_Id = "12.0";
                            break;  
                        case "OFFICE12"://检查本机是否安装Office2007  
                           // label1.Text = "您的Office版本是2007，Version12.0。";  
                           // label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.ACE.OLEDB.12.0";
                            main.Office_Version_Id = "12.0";
                            break;  
                        case "Office14"://检查本机是否安装Office2010  
                           // label1.Text = "您的Office版本是2010，Version14.0。";  
                           // label2.Text = "文件路径:" + path1; 
                            main.Office_Engen="Microsoft.ACE.OLEDB.12.0";
                            main.Office_Version_Id = "12.0";
                            break;  
                        case "OFFICE14"://检查本机是否安装Office2010  
                           // label1.Text = "您的Office版本是2010，Version14.0。";  
                            //label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.ACE.OLEDB.12.0";
                            main.Office_Version_Id = "12.0";
                            break;  
                        case "Office15"://检查本机是否安装Office2013  
                           // label1.Text = "您的Office版本是2013，Version15.0。";  
                           // label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.ACE.OLEDB.15.0";
                            main.Office_Version_Id = "15.0";
                            break;  
                        case "OFFICE15"://检查本机是否安装Office2013  
                            //label1.Text = "您的Office版本是2013，Version15.0。";  
                           // label2.Text = "文件路径:" + path1;  
                            main.Office_Engen="Microsoft.ACE.OLEDB.15.0";
                            main.Office_Version_Id = "15.0";
                            break;  
                    }  
                }  
                else  
                {  
                    
                     MessageBox.Show("您未安装office 2003 2007 2010 2013中的任一版本，请安装其一后运行！");
                     System.Environment.Exit(0);
                }  
            }  
            catch (Exception ex)  
            {  
                MessageBox.Show(ex.Message);  
            }  
         
        }
      
        
    }
}
