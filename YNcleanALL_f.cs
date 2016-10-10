using System;
using System.Windows.Forms;

namespace MyApp
{
    public partial class YNcleanALL_f : Form
    {
       public static Boolean YNdel=false;
        public YNcleanALL_f()
        {
          
            InitializeComponent();
        }

        private void YNcleanALL_f_Load(object sender, EventArgs e)
        {
           
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);

        }

        private void button2_Click(object sender, EventArgs e)
        {
            YNdel = false;
            this.Close();
        }

     
        private void button1_Click(object sender, EventArgs e)
        {


            YNdel = true;
            this.Close();
        }

        private void label1_Click(object sender, EventArgs e)
        {
           
        }
    }
}
