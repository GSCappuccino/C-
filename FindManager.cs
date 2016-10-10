using System;
using System.Data;
using System.Windows.Forms;

namespace MyApp
{
    public partial class FindManager : Form
    {
     
        public FindManager(DataTable a)
        {
        
            
            InitializeComponent();
            dataGridView1.DataSource = a;
        }

        private void FindManager_Load(object sender, EventArgs e)
        {
            
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
        }
    }
}
