using System;
using System.Windows.Forms;

namespace MyApp
{
    public partial class list_fourth : Form
    {
        object motherForm;
        Type motherType;
        public list_fourth(object aa,Type bb)
        {
            motherForm=aa;
            motherType = bb;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            this.Close();
        }

        private void list_fourth_Load(object sender, EventArgs e)
        {
            this.Left = (int)((Screen.PrimaryScreen.Bounds.Width - this.Width) / 2);
            this.Top = (int)((Screen.PrimaryScreen.Bounds.Height - this.Height) / 2);
            dataGridView2.DataSource = list_third.finalltable;

       
 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            string one=null;
            for(int i=0;i<list_third.finalltable.Rows.Count;i++)
            {
               one =one + (i + 1).ToString() +"."+ list_third.finalltable.Rows[i][4]+"； ";
            }
            if (uniquelistone.message == 1)
            {
            //  MessageBox.Show(MethodBase.GetCurrentMethod().ReflectedType);
                
             ((TextBox) main.text_19).Text= list_third.finalltable.Rows[0][1].ToString();

             
               ((RichTextBox) main.text_20).Text = one;
                list_third.finalltable.Clear();
                uniquelistone.message = 3;
                this.Close();

            }
            if (uniquelistone.message == 2)
            {
                ((TextBox) main.text_21).Text = list_third.finalltable.Rows[0][1].ToString();
                ((RichTextBox) main.text_22).Text = one;
                
                list_third.finalltable.Clear();
                uniquelistone.message = 3;
            
                this.Close();

            }         

        }
    }
}
