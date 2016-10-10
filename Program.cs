using System;
using System.Reflection;
using System.Windows.Forms;
//一定要引用
namespace MyApp
{
    class Program
    {
        [STAThread]
        
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
      
    
        static void Main()
        {
           
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
           Application.Run(new FileChose(new main()));

         

           

        }
    }
}
