using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace R2Tech_Valida_Arquivos
{
    static class Program
    {

        public static string valorDoParamentro = "";
        
        /// <summary>
        /// Ponto de entrada principal para o aplicativo.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {

            for (int i = 0; i < args.Count(); i++)
            {
                valorDoParamentro = args[i];
            }

            //MessageBox.Show(valorDoParamentro);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
