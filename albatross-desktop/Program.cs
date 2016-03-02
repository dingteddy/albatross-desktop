using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace albatross_desktop
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new DiffMergeForm(args));
            return;
            //Application.Run(new StartForm());
            /*args = new string[2];
            args[0] = @"E:\ThreeKingdoms\doc\静态数据表\type_copys_main.xlsx";
            args[1] = @"E:\ThreeKingdoms\doc\静态数据表\type_copys_daily.xlsx";*/
            if (args.Length == 0)
            {
                Application.Run(new StartForm());
            }
            else
            {
                Application.Run(new DiffMergeForm(args));
            }
        }
    }
}
