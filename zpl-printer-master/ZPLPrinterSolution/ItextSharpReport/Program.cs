using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
//using System;
//using System.Collections.Generic;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;
using System.Runtime.InteropServices;
using System.Media;
using System.Drawing;

namespace ItextSharpReport
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }


    }

}
