using ScalesData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportReport
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Launcher launcher = new Launcher();
            launcher.showView(true);
        }
    }
}
