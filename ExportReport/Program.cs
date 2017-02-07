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
            //Launcher launcher = new Launcher();
            //launcher.showView(true);
            try
            {
                Type type = Type.GetTypeFromProgID("ScalesData.Launcher");
                if (type != null)
                {
                    dynamic launcher = Activator.CreateInstance(type);
                    launcher.showView(true);
                } else
                {
                    MessageBox.Show("Couldn't load library file. System exit.");
                }
            } catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                MessageBox.Show(e.ToString());
            }
        }
    }
}
