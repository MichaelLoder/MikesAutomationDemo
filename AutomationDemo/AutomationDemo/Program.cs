using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutomationDemo.Demos;
using AutomationDemo.Views;

namespace AutomationDemo
{
    class Program
    {
        // make thread STA
        [STAThread]
        static void Main(string[] args)
        {
           // HookingIE.AttachToIEFromWebAddressExample();
            //HookingIE.AttachToIEUsingMainNativeWindowIntPtr();
            Application.Run(new ApplicationContext(new MainView()));
          
        }
    }
}
