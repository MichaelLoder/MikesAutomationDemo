using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Extensions;
using mshtml;

namespace AutomationDemo.Demos
{
    public static class HookingIE
    {
        //make sure IE is running and has the google page loaded
        public static HTMLDocument AttachToIEFromWebAddressExample()
        {
            // get ie instance from the windows api, and use the address to locate the page
            var ie = new SHDocVw.ShellWindowsClass().GetDocumentByAddress("google");
            // inject a simple alert
            return ie;
        }


        public static void AttachToIEUsingMainNativeWindowIntPtr()
        {
            // find all process by IE and select the MainWindowHandle
            var ieProcess = Process.GetProcessesByName("iexplore").Select(x => x.MainWindowHandle);
            HTMLDocument _doc;
            // loop through all main window handles ( selecting all processes will select sub processes)
            foreach (var intPtr in ieProcess)
            {
                // try to get the IE Instance
                _doc = new SHDocVw.ShellWindowsClass().GetDocumentByPtr(intPtr);
                // if you 
                if (_doc != null)
                {
                    CallScript(_doc, "AttachToIEUsingMainNativeWindowIntPtr");
                    break;
                }
            }

        }

        private static void CallScript(HTMLDocument doc, string message)
        {
            doc.InjectJavascript("alert('" + message +"')");
        }

        
        public static void IEEvent()
        {
            var _doc = new SHDocVw.ShellWindowsClass().GetDocumentByAddress("google");

            var docEvents = (_doc as HTMLDocumentEvents2_Event);
            if (docEvents == null)
            {
                Console.WriteLine("Doc events is null");
                return;
            }
            
            docEvents.onreadystatechange += DocEventsOnOnreadystatechange;
        }

        private static void DocEventsOnOnreadystatechange(IHTMLEventObj pEvtObj)
        {
            Console.WriteLine("ready state changed ");
        }

        
    }
}
