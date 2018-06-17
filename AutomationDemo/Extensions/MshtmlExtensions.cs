using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHDocVw;

namespace Extensions
{
    public static class MshtmlExtensions
    {
        public static mshtml.HTMLDocument GetDocumentByAddress(this SHDocVw.ShellWindowsClass shell, string address)
        {
            var wbSearched = new ShellWindowsClass().Cast<IWebBrowser2>().FirstOrDefault(x => x.LocationURL.ToLower().Contains(address.ToLower()));
            if (wbSearched == null)
                return null;
            if (!(wbSearched.Document is mshtml.HTMLDocument))
                return null;

            return (mshtml.HTMLDocument)wbSearched.Document;
        }

        public static mshtml.HTMLDocument GetDocumentByPtr(this SHDocVw.ShellWindowsClass shell, IntPtr pointer)
        {
            var wbSearched = new ShellWindowsClass().Cast<IWebBrowser2>().FirstOrDefault(x => x.HWND == (int)pointer);
            if (wbSearched == null)
            {
                return null;
            }
            if (!(wbSearched.Document is mshtml.HTMLDocument))
                return null;

            return (mshtml.HTMLDocument)wbSearched.Document;
        }

        public static bool InjectJavascript(this mshtml.HTMLDocument document, string script)
        {
            try
            {
                document.parentWindow.execScript(script, "Javascript");
                return true;
            }
            catch (Exception e)
            {
                return false;
            }

        }
    }
}

