using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AutomationDemo.Demos;
using Extensions;
using mshtml;

namespace AutomationDemo.Views
{
    public partial class MainView : Form
    {
        private mshtml.HTMLDocument _doc;
        public delegate void DHTMLEvent(IHTMLEventObj e);
        public MainView()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var ie = HookingIE.AttachToIEFromWebAddressExample();
            ie.InjectJavascript("alert('TEST')");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Items.Add("locating IE with google web page ");
            _doc = HookingIE.AttachToIEFromWebAddressExample();
          
            if (_doc == null)
            {
                MessageBox.Show("Unable to locate window with google chrome");
                return;
            }

            listBox1.Items.Add("IE found attaching load events");
            var windowEvts = (_doc.parentWindow as HTMLWindowEvents2_Event);
            windowEvts.onload += WindowEvtsOnOnload;

           // DOMEventHandler onmousedownhandler = new DOMEventHandler(_doc);
          //  onmousedownhandler.Handler += new DOMEvent(Mouse_Down);
           
        }

        private void WindowEvtsOnOnload(IHTMLEventObj pEvtObj)
        {
            Invoke(new MethodInvoker(() =>
            {
                listBox1.Items.Add("Window loaded");
            }));

        }

        public void Mouse_Down(mshtml.IHTMLEventObj e)
        {
            MessageBox.Show(e.srcElement.tagName);
        }
    }


        public delegate void DOMEvent(mshtml.IHTMLEventObj e);
        public class DOMEventHandler
        {
        public DOMEvent Handler;
        DispHTMLDocument Document;
        public DOMEventHandler(DispHTMLDocument doc)
        {
            this.Document = doc;
        }
        [DispId(0)]
        public void Call()
        {
            Handler(Document.parentWindow.@event);
        }
        }
    
}
