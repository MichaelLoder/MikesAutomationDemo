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
using AutomationDemo.BO;
using AutomationDemo.Demos;
using Extensions;
using mshtml;

namespace AutomationDemo.Views
{
    public partial class MainView : Form
    {
        private mshtml.HTMLDocument _doc;
        private Dictionary<int, IHTMLElement> _pageElements;
        private Element _selectedElement = null;
        private string currentColor = "";
        public delegate void DHTMLEvent(IHTMLEventObj e);
        public MainView()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Add("locating IE with google web page ");
            _doc = HookingIE.AttachToIEFromWebAddressExample();
            _doc.InjectJavascript("alert('TEST')");
            button3.Enabled = true;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.Items.Add("locating IE with google web page ");
            if (_doc == null)
            {
                _doc = HookingIE.AttachToIEFromWebAddressExample();
            }
            else
            {
                listBox1.Items.Add("Document already attached, using corrent document ");
            }
           
          
            if (_doc == null)
            {
                MessageBox.Show("Unable to locate window with google chrome");
                return;
            }

            listBox1.Items.Add("IE found attaching load events");
            var windowEvts = (_doc.parentWindow as HTMLWindowEvents2_Event);
            windowEvts.onload += WindowEvtsOnOnload;

            button3.Enabled = true;

            // DOMEventHandler onmousedownhandler = new DOMEventHandler(_doc);
            //  onmousedownhandler.Handler += new DOMEvent(Mouse_Down);

        }

        private void WindowEvtsOnOnload(IHTMLEventObj pEvtObj)
        {
            Invoke(new MethodInvoker(() =>
            {
                listBox1.Items.Add("Window loaded : Domain - " + _doc.domain);
            }));

        }

        public void Mouse_Down(mshtml.IHTMLEventObj e)
        {
            MessageBox.Show(e.srcElement.tagName);
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            _pageElements = new Dictionary<int, IHTMLElement>();
            var elements = GetElements().Select(x => new Element(x)
            {

            });

           listBox2.Items.AddRange(elements.ToArray());

        }

        private IEnumerable<IHTMLElement> GetElements()
        {
            foreach (var htmlElement in _doc.all.Cast<IHTMLElement>().ToList())
            {
                yield return htmlElement;
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            if (listBox2.SelectedItem != null)
            {
                if (this._selectedElement != null)
                {
                    if (string.IsNullOrEmpty(this.currentColor ))
                    {
                        this._selectedElement.HtmlElement.style.backgroundColor = "transparent";
                    }
                    else
                    {
                        this._selectedElement.HtmlElement.style.backgroundColor = currentColor;
                    }
                  
                }
                var element = (Element) listBox2.SelectedItem;
                this.currentColor = element.HtmlElement.style.backgroundColor?.ToString();
                element.HtmlElement.style.backgroundColor = "yellow";
                this._selectedElement = element;
            }
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
