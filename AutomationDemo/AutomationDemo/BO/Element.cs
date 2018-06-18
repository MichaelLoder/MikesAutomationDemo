using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using mshtml;

namespace AutomationDemo.BO
{
    public class Element
    {
        public IHTMLElement HtmlElement { get; set; }
        public int Id { get; set; }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(HtmlElement.id))
            {
                return "Element with Id : " + HtmlElement.id;
            }
            else
            {
                return "Element with tag : " +  HtmlElement.tagName;
            }
        }

        public Element(IHTMLElement element)
        {
            this.HtmlElement = element;
        }
    }
}
