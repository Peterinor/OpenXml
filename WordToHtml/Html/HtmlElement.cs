using System.Collections.Generic;
using System.Text;

namespace WordToHtml.Html
{
    public class HtmlElement
    {
        public HtmlElement(string tag = "html", bool needClose = true)
        {
            this.TargetName = tag;
            this.IsNeedClose = needClose;
        }

        public HtmlElement(string tag, string innerHtml)
            : this(tag, true)
        {
            this._innerHtml = innerHtml;
        }

        public bool IsNeedClose;
        public string TargetName { get; set; }
        private readonly string _innerHtml;
        public StringBuilder InnerHtml
        {
            get
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(this._innerHtml);
                foreach (var childElement in this.ChildElements)
                {
                    sb.Append(childElement);
                }
                return sb;
            }
        }

        public List<HtmlAttribute> Attributes = new List<HtmlAttribute>();
        public void AddAttribute(HtmlAttribute att)
        {
            this.Attributes.Add(att);
        }
        public void AddAttribute(string name, string value)
        {
            this.Attributes.Add(new HtmlAttribute(name, value));
        }
        public void AddStyle(CssStyle s)
        {
            this.Attributes.Add(new HtmlAttribute("style", s.ToString()));
        }

        public List<HtmlElement> ChildElements = new List<HtmlElement>();
        public void AddChild(HtmlElement child)
        {
            this.ChildElements.Add(child);
        }
        public override string ToString()
        {
            StringBuilder attrs = new StringBuilder();
            StringBuilder html = new StringBuilder();
            foreach (HtmlAttribute attr in this.Attributes)
            {
                attrs.Append(" " + attr);
            }
            html.Append("<" + this.TargetName + attrs);
            if (this.IsNeedClose)
            {
                html.Append(">")
                    .Append(this.InnerHtml)
                    .Append("</").Append(this.TargetName).Append(">");
            }
            else
            {
                html.Append(" />");
            }
            return html.ToString();
        }
    }
}