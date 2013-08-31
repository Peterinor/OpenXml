using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordToHtml
{
	public class HtmlAttribute
	{
		public string Name { get; set; }
		public string Value { get; set; }
		public HtmlAttribute(string name = null, string value = null)
		{
			this.Name = name;
			this.Value = value;
		}
		public string GetHtmlAttribute()
		{
			if (!string.IsNullOrEmpty(this.Name))
			{
				return this.Name + "=" + "\"" + (this.Value == null ? "" : this.Value) + "\"";
			}
			return "";
		}

	}
	public class HtmlElement
	{
		public HtmlElement(string tag = "html", bool close = true)
		{
			this.TargetName = tag;
			this.IsClose = close;
		}
		public bool IsClose;
		public string TargetName { get; set; }
		public string InnerHtml { get; set; }
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
			this.Attributes.Add(new HtmlAttribute("style", s.GetStyle()));
		}
		public List<HtmlElement> ChildElements = new List<HtmlElement>();
		public void AddChild(HtmlElement child)
		{
			this.ChildElements.Add(child);
			this.InnerHtml += child.GetHtml();
		}
		public string GetHtml()
		{
			var html = "";
			var attrs = "";
			foreach (HtmlAttribute attr in this.Attributes)
			{
				attrs += (" " + attr.GetHtmlAttribute());
			}
			html += "<" + this.TargetName + attrs;
			if (this.IsClose)
			{
				html += ">" + this.InnerHtml + "</" + this.TargetName + ">";
			}
			else
			{
				html += " />";
			}
			return html;
		}
	}
	public class HtmlHead : HtmlElement
	{
		public HtmlHead()
		{
			this.TargetName = "head";
		}
		public void AddMeta()
		{

		}
		public void AddStyleSheet(string src)
		{
			HtmlElement lnk = new HtmlElement("link", false);
			HtmlAttribute rel = new HtmlAttribute("rel", "stylesheet");
			HtmlAttribute type = new HtmlAttribute("type", "text/css");
			HtmlAttribute href = new HtmlAttribute("href", src);
			lnk.AddAttribute(rel);
			lnk.AddAttribute(type);
			lnk.AddAttribute(href);
			this.AddChild(lnk);
		}
		//List<CSS> css = new List<CSS>();
		public void AddStyle(CSS css)
		{
			HtmlElement c = new HtmlElement("style");
			c.InnerHtml = css.GetCSS();
			this.AddChild(c);
		}
		public void GetHtml()
		{
			base.GetHtml();
		}
		public void AddJavascript() { }
	}

	public class HtmlBody : HtmlElement
	{
		public HtmlBody()
		{
			this.TargetName = "body";
		}
	}
	public class HtmlPage
	{
		public HtmlPage()
		{
		}
		string Doctype(string ver = "5")
		{
			return "<!doctype html>";
		}

		HtmlHead head = new HtmlHead();
		HtmlBody body = new HtmlBody();

		public void AddHeadElement(HtmlElement h)
		{
			this.head.AddChild(h);
		}
		public void AddStyle(CSS c)
		{
			this.head.AddStyle(c);
		}
		public void AddElement(HtmlElement b)
		{
			this.body.AddChild(b);
		}
		public string GetHtml()
		{
			HtmlElement html = new HtmlElement("html");
			html.AddChild(this.head);
			html.AddChild(this.body);
			return this.Doctype() + html.GetHtml();
		}

	}
	class HtmlHelper
	{
	}
}
