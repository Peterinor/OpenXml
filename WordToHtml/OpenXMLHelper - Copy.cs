using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2010;
using DocumentFormat.OpenXml.Packaging;

namespace WordToHtml
{
	public class StyleConfig
	{
		public string Charset { get; set; }
	}
	public class StyleBudle
	{

	}
	public class CssStyle
	{
		public string Selector;
		public Dictionary<string, string> Styles = new Dictionary<string, string>();
		public CssStyle(string selector = null)
		{
			this.Selector = selector;
		}
		public void AddStyle(string name, string value)
		{
			this.Styles.Add(name, value);
		}
		public string GetStyle()
		{
			string _style = "";
			foreach (string css in this.Styles.Keys)
			{
				_style += css + ":" + this.Styles[css] + ";";
			}
			if (!string.IsNullOrEmpty(this.Selector))
			{
				_style = this.Selector + "{" + _style + "}";
			}
			return _style;
		}
		public void CombineStyles(CssStyle style)
		{
			foreach (string k in style.Styles.Keys)
			{
				if (this.Styles.Keys.Contains(k))
				{
					//this.Styles[k] += style.Styles[k];
					this.Styles[k] = style.Styles[k];
				}
				else
				{
					this.Styles.Add(k, style.Styles[k]);
				}
			}
		}
	}
	public class CSS
	{
		public List<CssStyle> styles = new List<CssStyle>();
		public void AddStyle(CssStyle s)
		{
			if (!string.IsNullOrEmpty(s.Selector))
			{
				this.styles.Add(s);
			}
			//Error
		}
		public string GetCSS()
		{
			var _css = "";
			for (var i = 0; i < this.styles.Count; i++)
			{
				_css += this.styles[i].GetStyle();
			}
			return _css;
		}
	}
	public class OpenXMLHelper
	{
		public CSS GetStyles(IEnumerable<Style> styles)
		{
			CSS css = new CSS();
			//生产Style模版对应的CSS Style
			foreach (Style s in styles)
			{
				var id = s.StyleId;
				var type = s.Type.ToString();
				Style linked_s = null;
				if (s.LinkedStyle != null)
				{
					var linked_id = s.LinkedStyle.Val;
					var s_1 = from ls in styles
							  where ls.StyleId.Value == linked_id
							  select ls;
					var k = s_1.Count();
					if (k > 0)
					{
						linked_s = s_1.ElementAt(0);
					}
				}
				css.AddStyle(this.GetStyle(s, linked_s));
			}
			return css;
		}
		public CssStyle GetStyle(Style s, Style linkedStyle)
		{
			CssStyle cs = this.GetStyle(s);
			cs.CombineStyles(this.GetStyle(linkedStyle));
			return cs;
		}

		public CssStyle GetStyle(Style s, StyleConfig config = null)
		{
			CssStyle cs = new CssStyle();
			if (s == null)
			{
				return cs;
			}
			var id = s.StyleId;
			var type = s.Type.ToString();
			if (type == "paragraph")
			{
				if (s.StyleParagraphProperties != null)
				{
					CssStyle c_s = this.GetStyleFromParagraphProperties(new ParagraphProperties(s.StyleParagraphProperties.OuterXml), config);
					c_s.Selector = ".Inner_P_" + id;
					if (s.StyleRunProperties != null)
					{
						CssStyle c_ss = this.GetStyleFromRunProperties(new RunProperties(s.StyleRunProperties.OuterXml), config);
						c_s.CombineStyles(c_ss);
					}
					return c_s;
				}
			}
			else if (type == "character")
			{
				if (s.StyleRunProperties != null)
				{
					CssStyle c_s = this.GetStyleFromRunProperties(new RunProperties(s.StyleRunProperties.OuterXml), config);
					c_s.Selector = ".Inner_R_" + id;
					return c_s;
				}
			}
			return cs;
		}

		public CssStyle GetStyleFromRunProperties(RunProperties r, StyleConfig config = null)
		{
			CssStyle _style = new CssStyle(null);
			var rPr = r;
			if (rPr != null)
			{
				if (rPr.RunFonts != null)
				{
					if (config.Charset == "Ascii")
					{
						if (rPr.RunFonts.Ascii != null || rPr.RunFonts.HighAnsi != null)
						{
							var fs = (rPr.RunFonts.Ascii != null ? rPr.RunFonts.Ascii.Value : "nofont")
							+ "," + (rPr.RunFonts.HighAnsi != null ? rPr.RunFonts.HighAnsi.Value : "nofont");
							_style.AddStyle("font-family", fs);
						}
					}
					if (config.Charset == "EastAsia")
					{
						if (rPr.RunFonts.EastAsia != null)
						{
							var fs = (rPr.RunFonts.EastAsia != null ? rPr.RunFonts.EastAsia.Value : "nofont");
							_style.AddStyle("font-family", fs);
						}
					}
				}
				if (rPr.FontSize != null && rPr.FontSize.Val != null)
				{
					_style.AddStyle("font-size", rPr.FontSize.Val + "px");
				}
				if (rPr.Bold != null)
				{
					if (rPr.Bold.Val != null && rPr.Bold.Val.Value)
					{
						_style.AddStyle("font-weight", "bold");
					}
				}
				if (rPr.Color != null)
				{
					_style.AddStyle("color", "#" + rPr.Color.Val);
				}
			}

			return _style;
		}

		public CssStyle GetStyleFromParagraphProperties(ParagraphProperties p, StyleConfig config = null)
		{
			CssStyle _style = new CssStyle(null);
			var pPr = p;
			if (pPr != null)
			{
				if (pPr.Justification != null)
				{
					var jc = pPr.Justification;
					_style.AddStyle("text-align", jc.Val);
				}
				//pPr.ParagraphMarkRunProperties;

				if (pPr.ParagraphMarkRunProperties != null)
				{
					RunProperties rPr = new RunProperties(pPr.ParagraphMarkRunProperties.OuterXml);
					CssStyle p_s = this.GetStyleFromRunProperties(rPr);
					_style.CombineStyles(p_s);
				}

			}
			return _style;
		}
	}
}
