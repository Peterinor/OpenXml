using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToHtml.Html;

namespace WordToHtml
{
    public class OpenXmlHelper
	{
		public CSS GetStyles(IEnumerable<Style> styles)
		{
			CSS css = new CSS();
			//生产Style模版对应的CSS Style
			foreach (Style s in styles)
			{
				var id = s.StyleId;
				var type = s.Type.ToString();
				Style linkedS = null;
				if (s.LinkedStyle != null)
				{
					var linkedId = s.LinkedStyle.Val;
					var s1 = from ls in styles
							  where ls.StyleId.Value == linkedId
							  select ls;
					var k = s1.Count();
					if (k > 0)
					{
						linkedS = s1.ElementAt(0);
					}
				}
				css.AddStyle(this.GetStyle(s, linkedS));
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
					CssStyle cS = this.GetStyleFromParagraphProperties(new ParagraphProperties(s.StyleParagraphProperties.OuterXml), config);
					cS.Selector = ".Inner_P_" + id;
					if (s.StyleRunProperties != null)
					{
						CssStyle cSs = this.GetStyleFromRunProperties(new RunProperties(s.StyleRunProperties.OuterXml), config);
						cS.CombineStyles(cSs);
					}
					return cS;
				}
			}
			else if (type == "character")
			{
				if (s.StyleRunProperties != null)
				{
					CssStyle cS = this.GetStyleFromRunProperties(new RunProperties(s.StyleRunProperties.OuterXml), config);
					cS.Selector = ".Inner_R_" + id;
					return cS;
				}
			}
			else if (type == "table")
			{
				if (s.StyleTableProperties != null)
				{
					CssStyle cS = this.GetStyleFromTableProperties(new TableProperties(s.StyleTableProperties.OuterXml), config);
					cS.Selector = ".Inner_T_" + id;
					return cS;
				}
			}
			else if (type == "numbering")
			{
			}
			else if (type == "direct")
			{

			}
			return cs;
		}

		public CssStyle GetStyleFromRunProperties(RunProperties r, StyleConfig config = null)
		{
			CssStyle style = new CssStyle(null);
			var rPr = r;
			if (rPr != null)
			{
				if (rPr.RunFonts != null)
				{
					if (config != null)
					{
						if (config.Charset == "Ascii")
						{
							if (rPr.RunFonts.Ascii != null || rPr.RunFonts.HighAnsi != null)
							{
								var fs = (rPr.RunFonts.Ascii != null ? rPr.RunFonts.Ascii.Value : "nofont")
								+ "," + (rPr.RunFonts.HighAnsi != null ? rPr.RunFonts.HighAnsi.Value : "nofont");
								style.AddStyle("font-family", fs);
							}
						}
						if (config.Charset == "EastAsia")
						{
							if (rPr.RunFonts.EastAsia != null)
							{
								var fs = (rPr.RunFonts.EastAsia != null ? rPr.RunFonts.EastAsia.Value : "nofont");
								style.AddStyle("font-family", fs);
							}
						}
					}
					else
					{
						var fs = (rPr.RunFonts.Ascii != null ? rPr.RunFonts.Ascii.Value : "nofont")
						+ "," + (rPr.RunFonts.HighAnsi != null ? rPr.RunFonts.HighAnsi.Value : "nofont")
						+ "," + (rPr.RunFonts.EastAsia != null ? rPr.RunFonts.EastAsia.Value : "nofont");
						style.AddStyle("font-family", fs);
					}
				}
				if (rPr.FontSize != null && rPr.FontSize.Val != null)
				{
					style.AddStyle("font-size", rPr.FontSize.Val + "px");
				}
				if (rPr.Bold != null)
				{
					if (rPr.Bold.Val != null && rPr.Bold.Val.Value)
					{
						style.AddStyle("font-weight", "bold");
					}
				}
				if (rPr.Color != null)
				{
					style.AddStyle("color", "#" + rPr.Color.Val);
				}
			}

			return style;
		}

		public CssStyle GetStyleFromParagraphProperties(ParagraphProperties p, StyleConfig config = null)
		{
			CssStyle style = new CssStyle(null);
			var pPr = p;
			if (pPr != null)
			{
				if (pPr.Justification != null)
				{
					var jc = pPr.Justification;
					style.AddStyle("text-align", jc.Val);
				}
				if (pPr.ParagraphMarkRunProperties != null)
				{
					RunProperties rPr = new RunProperties(pPr.ParagraphMarkRunProperties.OuterXml);
					CssStyle p_s = this.GetStyleFromRunProperties(rPr, config);
					style.CombineStyles(p_s);
				}

			}
			return style;
		}

		public CssStyle GetStyleFromTableProperties(TableProperties r, StyleConfig config = null)
		{
			CssStyle _style = new CssStyle(null);

			return _style;
		}
	}
}
