using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

using System.Text.RegularExpressions;

namespace WordToHtml
{
	public class ElementHandler
	{
		WordprocessingDocument Document;
		MainDocumentPart MainPart;
		public ElementHandler(WordprocessingDocument doc)
		{
			this.Document = doc;
			this.MainPart = doc.MainDocumentPart;
		}
		public HtmlElement Handle(DocumentFormat.OpenXml.OpenXmlElement elem)
		{
			HtmlElement no = new HtmlElement("no");
			if ((new Regex("Paragraph")).Match(elem.GetType().ToString()).Success)
			{
				return this.HandleParagraph((Paragraph)elem);
			}
			else if ((new Regex("Table")).Match(elem.GetType().ToString()).Success)
			{
				return this.HandleTable((Table)elem);
			}
			else
			{
				return no;
			}
		}

		public HtmlElement HandleParagraph(Paragraph p)
		{
			HtmlElement _pa = new HtmlElement("p");

			OpenXMLHelper helper = new OpenXMLHelper();

			StyleConfig config = new StyleConfig();
			config.Charset = "EastAsia";

			#region Paragraph
			var p_class = "";
			if (p.ParagraphProperties != null)
			{
				var pPr = p.ParagraphProperties;
				_pa.AddStyle(helper.GetStyleFromParagraphProperties(pPr, config));

				if (pPr.ParagraphStyleId != null)
				{
					var id = pPr.ParagraphStyleId.Val;
					p_class += "Inner_P_" + id;
				}
			}
			_pa.AddAttribute("class", p_class);
			//div.AddChild(_pa);
			var runs = p.Elements<Run>();
			foreach (var r in runs)
			{
				var span = new HtmlElement("span");
				var rpr = r.RunProperties;

				config.Charset = "EastAsia";
				var tt = r.InnerText;
				Regex reg = new Regex("[\x00-\xff]{1,}", RegexOptions.IgnoreCase | RegexOptions.Multiline);
				var mc = reg.Match(tt);
				if (mc.Success)
				{
					config.Charset = "Ascii";
				}
				span.AddStyle(helper.GetStyleFromRunProperties(rpr, config));

				if (r.RunProperties != null && r.RunProperties.RunStyle != null && r.RunProperties.RunStyle.Val != null)
				{
					var id = r.RunProperties.RunStyle.Val;
					span.AddAttribute("class", "Inner_P_" + id + " " + "Inner_R_" + id);
				}

				var ts = r.Elements<Text>();
				var pics = r.Elements<Drawing>();
				foreach (Text t in ts)
				{
					var _s = new HtmlElement("span");
					_s.InnerHtml = t.InnerText;
					span.AddChild(_s);
				}

				foreach (Drawing d in pics)
				{
					DocumentFormat.OpenXml.Drawing.GraphicData gd = d.Inline.Graphic.GraphicData;
					var pic = gd.Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ElementAt(0);
					var pb = pic.BlipFill;
					var blip = pb.Blip;
					var ppp = blip.Embed.Value;

					//var path = doc.GetReferenceRelationship(ppp).Uri;
					//var relationships = doc.Package.GetRelationships();
					//content += ppp;
					var imgp = this.MainPart.GetPartById(ppp);

					var img = new HtmlElement("img");
					img.AddAttribute("src", imgp.Uri.ToString().Replace("/word/", "./"));
					span.AddChild(img);
				}
				_pa.AddChild(span);

			}
			#endregion
			return _pa;
		}

		public HtmlElement HandleTable(Table t)
		{
			HtmlElement table = new HtmlElement("table");

			var tblPrs = t.Elements<TableProperties>();


			var trs = t.Elements<TableRow>();
			foreach (TableRow tr in trs)
			{
				HtmlElement h_tr = new HtmlElement("tr");
				var tcs = tr.Elements<TableCell>();
				foreach (TableCell tc in tcs)
				{
					HtmlElement h_td = new HtmlElement("td");
					var pars = tc.Elements<Paragraph>();
					foreach (Paragraph p in pars)
					{
						h_td.AddChild(this.HandleParagraph(p));
					}
					h_tr.AddChild(h_td);
				}
				table.AddChild(h_tr);
			}
			return table;
		}
	}
}
