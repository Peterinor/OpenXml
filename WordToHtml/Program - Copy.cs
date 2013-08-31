using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml;
using System.Xml.Linq;
using System.IO;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2010;
using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Drawing;

namespace WordToHtml
{
	class Program
	{
		static void Main(string[] args)
		{
			HtmlElement div = new HtmlElement("div");
			HtmlAttribute attr = new HtmlAttribute("style", "color:white");
			div.AddAttribute(attr);

			HtmlElement p = new HtmlElement("p");
			p.InnerHtml = "This is a Paragraph!";

			div.AddChild(p);
			//Console.WriteLine(div.GetHtml());

			HtmlHead head = new HtmlHead();
			head.AddStyleSheet("test.css");
			//head.AddStyle(
			//Console.WriteLine(head.GetHtml());

			HtmlPage page = new HtmlPage();
			//page.AddHeadElement(
			//page
			page.AddElement(div);
			page.AddElement(p);
			Console.WriteLine(page.GetHtml());

			xMain(args);
		}

		static void xMain(string[] args)
		{
			string file = "./word.docx";
			//file = "./DirectX_9_3D.docx";

			OpenXMLHelper helper = new OpenXMLHelper();

			HtmlPage _html = new HtmlPage();

			HtmlElement meta1 = new HtmlElement("meta", false);
			HtmlAttribute a1 = new HtmlAttribute("http-equiv", "content-type");
			HtmlAttribute a2 = new HtmlAttribute("content", "text/html; charset=UTF-8");
			meta1.AddAttribute(a1);
			meta1.AddAttribute(a2);

			_html.AddHeadElement(meta1);
			CSS css = new CSS();
			Style body = new Style("body");
			body.AddStyle("background-color", "gray");
			Style center = new Style(".center");
			center.AddStyle("text-align", "center");
			css.AddStyle(body);
			css.AddStyle(center);
			_html.AddStyle(css);


			HtmlElement div = new HtmlElement("div");
			HtmlAttribute div_class = new HtmlAttribute("class", "documentbody");
			div.AddAttribute(div_class);
			Style div_style = new Style();
			div_style.AddStyle("font-family","'Times New Roman' 宋体");
			div_style.AddStyle("font-size","10.5pt");
			div_style.AddStyle("margin","0 auto");
			div_style.AddStyle("width","600px");
			div_style.AddStyle("padding","100px 120px");
			div_style.AddStyle("border","2px solid gray");
			div_style.AddStyle("background-color","white");

			div.AddStyle(div_style);
			_html.AddElement(div);
			//HtmlAttribute
			Console.WriteLine(_html.GetHtml());

			StringBuilder html = new StringBuilder();
			html.Append("<!doctype html><html>");
			html.Append(@"<head>
							<style>
							html{
							}
							body{
								background-color:gray;
							}
							.center{
								text-align:center;
							}
							</style>"
						+ "<meta http-equiv=" + "\"X-UA-Compatible\" content=\"IE=edge\">"
						+ "<meta content=\"text/html; charset=UTF-8\" http-equiv=\"content-type\">"
						+ "</head>");
			html.Append("<body>");
			var _style = "\"font-family:'Times New Roman' 宋体;font-size:10.5pt;"
				+ "margin:0 auto; width:600px; padding:100px 120px;"
				+ "border:2px solid gray;background-color:white;\"";
			html.Append("<div class=\"docuemntbody\" style=" + _style + ">");

			#region docuemnt

			WordprocessingDocument doc = WordprocessingDocument.Open(file, false);

			IEnumerable<Paragraph> ps = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();

			foreach (var p in ps)
			{
				HtmlElement _pa = new HtmlElement("p");
				var _p = "<p ";
				var p_class = "";
				#region Paragraph
				//Console.WriteLine(p.InnerText);
				if (p.ParagraphProperties != null)
				{
					var _class = "class=\"noclass ";
					var pPr = p.ParagraphProperties;
					//Console.Write("ParagraphStyle:");
					if (pPr.Justification != null)
					{
						var jc = pPr.Justification;
						_class += jc.Val;
					}
					//if (pPr.ParagraphStyleId != null)
					//{
					//	Console.Write("innerStyle->{0}/", pPr.ParagraphStyleId.Val);
					//}
					//if (pPr.ParagraphMarkRunProperties != null)
					//{
					//	RunProperties rPr = new RunProperties(p.ParagraphProperties.ParagraphMarkRunProperties.OuterXml);
					//	if (rPr.FontSize != null)
					//	{
					//		Console.Write("fontsize->{0}/",rPr.FontSize.Val);
					//	}
					//	if (rPr.RunFonts != null)
					//	{
					//		Console.WriteLine("fontfamily->{0}/", rPr.RunFonts.Ascii + "/" + rPr.RunFonts.HighAnsi + "/" + rPr.RunFonts.ComplexScript);
					//	}
					//}
					//Console.WriteLine();
					_class += "\"";
					_p += _class;
					//_p.Replace("CLASS", _class);
				}
				_pa.AddAttribute("class",
				_p += ">";
				html.Append(_p);
				var runs = p.Elements<Run>();
				foreach (var r in runs)
				{
					var style = "";
					var rpr = r.RunProperties;
					style += helper.GetStyleFromRunProperties(rpr).GetStyle();

					var ts = r.Elements<Text>();
					var pics = r.Elements<Drawing>();
					var content = "";
					foreach (Text t in ts)
					{
						content += t.InnerText;
					}

					foreach (Drawing d in pics)
					{
						//DocumentFormat.OpenXml.Drawing.Pictures.Picture

						DocumentFormat.OpenXml.Drawing.GraphicData gd = d.Inline.Graphic.GraphicData;
						var pic = gd.Elements<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ElementAt(0);
						var pb = pic.BlipFill;
						var blip = pb.Blip;
						var ppp = blip.Embed.Value;
						//var path = doc.GetReferenceRelationship(ppp).Uri;
						//var relationships = doc.Package.GetRelationships();
						//content += ppp;
						content += "<img src=\"media/image1.png\" />";
					}
					var span = "<span style=\"" + style + "\">" + content + "</span>";
					html.Append(span);
					//Console.WriteLine(r.InnerText);

				}
				//Console.WriteLine("----------------------------");
				#endregion
				html.Append("</p>");
			}

			#endregion

			html.Append("</div>");
			html.Append("</body><html>");

			FileStream fs = new FileStream(file + ".html", FileMode.OpenOrCreate, FileAccess.Write);
			StreamWriter sw = new StreamWriter(fs);
			sw.WriteLine(html);
			sw.Close();
			fs.Close();

			XDocument styles = null;

			StylesPart sp = doc.MainDocumentPart.StyleDefinitionsPart == null
				? (StylesPart)doc.MainDocumentPart.StylesWithEffectsPart
				: (StylesPart)doc.MainDocumentPart.StyleDefinitionsPart;
			if (sp != null)
			{
				using (var reader = XmlNodeReader.Create(sp.GetStream(FileMode.Open, FileAccess.Read)))
				{
					styles = XDocument.Load(reader);
				}

			}
			if (styles != null)
			{
				//Console.WriteLine(styles.ToString());
			}


		}

		static string GetString(object obj)
		{

			if (obj != null)
			{
				return Convert.ToString(obj);
			}
			return null;

		}
	}
}
