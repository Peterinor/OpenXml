using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToHtml.Html;

namespace WordToHtml
{
    public class ElementHandler
    {
        readonly MainDocumentPart _mainPart;
        public ElementHandler(WordprocessingDocument doc)
        {
            this._mainPart = doc.MainDocumentPart;
        }
        public HtmlElement Handle(DocumentFormat.OpenXml.OpenXmlElement elem, ConvertConfig config)
        {
            HtmlElement no = new HtmlElement("no");
            if ((new Regex("Paragraph")).Match(elem.GetType().ToString()).Success)
            {
                return this.HandleParagraph((Paragraph)elem, config);
            }
            if ((new Regex("Table")).Match(elem.GetType().ToString()).Success)
            {
                return this.HandleTable((Table)elem, config);
            }
            return no;
        }

        public HtmlElement HandleParagraph(Paragraph p, ConvertConfig conConfig)
        {
            HtmlElement pa = new HtmlElement("p");

            OpenXmlHelper helper = new OpenXmlHelper();

            StyleConfig config = new StyleConfig
            {
                Charset = "EastAsia"
            };

            #region Paragraph
            var pClass = "";
            if (p.ParagraphProperties != null)
            {
                var pPr = p.ParagraphProperties;
                pa.AddStyle(helper.GetStyleFromParagraphProperties(pPr, config));

                if (pPr.ParagraphStyleId != null)
                {
                    var id = pPr.ParagraphStyleId.Val;
                    pClass += "Inner_P_" + id;
                }
            }
            pa.AddAttribute("class", pClass);
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
                    span.AddChild(new HtmlElement("span", t.InnerText));
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
                    var imgp = this._mainPart.GetPartById(ppp);

                    var img = new HtmlElement("img");
                    //img.AddAttribute("src", imgp.Uri.ToString().Replace("/word/", "./"));
                    img.AddAttribute("src", conConfig.ResourcePath + imgp.Uri.ToString().Substring(1));
                    span.AddChild(img);
                }
                pa.AddChild(span);

            }
            #endregion
            return pa;
        }

        public HtmlElement HandleTable(Table t, ConvertConfig conConfig)
        {
            HtmlElement table = new HtmlElement("table");

            var trs = t.Elements<TableRow>();
            foreach (TableRow tr in trs)
            {
                HtmlElement hTr = new HtmlElement("tr");
                var tcs = tr.Elements<TableCell>();
                foreach (TableCell tc in tcs)
                {
                    HtmlElement hTd = new HtmlElement("td");
                    var pars = tc.Elements<Paragraph>();
                    foreach (Paragraph p in pars)
                    {
                        hTd.AddChild(this.HandleParagraph(p, conConfig));
                    }
                    hTr.AddChild(hTd);
                }
                table.AddChild(hTr);
            }
            return table;
        }
    }
}
