using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordToHtml.Html;
using ZipHelper;

namespace WordToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            string file = "./data/word.docx";
            file = "./data/DirectX_9_3D.docx";

            const string ROOT = "./OUT/";

            string fileMd5 = Utities.GetMD5(Path.GetFileName(file));
            string docRoot = Path.Combine(new[] { ROOT, fileMd5 + "/" });

            ConvertConfig config = new ConvertConfig
            {
                ResourcePath = "./" + fileMd5 + "/"
            };

            if (!Directory.Exists(docRoot))
            {
                Console.WriteLine("Unzip File");
                UnZip un = new UnZip();
                un.UnZipToDir(file, docRoot);
            }

            Console.WriteLine("Convert Word to Html");

            OpenXmlHelper helper = new OpenXmlHelper();

            Html.Html html = new Html.Html();

            HtmlElement meta0 = new HtmlElement("meta", false);
            meta0.AddAttribute("http-equiv", "X-UA-Compatible");
            meta0.AddAttribute("content", "IE=edge,chrome=1");

            HtmlElement meta1 = new HtmlElement("meta", false);
            meta1.AddAttribute("name", "viewport");
            meta1.AddAttribute("content", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no");

            HtmlElement meta2 = new HtmlElement("meta", false);
            meta2.AddAttribute("name", "apple-mobile-web-app-capable");
            meta2.AddAttribute("content", "yes");

            HtmlElement meta3 = new HtmlElement("meta", false);
            meta3.AddAttribute("http-equiv", "content-type");
            meta3.AddAttribute("content", "text/html; charset=UTF-8");

            html.AddHeadElement(meta0);
            html.AddHeadElement(meta1);
            html.AddHeadElement(meta2);
            html.AddHeadElement(meta3);

            CSS css = new CSS();
            CssStyle body = new CssStyle("body");
            body.AddStyle("background-color", "gray");
            CssStyle center = new CssStyle(".center");
            center.AddStyle("text-align", "center");
            css.AddStyle(body);
            //css.AddStyle(center);
            html.AddStyle(css);

            HtmlElement div = new HtmlElement("div");
            HtmlAttribute divClass = new HtmlAttribute("class", "documentbody");
            div.AddAttribute(divClass);
            CssStyle divStyle = new CssStyle();
            divStyle.AddStyle("font-family", "'Times New Roman' 宋体");
            divStyle.AddStyle("font-size", "10.5pt");
            divStyle.AddStyle("margin", "0 auto");
            divStyle.AddStyle("width", "600px");
            divStyle.AddStyle("padding", "100px 120px");
            divStyle.AddStyle("border", "2px solid gray");
            divStyle.AddStyle("background-color", "white");

            div.AddStyle(divStyle);

            #region docuemnt

            WordprocessingDocument doc = WordprocessingDocument.Open(file, false);

            //StyleParts
            StylesPart docstyles = doc.MainDocumentPart.StyleDefinitionsPart == null
                ? doc.MainDocumentPart.StylesWithEffectsPart
                : (StylesPart)doc.MainDocumentPart.StyleDefinitionsPart;
            var styles = docstyles.Styles;
            //styles
            var styleEl = styles.Elements<Style>();
            //var i = __styles.Count();
            //生产Style模版对应的CSS Style
            html.AddStyle(helper.GetStyles(styleEl));

            var pps = doc.MainDocumentPart.Document.Body.ChildElements;

            ElementHandler handler = new ElementHandler(doc);

            //处理各个Word元素
            foreach (var pp in pps)
            {
                //Console.WriteLine(pp.GetType().ToString());
                div.AddChild(handler.Handle(pp, config));
            }

            #endregion
            html.AddElement(div);

            string htmlfile = ROOT + fileMd5 + ".html";
            FileStream fs = File.Exists(htmlfile) 
                ? new FileStream(htmlfile, FileMode.Truncate, FileAccess.Write) 
                : new FileStream(htmlfile, FileMode.CreateNew, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(html.ToString());
            sw.Close();
            fs.Close();

            ////
            //XDocument _styles = null;

            //if (docstyles != null)
            //{
            //    using (var reader = XmlReader.Create(docstyles.GetStream(FileMode.Open, FileAccess.Read)))
            //    {
            //        _styles = XDocument.Load(reader);
            //    }

            //}
            //if (_styles != null)
            //{
            //    //Console.WriteLine(_styles.ToString());
            //}
        }
    }
}
