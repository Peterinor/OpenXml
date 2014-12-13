﻿//using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Linq;
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
            //HtmlElement div = new HtmlElement("div");
            //HtmlAttribute attr = new HtmlAttribute("style", "color:white");
            //div.AddAttribute(attr);

            //HtmlElement p = new HtmlElement("p");
            //p.InnerHtml = "This is a Paragraph!";

            //div.AddChild(p);
            ////Console.WriteLine(div.GetHtml());

            //HtmlHead head = new HtmlHead();
            //head.AddStyleSheet("test.css");
            ////head.AddStyle(
            ////Console.WriteLine(head.GetHtml());

            //HtmlPage page = new HtmlPage();
            ////page.AddHeadElement(
            ////page
            //page.AddElement(div);
            //page.AddElement(p);
            //Console.WriteLine(page.GetHtml());

            xMain(args);
        }

        static void xMain(string[] args)
        {
            string file = "./data/word.docx";
            file = "./data/DirectX_9_3D.docx";

            string ROOT = "./OUT/";

            string fileMd5 = Utities.GetMD5(Path.GetFileName(file));
            string docRoot = Path.Combine(new string[] { ROOT, fileMd5 + "/" });// ROOT + fileMd5 + "/";

            ConvertConfig config = new ConvertConfig();
            config.ResourcePath = "./" + fileMd5 + "/";

            if (!Directory.Exists(docRoot))
            {
                Console.WriteLine("Unzip File");
                UnZip un = new UnZip();
                un.UnZipToDir(file, docRoot);
            }

            Console.WriteLine("Convert Word to Html");

            OpenXmlHelper helper = new OpenXmlHelper();

            Html.Html _html = new Html.Html();

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

            _html.AddHeadElement(meta0);
            _html.AddHeadElement(meta1);
            _html.AddHeadElement(meta2);
            _html.AddHeadElement(meta3);

            CSS css = new CSS();
            CssStyle body = new CssStyle("body");
            body.AddStyle("background-color", "gray");
            CssStyle center = new CssStyle(".center");
            center.AddStyle("text-align", "center");
            css.AddStyle(body);
            //css.AddStyle(center);
            _html.AddStyle(css);

            HtmlElement div = new HtmlElement("div");
            HtmlAttribute div_class = new HtmlAttribute("class", "documentbody");
            div.AddAttribute(div_class);
            CssStyle div_style = new CssStyle();
            div_style.AddStyle("font-family", "'Times New Roman' 宋体");
            div_style.AddStyle("font-size", "10.5pt");
            div_style.AddStyle("margin", "0 auto");
            div_style.AddStyle("width", "600px");
            div_style.AddStyle("padding", "100px 120px");
            div_style.AddStyle("border", "2px solid gray");
            div_style.AddStyle("background-color", "white");

            div.AddStyle(div_style);

            #region docuemnt

            WordprocessingDocument doc = WordprocessingDocument.Open(file, false);

            var mainPart = doc.MainDocumentPart;

            //StyleParts
            StylesPart docstyles = doc.MainDocumentPart.StyleDefinitionsPart == null
                ? (StylesPart)doc.MainDocumentPart.StylesWithEffectsPart
                : (StylesPart)doc.MainDocumentPart.StyleDefinitionsPart;
            var styles = docstyles.Styles;
            //styles
            var __styles = styles.Elements<Style>();
            //var i = __styles.Count();
            //生产Style模版对应的CSS Style
            _html.AddStyle(helper.GetStyles(__styles));

            IEnumerable<Paragraph> ps = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();

            var pps = doc.MainDocumentPart.Document.Body.ChildElements;

            ElementHandler handler = new ElementHandler(doc);

            //处理各个Word元素
            foreach (var pp in pps)
            {
                //Console.WriteLine(pp.GetType().ToString());
                div.AddChild(handler.Handle(pp, config));
            }

            #endregion
            _html.AddElement(div);

            FileStream fs;
            string htmlfile = ROOT + fileMd5 + ".html";
            if (File.Exists(htmlfile))
            {
                fs = new FileStream(htmlfile, FileMode.Truncate, FileAccess.Write);
            }else{
                fs = new FileStream(htmlfile, FileMode.CreateNew, FileAccess.Write);
            }
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(_html.ToString());
            sw.Close();
            fs.Close();

            //
            XDocument _styles = null;

            if (docstyles != null)
            {
                using (var reader = XmlNodeReader.Create(docstyles.GetStream(FileMode.Open, FileAccess.Read)))
                {
                    _styles = XDocument.Load(reader);
                }

            }
            if (_styles != null)
            {
                //Console.WriteLine(_styles.ToString());
            }


        }
    }
}
