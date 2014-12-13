namespace WordToHtml.Html
{
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
        public void AddStyle(CSS css)
        {
            HtmlElement c = new HtmlElement("style", css.ToString());
            this.AddChild(c);
        }
        public void AddJavascript() { }
    }
}