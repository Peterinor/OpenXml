namespace WordToHtml.Html
{
    public class Html : HtmlElement
    {
        public Html()
        {
            _head = new HtmlHead();
            _body = new HtmlBody();

            this.AddChild(this._body);
            this.AddChild(this._head);
        }
        string Doctype(string ver = "5")
        {
            if (ver == "5")
            {
                return "<!doctype html>";
            }
            return "";
        }

        private readonly HtmlHead _head;
        private readonly HtmlBody _body;

        public void AddHeadElement(HtmlElement h)
        {
            this._head.AddChild(h);
        }
        public void AddStyle(CSS c)
        {
            this._head.AddStyle(c);
        }
        public void AddElement(HtmlElement b)
        {
            this._body.AddChild(b);
        }
        public override string ToString()
        {
            return this.Doctype() + base.ToString();
        }
    }
}