namespace WordToHtml.Html
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
        public override string ToString()
        {
            if (!string.IsNullOrEmpty(this.Name))
            {
                return string.Format("{0}=\"{1}\"", this.Name, (this.Value ?? ""));
            }
            return "";
        }
    }
}