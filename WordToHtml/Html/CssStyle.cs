using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordToHtml.Html
{
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
        public override string ToString()
        {
            StringBuilder style = new StringBuilder();
            foreach (string css in this.Styles.Keys)
            {
                style.Append(css).Append(":").Append(this.Styles[css]).Append(";");
            }
            if (!string.IsNullOrEmpty(this.Selector))
            {
                style.Append(this.Selector).Append("{").Append(style).Append("}");
            }
            return style.ToString();
        }

        public void CombineStyles(CssStyle style)
        {
            foreach (string k in style.Styles.Keys)
            {
                if (this.Styles.Keys.Contains(k))
                {
                    this.Styles[k] = style.Styles[k];
                }
                else
                {
                    this.Styles.Add(k, style.Styles[k]);
                }
            }
        }
    }
}