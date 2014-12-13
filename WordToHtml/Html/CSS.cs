using System.Collections.Generic;
using System.Text;

namespace WordToHtml.Html
{
    public class CSS
    {
        public List<CssStyle> Styles = new List<CssStyle>();
        public void AddStyle(CssStyle s)
        {
            if (!string.IsNullOrEmpty(s.Selector))
            {
                this.Styles.Add(s);
            }
        }
        public override string ToString()
        {
            StringBuilder css = new StringBuilder();
            foreach (CssStyle style in this.Styles)
            {
                css.Append(style);
            }
            return css.ToString();
        }
    }
}