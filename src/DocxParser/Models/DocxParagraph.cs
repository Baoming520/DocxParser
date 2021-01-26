namespace DocxParser.Models
{
    #region Namespaces.
    using Newtonsoft.Json;
    #endregion

    public class DocxParagraph : DocxChildEntity
    {
        public DocxParagraph(string content, string section, int pageNum = -1, int orderNum = -1)
        {
            this.Content = content;
            this.Section = section;
            this.PageNum = pageNum;
            this.OrderNum = orderNum;
        }

        public override string Category
        {
            get
            {
                return "paragraph";
            }
        }

        public override string Section { get; set; }
        public override int PageNum { get; set; }
        public override int OrderNum { get; set; }

        [JsonProperty("content")]
        public string Content { get; private set; }

        // TODO:
        public string[] GetWords()
        {
            var words = this.Content.Split(new char[] { ' ', ',', '.', '"', '?', '!', '(', ')', '[', ']', '{', '}' });

            return words;
        }
    }
}
