namespace DocxParser.Models.TechDoc
{
    #region Namespaces.
    using System.Text.RegularExpressions;
    #endregion

    public class TechnicalDocument : Docx
    {
        public TechnicalDocument(string file)
            : base(file)
        {
            var shortNamePattern = "\\[MS-[a-zA-Z0-9]+(-[a-zA-Z0-9]+)*\\]:";
            if (Regex.IsMatch(this.Paragraphs[0].Content, shortNamePattern))
            {
                this.ShortName = Regex.Match(this.Paragraphs[0].Content, "MS-[a-zA-Z0-9]+(-[a-zA-Z0-9]+)*").Value;
                this.FullName = this.Paragraphs[1].Content;
            }
        }

        public string ShortName { get; set; }
        public string FullName { get; set; }
    }
}
