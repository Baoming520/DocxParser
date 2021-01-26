namespace DocxParser.Models
{
    #region Namespaces.
    using DocumentFormat.OpenXml;
    using DocxParser.Utils;
    using Newtonsoft.Json;
    using System;
    #endregion

    public class DocxToC : DocxChildEntity
    {
        public static bool IsToC(OpenXmlElement element)
        {
            string value = element.GetParagraphStyle();
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return value.StartsWith("TOC");
        }

        public DocxToC(OpenXmlElement element)
        {
            // By default: The element MUST be a ToC element.
            this.Level = element.GetParagraphStyle();
            if (this.Level == "TOCHeading")
            {
                string xPath = "//w:r/w:t";
                var elem = XPathProcessor.Select(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
                if (null == elem)
                {
                    throw new NullReferenceException("elem");
                }

                this.Content = elem.Value;
                this.Section = "-1";
                this.PageNum = -1;
            }
            else
            {
                string xPath = "//w:hyperlink/w:r/w:t";
                var elems = XPathProcessor.SelectAll(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
                if (elems[0].Value == "Table of Contents")
                {
                    this.Content = string.Empty;
                    this.Section = SECTION_INVALID;
                    this.PageNum = -1;
                    return;
                }
                // The magic number "3" indicates that the elements should contains: 
                // 1. section info; 
                // 2. content info;
                // 3. page number info.
                if (null == elems || elems.Count != 3)
                {
                    throw new FormatException("elems");
                }

                this.Section = elems[0].Value;
                this.Content = elems[1].Value;
                this.PageNum = Convert.ToInt32(elems[2].Value);
            }

            this.OrderNum = -1;
        }
        public override string Category
        {
            get
            {
                return "table_of_contents";
            }
        }
        public override string Section { get; set; }
        public override int PageNum { get; set; }
        public override int OrderNum { get; set; }

        [JsonProperty("level")]
        public string Level { get; private set; }

        [JsonProperty("content")]
        public string Content { get; private set; }

        private const string SECTION_INVALID = "Invalid";
    }
}
