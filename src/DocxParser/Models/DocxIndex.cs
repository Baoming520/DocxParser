namespace DocxParser.Models
{
    #region Namespaces.
    using DocumentFormat.OpenXml;
    using DocxParser.Utils;
    using System;
    #endregion

    public class DocxIndex : DocxChildEntity
    {
        
        public static bool IsIndexHeader(OpenXmlElement element)
        {
            var value = element.GetParagraphStyle();
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return value == "indexheader";
        }

        public static bool IsIndexEntry(OpenXmlElement element)
        {
            var value = element.GetParagraphStyle();
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return value == "indexentry0";
        }

        public DocxIndex(OpenXmlElement element, string section, int pageNum, int orderNum)
        {
            if (DocxIndex.IsIndexHeader(element))
            {
                this.IndexHeader = true;
                this.Content = element.InnerText;
            }
            else if (DocxIndex.IsIndexEntry(element))
            {
                this.IndexHeader = false;
                this.Content = element.GetInnerText();
            }
            else
            {
                throw new ArgumentException("element");
            }

            this.Section = section;
            this.PageNum = pageNum;
            this.OrderNum = orderNum;
        }

        public override string Category
        {
            get
            {
                return "index";
            }
        }
        public override string Section { get; set; }
        public override int PageNum { get; set; }
        public override int OrderNum { get; set; }
        public bool IndexHeader { get; set; }
        public string Content { get; set; }
    }
}
