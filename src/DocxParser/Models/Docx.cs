namespace DocxParser.Models
{
    #region Namespaces.
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocxParser.Utils;
    using Newtonsoft.Json;
    using System.Collections.Generic;
    #endregion

    public class Docx
    {
        public Docx(string file)
        {
            this.Initialize();

            int pageNum = 0;
            var getSectionNum = OpenXmlExtension.GetSectionNumberFactory();

            this.hyperlinkRelationshipDict = new Dictionary<string, HyperlinkRelationship>();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(file, true))
            {
                foreach (var hlrs in doc.MainDocumentPart.HyperlinkRelationships)
                {
                    this.hyperlinkRelationshipDict.Add(hlrs.Id, hlrs);
                }

                var isCodeElem = false;
                int orderNum = 0;
                var elements = doc.MainDocumentPart.Document.Body.Elements();
                foreach (var elem in elements)
                {
                    var run = elem.GetFirstChild<Run>();
                    if (run != null)
                    {
                        var lastRenderedPageBreak = run.GetFirstChild<LastRenderedPageBreak>();
                        var pageBreak = run.GetFirstChild<Break>();
                        if (lastRenderedPageBreak != null || pageBreak != null)
                        {
                            // 此方法只可找到大概的页码，真实页码有本地APP决定，并伴随着行间距，字体大小等因素而改变
                            pageNum++;
                        }
                    }

                    var section = getSectionNum(elem);
                    if (DocxToC.IsToC(elem))
                    {
                        this.TableOfContents.Add(new DocxToC(elem));
                    }

                    if (elem.GetType() == typeof(Paragraph))
                    {
                        this.Paragraphs.Add(new DocxParagraph(elem.InnerText, section, pageNum, orderNum));
                        this.GetHyperlinks(((Paragraph)elem).ChildElements, section, pageNum, orderNum);

                        var pStyleVal = elem.GetParagraphStyle();
                        if (!string.IsNullOrEmpty(pStyleVal) && pStyleVal == "Code")
                        {
                            this.Codes.Add(new DocxCode(elem.InnerText, section, pageNum, orderNum) { IsBoundary = !isCodeElem });
                            isCodeElem = true;
                        }
                        else
                        {
                            isCodeElem = false;
                        }

                        if (DocxIndex.IsIndexHeader(elem) || DocxIndex.IsIndexEntry(elem))
                        {
                            this.Indexes.Add(new DocxIndex(elem, section, pageNum, orderNum));
                        }
                    }
                    else if (elem.GetType() == typeof(Table))
                    {
                        this.Tables.Add(new DocxTable(elem.Elements(), section, pageNum, orderNum));
                        var tParagraphs = elem.Descendants<Paragraph>();
                        foreach (var tpElem in tParagraphs)
                        {
                            this.GetHyperlinks(tpElem.ChildElements, section, pageNum, orderNum);

                            // Get the codes from table.
                            var pStyleVal = tpElem.GetParagraphStyle();
                            if (!string.IsNullOrEmpty(pStyleVal) && pStyleVal == "Code")
                            {
                                this.Codes.Add(new DocxCode(elem.InnerText, section, pageNum, orderNum) { IsBoundary = !isCodeElem });
                                isCodeElem = true;
                            }
                            else
                            {
                                isCodeElem = false;
                            }
                        }
                    }

                    orderNum++;
                }
            }
        }

        [JsonProperty("hyperlinks")]
        public List<DocxHyperlink> Hyperlinks { get; set; }

        [JsonProperty("table_of_contents")]
        public List<DocxToC> TableOfContents { get; set; }

        [JsonProperty("paragraphs")]
        public List<DocxParagraph> Paragraphs { get; set; }

        [JsonProperty("codes")]
        public List<DocxCode> Codes { get; set; }

        [JsonProperty("tables")]
        public List<DocxTable> Tables { get; set; }

        [JsonProperty("indexes")]
        public List<DocxIndex> Indexes { get; set; }

        private Dictionary<string, HyperlinkRelationship> hyperlinkRelationshipDict;

        private void Initialize()
        {
            this.Hyperlinks = new List<DocxHyperlink>();
            this.TableOfContents = new List<DocxToC>();
            this.Paragraphs = new List<DocxParagraph>();
            this.Codes = new List<DocxCode>();
            this.Tables = new List<DocxTable>();
            this.Indexes = new List<DocxIndex>();
        }

        private void GetHyperlinks(IEnumerable<OpenXmlElement> elements, string section, int pageNum, int orderNum)
        {
            foreach (var elem in elements)
            {
                if (elem.LocalName == "hyperlink")
                {
                    var hyperlink = (Hyperlink)elem;
                    if (hyperlink.Id != null)
                    {
                        var id = hyperlink.Id.Value;
                        var name = elem.InnerText;
                        this.Hyperlinks.Add(new DocxHyperlink(this.hyperlinkRelationshipDict[id].Uri) { Id = id, Name = name, Section = section, PageNum = pageNum, OrderNum = orderNum, Abstruct = elem.Parent.InnerText });
                    }
                    else
                    {
                        /*
                        * 有如下几种情况会导致条件"hl.id == null"成立：
                        * 1, hyperlink是一个目录条目
                        * 2, Glossary章节中的术语的引用
                        * 3, 各个章节引用的链接
                        * 4, Index章节中的链接
                        */
                    }
                }
            }
        }
    }
}
