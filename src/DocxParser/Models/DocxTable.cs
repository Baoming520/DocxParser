namespace DocxParser.Models
{
    #region Namespaces.
    using DocumentFormat.OpenXml;
    using DocxParser.Utils;
    using System;
    using System.Collections.Generic;
    using System.Xml.Linq;
    #endregion

    public class DocxTable : DocxChildEntity
    {
        public DocxTable(IEnumerable<OpenXmlElement> childElems, string section, int pageNum = -1, int orderNum = -1)
        {
            this.Rows = new List<List<string>>();
            foreach (var elem in childElems)
            {
                if (elem.LocalName == "tblPr")
                {
                    continue;
                }
                if (elem.LocalName == "tblGrid")
                {
                    continue;
                }
                if (elem.LocalName == "tr")
                {
                    string xPath = string.Empty;
                    var cols = new List<string>();
                    foreach (var tcElem in elem.Elements())
                    {
                        if (tcElem.LocalName != "tc")
                        {
                            continue;
                        }

                        var colText = this.GetColumnContent(tcElem);
                        cols.Add(colText);
                    }
                    this.Rows.Add(cols);
                }
            }

            this.Section = section;
            this.PageNum = pageNum;
            this.OrderNum = orderNum;
        }

        public override string Category
        {
            get
            {
                return "table";
            }
        }

        public override string Section { get; set; }

        public override int PageNum { get; set; }
        public override int OrderNum { get; set; }

        public List<List<string>> Rows { get; private set; }

        public string[,] Content
        {
            get
            {
                string[,] content = null;
                if (null != this.Rows && this.Rows.Count > 0 && this.Rows[0].Count > 0)
                {
                    int w = Rows[0].Count, h = Rows.Count;
                    content = new string[h, w];
                    for (int i = 0; i < h; i++)
                    {
                        for (int j = 0; j < w; j++)
                        {
                            content[i, j] = Rows[i][j];
                        }
                    }
                }

                return content;
            }
        }

        #region Private members.
        private string GetColumnContent(OpenXmlElement tcElem)
        {
            var colContent = String.Empty;

            // Get all paragraph elements in the column element of a table
            string xPath = "//w:p";
            var pElems = tcElem.OuterXml.SelectAll(xPath, OpenXmlNamespaceManager.Instance);
            if (pElems == null || pElems.Count == 0)
            {
                return colContent;
            }

            int index = 0;
            foreach (var pElem in pElems)
            {
                xPath = "//w:r/w:t";
                var tElems = pElem.ToString().SelectAll(xPath, OpenXmlNamespaceManager.Instance);
                if (tElems == null || tElems.Count == 0)
                {
                    xPath = "//w:hyperlink/w:r/w:t";
                    tElems = pElem.ToString().SelectAll(xPath, OpenXmlNamespaceManager.Instance);
                }

                if (tElems == null || tElems.Count == 0)
                {
                    continue;
                }

                colContent += this.MergeContent(tElems);
                if (index < pElems.Count - 1)
                {
                    colContent += Constants.CRLF;
                }

                index++;
            }

            return colContent;
        }

        private string MergeContent(IEnumerable<XElement> textElems)
        {
            var colText = string.Empty;
            foreach (var tElem in textElems)
            {
                colText += tElem == null ? string.Empty : tElem.Value;
            }

            return colText;
        }
        #endregion
    }
}
