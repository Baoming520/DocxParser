namespace DocxParser.Utils
{
    #region Namespaces.
    using DocumentFormat.OpenXml;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;
    #endregion

    public static class OpenXmlExtension
    {
        public delegate string GetSectionNumber(OpenXmlElement element);

        public static GetSectionNumber GetSectionNumberFactory()
        {
            var sectionDict = new Dictionary<SectionIndex, int>();
            return (element) => 
            {
                var pStyle = element.GetParagraphStyle();
                if (!String.IsNullOrEmpty(pStyle) && pStyle.StartsWith("Heading"))
                {
                    if (!sectionDict.ContainsKey(pStyle))
                    {
                        sectionDict.Add(pStyle, 1);
                    }
                    else
                    {
                        sectionDict[pStyle] += 1;
                        int headingNum = Convert.ToInt32(pStyle.Substring("Heading".Length));
                        var removedList = new List<SectionIndex>();
                        foreach (var key in sectionDict.Keys)
                        {
                            int hNum = Convert.ToInt32(key.HeadingNum);
                            if (hNum > headingNum)
                            {
                                removedList.Add(key);
                            }
                        }

                        // Remove from section dictionary
                        foreach (var item in removedList)
                        {
                            sectionDict.Remove(item);
                        }
                    }
                }

                string sectionNum = String.Empty;
                if (sectionDict != null && sectionDict.Keys.Count > 0)
                {
                    var subSections = sectionDict.Keys.ToList();
                    subSections.Sort();
                    for (var i = 0; i < subSections.Count; i++)
                    {
                        sectionNum += sectionDict[subSections[i]];
                        if (i != subSections.Count - 1)
                        {
                            sectionNum += ".";
                        }
                    }

                    return sectionNum;
                }

                return sectionNum;
            };
        }

        public static string GetParagraphStyle(this OpenXmlElement element)
        {
            string xPath = "//w:pPr/w:pStyle";
            var elem = XPathProcessor.Select(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
            if (null == elem)
            {
                return string.Empty;
            }

            XNamespace wNS = OpenXmlNamespaceManager.Instance.LookupNamespace("w");
            var valAttr = elem.Attribute(wNS + "val");
            if (null == valAttr)
            {
                return string.Empty;
            }

            return valAttr.Value;
        }

        public static string[] GetHyperlinkRIds(this OpenXmlElement element, out string[] rNames)
        {
            rNames = null;
            var rIdList = new List<string>();
            var rNameList = new List<string>();
            string xPath = "//w:hyperlink[@r:id]";
            var elems = XPathProcessor.SelectAll(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
            if (elems == null || elems.Count == 0)
            {
                return null;
            }


            XNamespace rNS = OpenXmlNamespaceManager.Instance.LookupNamespace("r");
            foreach (var elem in elems)
            {
                var idAttr = elem.Attribute(rNS + "id");
                if (idAttr == null)
                {
                    continue;
                }

                rNameList.Add(elem.Value);
                rIdList.Add(idAttr.Value);
            }

            rNames = rNameList.ToArray();

            return rIdList.ToArray();
        }

        public static string GetInnerText(this OpenXmlElement element)
        {
            string xPath = "//w:r/w:t";
            var elems = XPathProcessor.SelectAll(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
            var text = string.Empty;
            foreach (var elem in elems)
            {
                if (elem.Name.LocalName != "t")
                {
                    continue;
                }

                text += elem == null ? string.Empty : elem.Value;
            }

            return text;
        }

        public static bool HasBookmark(this OpenXmlElement element)
        {
            string xPath = "//w:bookmarkStart";
            var elem = XPathProcessor.Select(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
            if (null == elem)
            {
                return false;
            }

            xPath = "//w:bookmarkEnd";
            elem = XPathProcessor.Select(element.OuterXml, xPath, OpenXmlNamespaceManager.Instance);
            if (null == elem)
            {
                return false;
            }

            return true;
        }
    }
}
