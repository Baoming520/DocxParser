namespace DocxParser.Utils
{
    #region Namespaces.
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.XPath;
    #endregion

    public static class XPathProcessor
    {
        public static XElement Select(this string xml, string xPath, XmlNamespaceManager nsManager)
        {
            XElement xelem = null;
            XElement res = null;
            try
            {
                xelem = XElement.Parse(xml);
                res = xelem.XPathSelectElement(xPath, nsManager);
            }
            catch
            {
                return res;
            }

            return res;
        }

        public static List<XElement> SelectAll(this string xml, string xPath, XmlNamespaceManager nsManager)
        {
            XElement xelem = null;
            IEnumerable<XElement> xelems = null;
            try
            {
                xelem = XElement.Parse(xml);
                xelems = xelem.XPathSelectElements(xPath, nsManager);
            }
            catch
            {
                return null;
            }

            return xelems.ToList();
        }
    }
}
