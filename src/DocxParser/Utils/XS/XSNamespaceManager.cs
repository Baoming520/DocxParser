namespace DocxParser.Utils.XS
{
    #region Namespaces
    using System.Collections.Generic;
    using System.Xml;
    #endregion

    public class XSNamespaceManager : XmlNamespaceManager
    {
        public XSNamespaceManager(Dictionary<string, string> relationships)
            : base(new NameTable())
        {
            this.AddNamespace("xs", "http://www.w3.org/2001/XMLSchema");
            this.AddNamespace("xsd", "http://www.w3.org/2001/XMLSchema");
            this.AddNamespace("wsdl", "http://schemas.xmlsoap.org/wsdl/");
            foreach (var item in relationships)
            {
                if (this.DefaultNamespace.Contains(item.Key))
                {
                    continue;
                }

                this.AddNamespace(item.Key, item.Value);
            }
        }
    }
}
