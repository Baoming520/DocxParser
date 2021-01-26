namespace DocxParser.Utils
{
    #region Namespaces.
    using System.Xml;
    #endregion

    public class OpenXmlNamespaceManager : XmlNamespaceManager
    {
        public static OpenXmlNamespaceManager Instance
        {
            get
            {
                return OpenXmlNamespaceManager.instance;
            }
        }

        private static OpenXmlNamespaceManager instance = new OpenXmlNamespaceManager();

        private OpenXmlNamespaceManager()
            : base(new NameTable())
        {
            this.AddNamespace("xml", "http://www.w3.org/XML/1998/namespace");
            this.AddNamespace("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            this.AddNamespace("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            this.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            this.AddNamespace("o", "urn:schemas-microsoft-com:office:office");
            this.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            this.AddNamespace("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            this.AddNamespace("v", "urn:schemas-microsoft-com:vml");
            this.AddNamespace("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            this.AddNamespace("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            this.AddNamespace("w10", "urn:schemas-microsoft-com:office:word");
            this.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            this.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            this.AddNamespace("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            this.AddNamespace("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            this.AddNamespace("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            this.AddNamespace("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            this.AddNamespace("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            this.AddNamespace("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
        }
    }
}
