namespace DocxParser.Models
{
    #region Namespaces.
    using System;
    using Newtonsoft.Json;
    #endregion

    public class DocxHyperlink : DocxChildEntity, IComparable
    {
        
        public DocxHyperlink(Uri uri)
        {
            if (uri == null)
            {
                throw new ArgumentNullException("uri");
            }

            this.uri = uri;
        }

        public override string Category
        {
            get
            {
                return "hyperlink";
            }
        }

        public override string Section { get; set; }

        public override int PageNum { get; set; }

        public override int OrderNum { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("url")]
        public string URL
        {
            get
            {
                if (uri.OriginalString.EndsWith(".pdf"))
                {
                    return uri.OriginalString;
                }
                else if (!String.IsNullOrEmpty(uri.AbsoluteUri))
                {
                    return uri.AbsoluteUri;
                }
                else
                {
                    return String.Empty;
                }
            }
        }

        [JsonProperty("abstruct")]
        public string Abstruct { get; set; }

        public override bool Equals(object obj)
        {
            DocxHyperlink hl = (DocxHyperlink)obj;

            return hl.Id == this.Id;
        }

        public override int GetHashCode()
        {
            return this.Id.GetHashCode();
        }

        public int CompareTo(object obj)
        {
            DocxHyperlink hl = (DocxHyperlink)obj;

            return Convert.ToInt32(this.Id.Substring(ID_PREFIX.Length)) - Convert.ToInt32(hl.Id.Substring(ID_PREFIX.Length));
        }

        private const string ID_PREFIX = "rId";

        private Uri uri;

        
    }
}
