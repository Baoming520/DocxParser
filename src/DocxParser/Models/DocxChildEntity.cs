namespace DocxParser.Models
{
    // Namespaces
    #region Namespaces.
    using Newtonsoft.Json;
    #endregion

    /// <summary>
    /// A based class of docx entity.
    /// </summary>
    public abstract class DocxChildEntity
    {
        /// <summary>
        /// Identifies the category of the current entity.
        /// </summary>
        [JsonProperty("category")]
        public abstract string Category { get; }

        /// <summary>
        /// Identifies the section number
        /// </summary>
        [JsonProperty("section")]
        public abstract string Section { get; set; }

        /// <summary>
        /// An inaccurate page number.
        /// </summary>
        [JsonProperty("page_num")]
        public abstract int PageNum { get; set; }

        /// <summary>
        /// A number used to identify the order in the document.
        /// </summary>
        [JsonProperty("order_num")]
        public abstract int OrderNum { get; set; }
    }
}
