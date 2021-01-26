namespace DocxParser.Models
{
    // Namespaces
    #region Namespaces.
    using Newtonsoft.Json;
    #endregion

    public class DocxCode : DocxChildEntity
    {
        public DocxCode(string code, string section, int pageNum = -1, int orderNum = -1)
        {
            this.Code = code;
            this.Section = section;
            this.PageNum = pageNum;
            this.OrderNum = orderNum;
            this.OrderRange = new int[0];
        }

        public override string Category
        {
            get
            {
                return "code";
            }
        }

        public override string Section { get; set; }

        public override int PageNum { get; set; }

        public override int OrderNum { get; set; }

        [JsonProperty("order_range")]
        public int[] OrderRange { get; set; }

        [JsonProperty("code")]
        public string Code { get; set; }

        [JsonIgnore]
        public bool IsBoundary { get; set; }

        [JsonIgnore]
        public int Delta
        {
            get
            {
                return this.OrderRange != null && this.OrderRange.Length == 2 && this.OrderRange[1] >= this.OrderRange[0] ?
                    this.OrderRange[1] - this.OrderRange[0] + 1 :
                    1;
            }
        }
    }
}
