namespace DocxParser.Models.Code
{
    #region Namespaces
    using DocxParser.Utils;
    using Newtonsoft.Json;
    using System;
    using System.Collections.Generic;
    using System.Linq;
    #endregion

    public class CodeCollector
    {
        public CodeCollector(IEnumerable<DocxCode> codes)
        {
            this.RawCodes = codes.ToList<DocxCode>();
            this.Codes = new List<DocxCode>();
            this.Sections = new HashSet<string>();
        }

        public void Group()
        {
            var codeCache = String.Empty;
            var sectionCache = string.Empty;
            int pageNumCache = -1;
            int[] orderRange = new int[2];
            foreach (var code in this.RawCodes)
            {
                if (code.IsBoundary)
                {
                    if (!String.IsNullOrEmpty(codeCache))
                    {
                        this.Codes.Add(new DocxCode(codeCache, sectionCache, pageNumCache) { OrderRange = orderRange });
                        orderRange = new int[2];
                    }

                    codeCache = String.Empty;
                    sectionCache = String.Empty;
                    pageNumCache = -1;
                    orderRange[0] = code.OrderNum;
                }

                codeCache += code.Code + Constants.CRLF;
                orderRange[1] = code.OrderNum;
                if (String.IsNullOrEmpty(sectionCache) || pageNumCache == -1)
                {
                    sectionCache = code.Section;
                    pageNumCache = code.PageNum;
                }
            }
        }

        public virtual List<DocxCode> GetCodesBySection(string section)
        {
            var ret = new List<DocxCode>();
            foreach (var code in this.Codes)
            {
                if (code.Section == section)
                {
                    ret.Add(code);
                }
            }

            return ret;
        }

        [JsonProperty("raw_codes")]
        public List<DocxCode> RawCodes { get; private set; }

        [JsonProperty("codes")]
        public List<DocxCode> Codes { get; protected set; }

        [JsonProperty("sections")]
        public HashSet<string> Sections { get; protected set; }
    }
}
