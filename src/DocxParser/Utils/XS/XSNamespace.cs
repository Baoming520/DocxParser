namespace DocxParser.Utils.XS
{
    #region Namespaces
    using DocxParser.Models;
    using System;
    using System.Collections.Generic;
    using System.Text.RegularExpressions;
    #endregion

    public class XSNamespace
    {
        public XSNamespace(DocxToC toc, DocxTable table, string sectionName, string[] tableHeader)
        {
            if (toc == null)
            {
                throw new ArgumentException("toc");
            }
            if (table == null)
            {
                throw new ArgumentException("table");
            }
            if (toc.Section != table.Section)
            {
                throw new ArgumentException("mismatch section number of arguments");
            }
            if (toc.Content != sectionName)
            {
                throw new ArgumentException(String.Format("The section name is not \"Namespaces\" in section {0}", toc.Section));
            }
            if (table.Rows == null ||
                table.Rows[0] == null ||
                table.Rows[0].Count < 2 ||
                table.Rows[0].Count != tableHeader.Length)
            {
                throw new ArgumentException("table.Rows");
            }

            for (int i = 0; i < table.Rows[0].Count; i++)
            {
                if (table.Rows[0][i] != tableHeader[i])
                {
                    throw new ArgumentException("table.Rows[0]");
                }
            }

            this.relationships = new Dictionary<string, string>();
            for (int i = 1; i < table.Rows.Count; i++)
            {
                var c = 0;
                var prefixes = table.Rows[i][c++];
                string[] prefixArr = null;
                if (prefixes.Contains(","))
                {
                    prefixArr = prefixes.Split(',');
                }

                if (prefixArr == null)
                {
                    this.relationships.Add(prefixes, this.FixNamespaceURL(table.Rows[i][c]));
                }
                else
                {
                    foreach (var prefix in prefixArr)
                    {
                        this.relationships.Add(prefix.Trim(), this.FixNamespaceURL(table.Rows[i][c]));
                    }
                }
            }
        }

        public string this[string key]
        {
            get
            {
                return this.relationships[key];
            }
        }

        public Dictionary<string, string> Relationships
        {
            get
            {
                return this.relationships;
            }
        }

        private Dictionary<string, string> relationships;
        private string FixNamespaceURL(string nsURL)
        {
            if (Uri.IsWellFormedUriString(nsURL, UriKind.RelativeOrAbsolute))
            {
                return nsURL;
            }

            string pattern = "<[0-9A-Za-z_.-]+>";
            var matches = Regex.Matches(nsURL, pattern);
            if (matches == null || matches.Count == 0)
            {
                throw new FormatException(String.Format("The parameter 'nsURL={0}' is not a well format URL.", nsURL));
            }

            foreach (var match in matches)
            {
                var m = (Match)match;
                int index = nsURL.IndexOf(m.Value);
                nsURL = nsURL.Remove(index, m.Value.Length);
            }

            return nsURL;
        }
    }
}
