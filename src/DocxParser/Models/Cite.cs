namespace DocxParser.Models
{
    #region Namespaces
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    #endregion

    public class Cite
    {
        public static Cite GetAppropriateCite(IEnumerable<Cite> cites)
        {
            if (cites.Count() == 0)
            {
                return null;
            }
            if (cites.Count() == 1)
            {
                return cites.ToArray()[0];
            }

            Cite ret = null;
            double minDisParent = Int32.MaxValue, minDisChild = Int32.MaxValue, minDisRelative = Int32.MaxValue;
            Cite citeChild = null, citeParent = null, citeRelative = null;
            foreach (var cite in cites)
            {
                if (cite.Position == CitePosition.Self)
                {
                    ret = cite;
                    break;
                }

                if (cite.Position == CitePosition.Child)
                {
                    if (cite.Distance < minDisChild)
                    {
                        minDisChild = cite.Distance;
                        citeChild = cite;
                    }
                }

                if (cite.Position == CitePosition.Parent)
                {
                    if (cite.Distance < minDisParent)
                    {
                        minDisParent = cite.Distance;
                        citeParent = cite;
                    }
                }

                if (cite.Position == CitePosition.Relative)
                {
                    if (cite.Distance < minDisRelative)
                    {
                        minDisRelative = cite.Distance;
                        citeRelative = cite;
                    }
                }
            }

            if (ret == null && (citeChild != null || citeParent != null))
            {
                ret = minDisChild <= minDisParent ? citeChild : citeParent;
            }

            if (ret == null && citeRelative != null)
            {
                ret = citeRelative;
            }

            return ret;
        }

        public Cite(string citedSection, string section, string code)
        {
            if (!Regex.IsMatch(citedSection, SECTION_PATTERN) ||
                !Regex.IsMatch(section, SECTION_PATTERN))
            {
                throw new FormatException("Bad format for section number.");
            }

            int[] citedSectionPathes = citedSection.Split('.').Select(x => Convert.ToInt32(x)).ToArray();
            int[] sectionPathes = section.Split('.').Select(x => Convert.ToInt32(x)).ToArray();
            int affinity = 0;
            for (int i = 0; i < Math.Min(citedSectionPathes.Length, sectionPathes.Length); i++)
            {
                if (citedSectionPathes[i] == sectionPathes[i])
                {
                    affinity++;
                }
                else
                {
                    break;
                }
            }

            this.Section = section;
            this.Content = code;
            if (citedSection == section)
            {
                this.Position = CitePosition.Self;
                this.Distance = 0;
            }
            else if (section.StartsWith(citedSection))
            {
                this.Position = CitePosition.Child;
                this.Distance = Math.Abs(citedSectionPathes.Length - sectionPathes.Length);
            }
            else if (citedSection.StartsWith(section))
            {
                this.Position = CitePosition.Parent;
                this.Distance = Math.Abs(citedSectionPathes.Length - sectionPathes.Length);
            }
            else if (affinity > 0)
            {
                this.Position = CitePosition.Relative;
                this.Distance = Math.Sqrt(Math.Pow(citedSectionPathes.Length - affinity, 2.0) + Math.Pow(sectionPathes.Length - affinity, 2.0));
            }
            else
            {
                this.Position = CitePosition.Puzzle;
                this.Distance = -1;
            }
        }

        public string Section { get; set; }
        public string Content { get; set; }
        public CitePosition Position { get; set; }
        public double Distance { get; set; }

        public override string ToString()
        {
            string pattern = "Section: {0}\r\nPosition: {1}\r\nDistance: {2}\r\nContent: \r\n--->\r\n{3}\r\n<---";
            return String.Format(pattern, this.Section, this.Position, this.Distance, this.Content);
        }

        private const string SECTION_PATTERN = "^[0-9]+(\\.[0-9]+)*$";
    }

    public enum CitePosition : int
    {
        Self = 0,
        Child = 1,
        Parent = 2,
        Relative = 3,
        Puzzle = -1
    }
}
