namespace DocxParser.Models.Extensions
{
    #region Namespaces
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text.RegularExpressions;
    #endregion

    public static class extDocx
    {
        const string SectionLevelGT9Pattern = "[0-9]+(\\.[0-9]+){9,}";
        const string SectionPattern = "[0-9]+(\\.[0-9]+)*";

        /// <summary>
        /// Get the scope of the specified section in the document.
        /// </summary>
        /// <param name="paragraphs">All the paragraphs in the document.</param>
        /// <param name="section">The specified section number like "x.x.x.x..."</param>
        /// <returns>Returns the scope with a tuple. 
        /// Tuple.Item1 indicates closed interval;
        /// Tuple.Item2 indicates open interval.</returns>
        public static Tuple<int, int> GetSectionScope(this IEnumerable<DocxParagraph> paragraphs, string section)
        {
            int start = -1, end = -1;
            int level = section.GetSectionLevel();
            if (level < 9)
            {
                foreach (var paragraph in paragraphs)
                {
                    if (paragraph.Section == section && start == -1)
                    {
                        start = paragraph.OrderNum;
                        continue;
                    }
                    else if (paragraph.Section != section && start != -1 && end == -1)
                    {
                        end = paragraph.OrderNum;
                        break;
                    }
                }


            }
            else if (level == 9)
            {
                foreach (var paragraph in paragraphs)
                {
                    if (paragraph.Section == section && start == -1)
                    {
                        start = paragraph.OrderNum;
                    }

                    if (start >= 0)
                    {
                        // The following three scenarios will indicate the new section will start at the current paragraph:
                        // 1. The level of the section is greater than 9; (Finds a sub-section of current.)
                        // 2. The level of the section is less than 9;
                        // 3. The level of the section is 9, but the section number is not equal to specified one.
                        var regex = new Regex(SectionPattern);
                        if (regex.IsMatch(paragraph.Content))
                        {
                            var match = regex.Match(paragraph.Content);
                            if (match.Value.GetSectionLevel() > 9 && paragraph.Content.StartsWith(match.Value) && match.Value.Contains(section))
                            {
                                end = paragraph.OrderNum;
                                break;
                            }
                        }
                        else if (paragraph.Section.GetSectionLevel() < 9)
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                        else if (paragraph.Section.GetSectionLevel() == 9 &&
                            paragraph.Section.GetTopLevelSection() != section.GetTopLevelSection())
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                    }
                }
            }
            else // The level of the section number is greater than 9.
            {
                var regex = new Regex(SectionLevelGT9Pattern);
                foreach (var paragraph in paragraphs)
                {
                    if (start == -1 && regex.IsMatch(paragraph.Content))
                    {
                        var match = regex.Match(paragraph.Content);
                        if (match.Success && paragraph.Content.StartsWith(match.Value) && match.Value == section)
                        {
                            start = paragraph.OrderNum;
                        }
                    }

                    if (start >= 0 && end == -1)
                    {
                        regex = new Regex(SectionPattern);
                        var match = regex.Match(paragraph.Content);
                        if (match.Success &&
                            match.Value.GetSectionLevel() > 9 &&
                            match.Value.Contains(section) &&
                            match.Value != section &&
                            paragraph.Content.StartsWith(match.Value)) // sub-section
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                        else if (match.Success &&
                           match.Value.GetSectionLevel() > 9 &&
                           match.Value.GetSectionLevel() == section.GetSectionLevel() &&
                           paragraph.Content.StartsWith(match.Value) && section != match.Value) // neighbour section
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                        else if (!match.Success &&
                            paragraph.Section.GetSectionLevel() <= 9 &&
                            paragraph.Section.GetTopLevelSection() != section.GetTopLevelSection())
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                    }
                }
            }

            if (start == -1 || end == -1)
            {
                throw new Exception("The scope of the section cannot be found.");
            }

            return new Tuple<int, int>(start, end);
        }

        public static List<DocxParagraph> GetParagraphs(this IEnumerable<DocxParagraph> paragraphs, string section)
        {
            var ret = new List<DocxParagraph>();
            int level = section.GetSectionLevel();
            if (level < 9)
            {
                foreach (var paragraph in paragraphs)
                {
                    if (paragraph.Section == section)
                    {
                        ret.Add(paragraph);
                    }
                }
            }
            else if (level == 9)
            {
                int start = -1;
                foreach (var paragraph in paragraphs)
                {
                    if (paragraph.Section == section && start == -1)
                    {
                        start = paragraph.OrderNum;
                    }

                    if (start >= 0)
                    {
                        // The following three scenarios will indicate the new section will start at the current paragraph:
                        // 1. The level of the section is greater than 9; (Finds a sub-section of current.)
                        // 2. The level of the section is less than 9;
                        // 3. The level of the section is 9, but the section number is not equal to specified one.
                        var regex = new Regex(SectionPattern);
                        if (regex.IsMatch(paragraph.Content))
                        {
                            var match = regex.Match(paragraph.Content);
                            if (match.Value.GetSectionLevel() > 9 &&
                                match.Value.Contains(section) &&
                                paragraph.Content.StartsWith(match.Value))
                            {
                                // If its sub-section is found, the program will terminate the current foreach loop.
                                break;
                            }
                        }
                        else if (paragraph.Section.GetSectionLevel() < 9)
                        {
                            // If the other section is found here, the program will terminate the current foreach loop.
                            break;
                        }
                        else if (paragraph.Section.GetSectionLevel() == 9 &&
                            paragraph.Section.GetTopLevelSection() != section.GetTopLevelSection())
                        {
                            // If its neighbour section is found, the program will terminate the current foreach loop.
                            break;
                        }

                        ret.Add(paragraph);
                    }
                }
            }
            else // The level of the section number is greater than 9.
            {
                int start = -1, end = -1;
                var regex = new Regex(SectionLevelGT9Pattern);
                foreach (var paragraph in paragraphs)
                {
                    if (start == -1 && regex.IsMatch(paragraph.Content))
                    {
                        var match = regex.Match(paragraph.Content);
                        if (paragraph.Content.StartsWith(section))
                        {
                            start = paragraph.OrderNum;
                        }
                    }

                    if (start >= 0 && end == -1)
                    {
                        regex = new Regex(SectionPattern);
                        var match = regex.Match(paragraph.Content);
                        if (match.Success &&
                            match.Value.GetSectionLevel() > 9 &&
                            match.Value.Contains(section) &&
                            match.Value != section &&
                            paragraph.Content.StartsWith(match.Value) && section != match.Value) // sub-section
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                        else if (match.Success &&
                            match.Value.GetSectionLevel() > 9 &&
                            match.Value.GetSectionLevel() == section.GetSectionLevel() &&
                            paragraph.Content.StartsWith(match.Value) && section != match.Value) // neighbour section
                        {
                            end = paragraph.OrderNum;
                            break;
                        }
                        else if (!match.Success && paragraph.Section.GetSectionLevel() <= 9 &&
                            paragraph.Section.GetTopLevelSection() != section.GetTopLevelSection())
                        {
                            end = paragraph.OrderNum;
                            break;
                        }

                        if (end == -1)
                        {
                            if (match.Success && match.Value == section && paragraph.Content.StartsWith(section))
                            {
                                // Remove the section number from the head of the paragraph.
                                ret.Add(new DocxParagraph(paragraph.Content.Remove(0, section.Length).Trim(), section, paragraph.PageNum, paragraph.OrderNum));
                            }
                            else
                            {
                                ret.Add(new DocxParagraph(paragraph.Content, section, paragraph.PageNum, paragraph.OrderNum));
                            }
                        }
                    }
                }
            }

            return ret;
        }

        public static List<DocxTable> GetTables(this IEnumerable<DocxTable> tables, string section, Tuple<int, int> scope = null)
        {
            var ret = new List<DocxTable>();
            int level = section.GetSectionLevel();
            if (level >= 9 && scope == null)
            {
                throw new ArgumentNullException("If the level of section number is grater than or equal to 9, the scope must not be nullable.");
            }

            // Word document only support 9 levels of heading:
            // 1. When try to get a 9-level section, it will return all the contents in its sub-sections.
            // 2. When try to get a section which its level is greater than or equal to 9, it will get the tables by the parameter "scope". 
            //    So the parameter "scope" cannot be nullable in this case.
            if (level < 9)
            {
                foreach (var table in tables)
                {
                    if (table.Section == section)
                    {
                        ret.Add(table);
                    }
                }
            }
            else
            {
                foreach (var table in tables)
                {
                    if (table.OrderNum >= scope.Item1 && table.OrderNum < scope.Item2)
                    {
                        ret.Add(table);
                    }
                }
            }

            return ret;
        }

        public static DocxToC GetToC(this IEnumerable<DocxToC> tocs, string section)
        {
            foreach (var tocItem in tocs)
            {
                if (tocItem.Section == section)
                {
                    return tocItem;
                }
            }

            return null;
        }

        public static string GetSectionNumberByOrder(this IEnumerable<DocxParagraph> paragraphs, int orderNum)
        {
            string section = String.Empty;
            int index = 0;
            foreach (var paragraph in paragraphs)
            {
                if (paragraph.OrderNum == orderNum)
                {
                    var level = paragraph.Section.GetSectionLevel();
                    if (level < 9)
                    {
                        section = paragraph.Section;
                        break;
                    }
                    else // level == 9, there is no possible to consider the level is greater than 9.
                    {
                        var pArr = paragraphs.ToArray();
                        string curr = String.Empty;
                        string cache = String.Empty;
                        for (int i = index; i >= 0; i--)
                        {
                            
                            curr = pArr[i].Section;
                            if (i == index)
                            {
                                cache = pArr[i].Section;
                            }

                            #region Actually, level = 9
                            // The order number is belong to the section whose real level is 9
                            if (curr != cache)
                            {
                                section = cache;
                                break;
                            }
                            #endregion

                            #region Actually, level > 9
                            // The order number is belong to the section whose real level is greater than 9
                            var match = Regex.Match(pArr[i].Content, SectionLevelGT9Pattern);
                            if (curr == cache && match.Success && pArr[i].Content.StartsWith(match.Value))
                            {
                                section = match.Value;
                                break;
                            }
                            #endregion
                        }

                        if (!String.IsNullOrEmpty(section))
                        {
                            break;
                        }
                    }
                }

                index++;
            }

            return section;
        }

        public static string GetSectionNumberByName(this IEnumerable<DocxToC> tocs, string sectionName, bool ignoreCase = true)
        {
            foreach (var tocItem in tocs)
            {
                if (ignoreCase ?
                    tocItem.Content.ToLower() == sectionName.ToLower() :
                    tocItem.Content == sectionName)
                {
                    return tocItem.Section;
                }
            }

            return String.Empty;
        }

        public static string GetSectionNameByNumber(this IEnumerable<DocxToC> tocs, string sectionNumber)
        {
            foreach (var tocItem in tocs)
            {
                if (tocItem.Section == sectionNumber)
                {
                    return tocItem.Content;
                }
            }

            return String.Empty;
        }

        public static string GetParentSection(this string section)
        {
            string sectionPattern = "^[0-9]+(\\.[0-9]+)*$";
            if (!Regex.IsMatch(section, sectionPattern))
            {
                return null;
            }

            int lastIndex = section.LastIndexOf('.');
            if (lastIndex == -1)
            {
                return section;
            }

            return section.Substring(0, lastIndex);
        }

        public static string GetNeighborSection(this string section)
        {
            throw new NotImplementedException();
        }

        public static string GetTopLevelSection(this string section)
        {
            if (section.GetSectionLevel() < 9)
            {
                return section;
            }

            return section.GetAnyLevelSection(level: 9);
        }

        public static string GetAnyLevelSection(this string section, int level = 9)
        {
            if (!Regex.IsMatch(section, SectionPattern))
            {
                throw new FormatException("The format of the argument 'section' is incorrect.");
            }

            var splits = section.Split('.');
            if (splits.Length < level)
            {
                throw new ArgumentException("level");
            }
            else if (splits.Length == level)
            {
                return section;
            }

            var nums = new string[level];
            for (int i = 0; i < nums.Length; i++)
            {
                nums[i] = splits[i];
            }

            return String.Join(".", nums);
        }

        public static int GetSectionLevel(this string section)
        {
            string sectionPattern = "^[0-9]+(\\.[0-9]+)*$";
            if (!Regex.IsMatch(section, sectionPattern))
            {
                return -1;
            }

            string[] nums = section.Split('.');

            return nums != null ? nums.Length : 0;
        }
    }
}
