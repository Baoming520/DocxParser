namespace DocxParser.Utils
{
    #region Namespaces.
    using System;
    #endregion

    public class SectionIndex : IComparable
    {
        public SectionIndex(string heading)
        {
            this.Heading = heading;
            this.HeadingNum = Convert.ToInt32(heading.Substring("Heading".Length));
        }

        public string Heading { get; private set; }
        public int HeadingNum { get; private set; }

        public override string ToString()
        {
            return this.Heading;
        }

        public override bool Equals(object obj)
        {
            SectionIndex si = (SectionIndex)obj;

            return si.HeadingNum == this.HeadingNum;
        }

        public override int GetHashCode()
        {
            return this.Heading.GetHashCode();
        }

        public int CompareTo(object obj)
        {
            SectionIndex si = (SectionIndex)obj;

            return this.HeadingNum - si.HeadingNum;
        }

        public static implicit operator SectionIndex(string heading)
        {
            return new SectionIndex(heading);
        }
    }
}
