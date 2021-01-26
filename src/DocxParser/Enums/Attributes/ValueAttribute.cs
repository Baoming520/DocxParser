namespace DocxParser.Enums.Attributes
{
    #region Namespaces.
    using System;
    #endregion

    public class ValueAttribute : Attribute
    {
        public ValueAttribute(string value)
        {
            this.Value = value;
        }

        public string Value { get; private set; }
    }
}
