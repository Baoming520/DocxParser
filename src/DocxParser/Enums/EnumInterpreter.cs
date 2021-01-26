namespace DocxParser.Enums
{
    #region Namespaces.
    using System;
    using System.ComponentModel;
    using System.ComponentModel.DataAnnotations;
    using System.Reflection;
    #endregion

    public static class EnumInterpreter
    {
        public static string GetEnumDescription(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());
            DescriptionAttribute[] descriptionAttrs = (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);
            if (descriptionAttrs.Length > 0)
            {
                return descriptionAttrs[0].Description;
            }
            else
            {
                return value.ToString();
            }
        }

        public static string GetEnumDisplayName(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());
            DisplayAttribute[] descriptionAttrs = (DisplayAttribute[])fi.GetCustomAttributes(typeof(DisplayAttribute), false);
            if (descriptionAttrs.Length > 0)
            {
                return descriptionAttrs[0].Name;
            }
            else
            {
                return value.ToString();
            }
        }
    }
}
