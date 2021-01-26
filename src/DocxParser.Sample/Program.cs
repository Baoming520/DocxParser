namespace DocxParser.Sample
{
    #region Namespaces
    using DocxParser.Models.TechDoc;
    using System;
    #endregion

    class Program
    {
        static void Main(string[] args)
        {
            var td = new TechnicalDocument(@"C:\Users\dxl\Desktop\Protocols\[MS-SSAS].docx");
            foreach (var table in td.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row)
                    {
                        if (cell.Contains("MDACTION_COORDINATE_DIMENSION (2)"))
                        {
                            Console.WriteLine(cell);
                        }
                    }
                }
            }
        }
    }
}
