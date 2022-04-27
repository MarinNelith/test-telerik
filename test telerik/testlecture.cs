using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Styles;
using Telerik.WinForms.Documents.Spreadsheet.Model;

namespace test_telerik
{
    internal static class testlecture
    {
        [STAThread]
        static void Main()
        {

            RadFlowDocument document = new RadFlowDocument();

            Table table = document.Sections.AddSection().Blocks.AddTable();
            document.StyleRepository.AddBuiltInStyle(BuiltInStyleNames.TableGridStyleId);
            table.StyleId = BuiltInStyleNames.TableGridStyleId;

            for (int i = 0; i < 5; i++)
            {
                TableRow row = table.Rows.AddTableRow();

                for (int j = 0; j < 10; j++)
                {
                    TableCell cell = row.Cells.AddTableCell();
                    cell.Blocks.AddParagraph().Inlines.AddRun(string.Format("test {0}, {1}", i, j));
                    cell.PreferredWidth = new TableWidthUnit(50);
                }
            }

            //////////////////////////////////////////// code permettant de créer des imageInline

            using (Stream output = new FileStream("C:/Users/marin/test.docx", FileMode.OpenOrCreate)) // --> extrait en un document word
            {
                Telerik.Windows.Documents.Flow.FormatProviders.Docx.DocxFormatProvider provider = new Telerik.Windows.Documents.Flow.FormatProviders.Docx.DocxFormatProvider();
                provider.Export(document, output);
            }

        }
    }
}
