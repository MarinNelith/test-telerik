using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Styles;
using Telerik.WinForms.Documents.Spreadsheet.Model;

namespace test_telerik
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Telerik.Windows.Documents.Flow.Model.RadFlowDocument document = new
            //Telerik.Windows.Documents.Flow.Model.RadFlowDocument();
            //Telerik.Windows.Documents.Flow.Model.Editing.RadFlowDocumentEditor editor = new Telerik.Windows.Documents.Flow.Model.Editing.RadFlowDocumentEditor(document);
            //editor.InsertText("Ceci est un test !");

            RadFlowDocument document = new RadFlowDocument();

            //////////////////////////////////////////// fonctionnement des sections 

            //Section section = new Section(document); // --> par le constructeur 
            //document.Sections.Add(section);

            Section section = document.Sections.AddSection(); //--> identique que les deux lignes au dessus mais plus simple

            //////////////////////////////////////////// fonctionnement des tables 

            //Lorsque vous créez une instance de la classe Table,
            //vous devez transmettre le document auquel la table appartient comme paramètre au constructeur.    

            //Table emptyTable = new Table(document, 10, 5); // Table object with 10 rows and 5 columns. 
            //section.Blocks.Add(emptyTable);

            //////////////////////////////////////////// code permettant de créer une table numéroter


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

            using (Stream output = File.OpenWrite("test.pdf")) // --> extrait en un pdf
            {
                Telerik.Windows.Documents.Flow.FormatProviders.Pdf.PdfFormatProvider providerPDF = new Telerik.Windows.Documents.Flow.FormatProviders.Pdf.PdfFormatProvider();
                providerPDF.Export(document, output);
            }
        }
    }
}