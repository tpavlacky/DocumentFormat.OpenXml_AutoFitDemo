using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Printing_AutoFit;

internal class Program
{
  static void Main()
  {
    var docuName = "test.docx";
    using var doc = WordprocessingDocument.Create(docuName, WordprocessingDocumentType.Document);
    var mainDocumentPart = doc.AddMainDocumentPart();
    mainDocumentPart.Document = new();
    var stylePart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
    var root = new Styles();
    root.Save(stylePart);
    
    stylePart.Styles.Append(TableStyleFactory.GenerateStyle());

    var body = mainDocumentPart.Document.AppendChild(new Body());

    // ================== AUTO-SIZED TABLE ==================
    body.AppendChild(new Paragraph(new Run(new Text("Auto-sized table"))));

    var autoWidthTableBuilder = new AutoWidthTableBuilder();
    autoWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("Karel");
      cfg.WithCell("Novak");
      cfg.WithCell("Novak");
    });
    autoWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("Pepik");
      cfg.WithCell("Zima");
      cfg.WithCell("dsdasdads");
    });
    autoWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("4564sad65as4d6");
      cfg.WithCell("ZimaZdadsaimaZimaZima");
      cfg.WithCell("ZimaZimaZimaZima");
    });
    autoWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("PepikPepikPepik");
      cfg.WithCell("dsadsa");
      cfg.WithCell("adsas");
    });
    autoWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("aa");
      cfg.WithCell("vv");
      cfg.WithCell("cc");
    });

    body.AppendChild(autoWidthTableBuilder.Build());

    // ================== FIXED-SIZED TABLE ==================
    body.AppendChild(new Paragraph());
    body.AppendChild(new Paragraph(new Run(new Text("Fixed-sized table"))));

    var fixedWidthTableBuilder = new FixedWidthTableBuilder();
    fixedWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("Vlasta");
      cfg.WithCell("Novakova");
      cfg.WithCell("11");
    });
    fixedWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("Pepicka");
      cfg.WithCell("Zimova");
      cfg.WithCell("22");
    });
    fixedWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("PepickaPepickaPepickaPepicka");
      cfg.WithCell("ZimovaZimovaZimova");
      cfg.WithCell("33");
    });
    fixedWidthTableBuilder.WithRow(cfg =>
    {
      cfg.WithCell("1");
      cfg.WithCell("2");
      cfg.WithCell("3");
    });

    body.AppendChild(fixedWidthTableBuilder.Build());


    mainDocumentPart.Document.Save();

    Process.Start(docuName);
  }
}