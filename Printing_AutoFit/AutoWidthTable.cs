using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Printing_AutoFit;

internal class AutoWidthTableBuilder
{
  private readonly Table _table;

  public AutoWidthTableBuilder()
  {
    _table = new();
    _table.AppendChild(CreateTableProps());
  }

  internal AutoWidthTableBuilder WithRow(Action<AutoWidthRowBuilder> cfg)
  {
    var rowBuilder = new AutoWidthRowBuilder();
    cfg(rowBuilder);
    _table.AppendChild(rowBuilder.Build());

    return this;
  }

  internal Table Build()
  {
    return _table;
  }

  private static TableProperties CreateTableProps()
  {
    return new(
      new TableStyle() { Val = "MyGrid" },
      new TableLayout() { Type = TableLayoutValues.Autofit }, // based on docu (http://officeopenxml.com/WPtableLayout.php) can be omitted but lets be exclusive ... I mean explicit
      new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto });
  }
}

internal class AutoWidthRowBuilder
{
  private readonly TableRow _row = new();

  internal AutoWidthRowBuilder WithCell(string text)
  {
    _row.AppendChild(CreateTableCellWithText(text));
    return this;
  }

  internal TableRow Build()
  {
    return _row;
  }

  private static TableCell CreateTableCellWithText(string text)
  {
    return new(
      new TableCellProperties(
        new TableCellWidth()
        {
          Width = "0",
          Type = TableWidthUnitValues.Auto
        }),
      new Paragraph(
        new Run(
          new Text(text))));
  }
}