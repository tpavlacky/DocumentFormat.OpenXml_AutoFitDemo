using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace Printing_AutoFit;

internal class FixedWidthTableBuilder
{
  private readonly Table _table;

  public FixedWidthTableBuilder()
  {
    _table = new();
    _table.AppendChild(CreateTableProps());
  }

  internal FixedWidthTableBuilder WithRow(Action<FixedWidthRowBuilder> cfg)
  {
    var rowBuilder = new FixedWidthRowBuilder();
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
      new TableWidth() { Width = "10000", Type = TableWidthUnitValues.Dxa });
  }
}

internal class FixedWidthRowBuilder
{
  private readonly TableRow _row = new();

  internal FixedWidthRowBuilder WithCell(string text)
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
        new Paragraph(
        new Run(
          new Text(text)))));
  }
}