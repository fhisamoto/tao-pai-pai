using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConsoleApplication1
{
    public class Program
    {
        private static void Main()
        {
            var sFile = @"teste.xlsx";
            if (File.Exists(sFile))
            {
                File.Delete(sFile);
            }
            WriteXlsx(sFile);
        }

        [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
        private static void WriteXlsx(string docName)
        {
            using (var spreadsheet = SpreadsheetDocument.Create(docName, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());

                var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = CreateStylesheet();
                workbookStylesPart.Stylesheet.Save();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheet = new Sheet
                {
                    Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    var b = new Builder(writer);
                    b.StartEndElement(new Worksheet(),
                        worksheet =>
                        {
                            b.StartEndElement(new SheetData(),
                                sheetData =>
                                {
                                    b.StartEndElement(new Row(), row =>
                                    {
                                        b.Element(new Cell
                                        {
                                            CellValue =
                                                new CellValue(DateTime.Now.ToOADate().ToString(new NumberFormatInfo())),
                                            StyleIndex = 1
                                        });
                                        b.Element(new Cell
                                        {
                                            CellValue = new CellValue("lalala"),
                                            DataType = CellValues.String,
                                            StyleIndex = 2
                                        });
                                    });
                                });
                        });
                }
                // Close the document.
                workbookPart.Workbook.Save();
                spreadsheet.Close();
            }
        }

        [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
        private static Stylesheet CreateStylesheet()
        {
            var stylesheet = new Stylesheet();

            var numberingFormat = new NumberingFormat
            {
                NumberFormatId = UInt32Value.FromUInt32(165),
                FormatCode = StringValue.FromString("dd/MM/yyyy HH:mm:ss")
            };
            stylesheet.Append(AddInComposit(numberingFormat, new NumberingFormats()));

            var fonts = new Fonts();
            AddInComposit(new Font
            {
                FontName = new FontName {Val = "Calibri"},
                FontSize = new FontSize {Val = 11},
                Color = new Color {Indexed = 8},
                FontFamilyNumbering = new FontFamilyNumbering {Val = 2},
                FontScheme = new FontScheme {Val = FontSchemeValues.Minor}
            }, fonts);
            AddInComposit(new Font
            {
                Bold = new Bold(),
                FontSize = new FontSize {Val = 11},
                FontName = new FontName {Val = "Calibri"}
            }, fonts);
            stylesheet.Append(fonts);

            var fills = new Fills();
            AddInComposit(new Fill(), fills);
            AddInComposit(new Fill
            {
                PatternFill = new PatternFill {PatternType = PatternValues.DarkGray}
            }, fills);
            stylesheet.Append(fills);

            stylesheet.Append(AddInComposit(new Border
            {
                LeftBorder = new LeftBorder(),
                RightBorder = new RightBorder(),
                TopBorder = new TopBorder(),
                BottomBorder = new BottomBorder(),
                DiagonalBorder = new DiagonalBorder()
            }, new Borders()));

            stylesheet.Append(AddInComposit(new CellFormat
            {
                ApplyAlignment = new BooleanValue(false),
                ApplyFill = new BooleanValue(false),
                ApplyNumberFormat = new BooleanValue(false),
                ApplyProtection = new BooleanValue(false),
                BorderId = 0,
                FontId = 0,
                FillId = 0
            }, new CellStyleFormats()));

            var cellFormats = new CellFormats();
            AddInComposit(new CellFormat
            {
                ApplyAlignment = new BooleanValue(false),
                ApplyFill = new BooleanValue(false),
                ApplyNumberFormat = new BooleanValue(false),
                ApplyProtection = new BooleanValue(false),
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 0
            }, cellFormats);
            AddInComposit(new CellFormat()
            {
                ApplyAlignment = new BooleanValue(false),
                ApplyFill = new BooleanValue(false),
                ApplyNumberFormat = new BooleanValue(true),
                ApplyProtection = new BooleanValue(false),
                BorderId = 0,
                FillId = 0,
                FontId = 0,
                NumberFormatId = 165
            }, cellFormats);
            AddInComposit(new CellFormat()
            {
                ApplyAlignment = new BooleanValue(false),
                ApplyFill = new BooleanValue(false),
                ApplyFont = new BooleanValue(true),
                ApplyNumberFormat = new BooleanValue(false),
                ApplyProtection = new BooleanValue(false),
                BorderId = 0,
                FillId = 0,
                FontId = 1,
                NumberFormatId = 0
            }, cellFormats);
            stylesheet.Append(cellFormats);

            return stylesheet;
        }

        public class Builder
        {
            public Builder(OpenXmlWriter writer)
            {
                Writer = writer;
            }

            public OpenXmlWriter Writer { get; private set; }

            public void StartEndElement(OpenXmlElement element, Action<OpenXmlElement> action)
            {
                Writer.WriteStartElement(element);
                action(element);
                Writer.WriteEndElement();
            }

            public void Element(OpenXmlElement element, Action<OpenXmlElement> action)
            {
                action(element);
                Writer.WriteElement(element);
            }

            public void Element(OpenXmlElement element)
            {
                Element(element, e => { });
            }
        }

        private static OpenXmlCompositeElement AddInComposit(OpenXmlElement element,
            OpenXmlCompositeElement compositeElement)
        {
            // ReSharper disable once PossiblyMistakenUseOfParamsMethod
            compositeElement.Append(element);
            var propertyInfo = compositeElement.GetType().GetProperty("Count");
            propertyInfo.SetValue(compositeElement, UInt32Value.FromUInt32((uint) compositeElement.ChildElements.Count));
            return compositeElement;
        }
    }
}