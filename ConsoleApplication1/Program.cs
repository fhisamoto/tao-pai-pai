﻿using System;
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
            Lala(sFile);
        }

        [SuppressMessage("ReSharper", "PossiblyMistakenUseOfParamsMethod")]
        private static void Lala(string docName)
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

                var writer = OpenXmlWriter.Create(worksheetPart);

                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                writer.WriteStartElement(new Row());
                writer.WriteElement(new Cell
                {
                    CellValue = new CellValue(DateTime.Now.ToOADate().ToString(new NumberFormatInfo())),
                    StyleIndex = 0
                });
                writer.WriteEndElement();


                writer.WriteEndElement(); //end of SheetData
                writer.WriteEndElement(); //end of worksheet

                writer.Close();

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

            stylesheet.Append(AddInComposit(new Font
            {
                FontName = new FontName {Val = "Calibri"},
                FontSize = new FontSize {Val = 11}
            }, new Fonts()));

            stylesheet.Append(AddInComposit(new Fill
            {
                PatternFill = new PatternFill {PatternType = PatternValues.None}
            }, new Fills()));

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
                BorderId = 0,
                FontId = 0,
                FillId = 0,
                
            }, new CellStyleFormats()));

            stylesheet.Append(AddInComposit(new CellFormat()
            {
                NumberFormatId = numberingFormat.NumberFormatId,
                ApplyNumberFormat = true,
                FormatId = 0
            }, new CellFormats()));

            stylesheet.Append(AddInComposit(new CellStyle
            {
                FormatId = 0,
                Name = "Normal",
                BuiltinId = 0
            }, new CellStyles()));

            return stylesheet;
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