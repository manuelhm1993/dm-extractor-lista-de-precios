using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using FromMarker2 = DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker;

using System.Linq;
//using System.Runtime.InteropServices;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Softech.Administrativo.Generacion
{
    public static class XLSLIS
    {
        public static void Generar(DataSet datos, Dictionary<string, object> filtros, string rutaSalida, string nombreArchivo)
        {
            if (datos == null || datos.Tables.Count == 0)
                throw new ArgumentException("No hay datos para generar el archivo.");

            if (!datos.Tables.Contains("XLSLISTA"))
                throw new ArgumentException("La tabla 'XLSLISTA' no está presente en el DataSet.");

            DataTable tabla = datos.Tables["XLSLISTA"];
            string rutaArchivo = Path.Combine(rutaSalida, nombreArchivo);
            string rutaPlantilla = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Generadores" + "\\PlantillaXLSLISTA.xlsm");

            File.Copy(rutaPlantilla, rutaArchivo, true);

            using (SpreadsheetDocument documento = SpreadsheetDocument.Open(rutaArchivo, true))
            {
                int ultimaFila = 3;

                Imprimir_02_datos_0(datos, rutaArchivo, nombreArchivo, documento, ref ultimaFila);

                // InsertarImagenesEnExcelOptimizado(rutaArchivo, @"C:\EXPORTA");
                InsertarImagenesEnExcelOptimizado(documento, @"C:\EXPORTA"); // Pasar el objeto SpreadsheetDocument
            }
        }

        #region Imprimir
        private static SpreadsheetDocument Imprimir_02_datos_0(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_02_renglones_0(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            return documento;
        }
        #endregion

        #region InsertCellInWorksheet

        private static Cell InsertCellInWorksheet(Worksheet ws, string addressName)
        {

            SheetData sheetData = ws.GetFirstChild<SheetData>();
            Cell cell = null;

            UInt32 rowNumber = GetRowIndex(addressName);
            Row row = GetRow(sheetData, rowNumber);

            // If the cell you need already exists, return it.
            // If there is not a cell with the specified column name, insert one.  
            Cell refCell = row.Elements<Cell>().
                Where(c => c.CellReference.Value == addressName).FirstOrDefault();
            if (refCell != null)
            {
                cell = refCell;
            }
            else
            {
                cell = CreateCell(row, addressName);
            }
            return cell;
        }

        #endregion

        #region GetRowIndex

        private static UInt32 GetRowIndex(string address)
        {
            string rowPart;
            UInt32 l;
            UInt32 result = 0;

            for (int i = 0; i < address.Length; i++)
            {
                if (UInt32.TryParse(address.Substring(i, 1), out l))
                {
                    rowPart = address.Substring(i, address.Length - i);
                    if (UInt32.TryParse(rowPart, out l))
                    {
                        result = l;
                        break;
                    }
                }
            }
            return result;
        }

        #endregion

        #region GetRow

        private static Row GetRow(SheetData wsData, UInt32 rowIndex)
        {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row();
                row.RowIndex = rowIndex;
                wsData.Append(row);
            }
            return row;
        }

        #endregion

        #region CreateCell

        private static Cell CreateCell(Row row, String address)
        {
            Cell cellResult;
            Cell refCell = null;

            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, address, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            cellResult = new Cell();
            cellResult.CellReference = address;

            row.Append(cellResult);
            return cellResult;
        }

        #endregion

        #region InsertSharedStringItem

        private static int InsertSharedStringItem(WorkbookPart wbPart, object value)
        {
            int index = 0;
            bool found = false;
            var stringTablePart = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

            if (stringTablePart == null)
            {

                stringTablePart = wbPart.AddNewPart<SharedStringTablePart>();
            }

            var stringTable = stringTablePart.SharedStringTable;
            if (stringTable == null)
            {
                stringTable = new SharedStringTable();
            }

            foreach (SharedStringItem item in stringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == value)
                {
                    found = true;
                    break;
                }
                index += 1;
            }

            if (!found)
            {
                stringTable.AppendChild(new SharedStringItem(new Text((String)value)));
                stringTable.Save();
            }
            return index;
        }

        #endregion

        #region Renglones
        private static SpreadsheetDocument Imprimir_02_02_renglones_0(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 3).ToString();
                documento = UpdateValue("A" + fila, r["co_art"], 11, CellValues.String, documento);
                documento = UpdateValue("B" + fila, r["art_des"], 11, CellValues.String, documento);
                documento = UpdateValue("C" + fila, r["des_uni"], 11, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["cat_des"], 11, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["Precio01"], 11, CellValues.Number, documento);
                documento = UpdateValue("F" + fila, r["StockActual"], 11, CellValues.Number, documento);
                documento = UpdateValue("G" + fila, Decimal.Zero, 11, CellValues.Number, documento);
                documento = UpdateValue("H" + fila, r["co_art"], 11, CellValues.String, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }
        #endregion

        /* ************************ INSERTAR IMAGENES CON OPENXML ************************ */
        private static List<string> GetCodes(DataSet dt)
        {
            List<string> codigos = new List<string>();

            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 1).ToString();

                codigos.Add(r["co_art"].ToString());
            }

            return codigos;
        }

        private static Dictionary<string, string> GetImages(DataSet dt)
        {
            Dictionary<string, string> imagenes = new Dictionary<string, string>();

            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 1).ToString();

                imagenes.Add(r["co_art"].ToString(), "");
            }

            return imagenes;
        }

        public static void InsertarImagenesEnExcelOptimizado(SpreadsheetDocument documento, string imagesFolder) // Recibir el objeto SpreadsheetDocument
        {
            WorkbookPart workbookPart = documento.WorkbookPart; // Ahora es accesible

            // 1. Buscar específicamente la hoja "IMAGEN"
            Sheet hojaImagen = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == "IMAGEN");

            if (hojaImagen == null)
                throw new InvalidOperationException("No se encontró la hoja 'IMAGEN' en el archivo Excel.");

            // 2. Obtener la parte de la hoja específica
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(hojaImagen.Id);

            // 3. Configurar el DrawingsPart para las imágenes
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
            {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new Drawing() { Id = workbookPart.GetIdOfPart(drawingsPart) });
                worksheetPart.Worksheet.Save();
            }

            // 4. Configurar WorksheetDrawing
            var worksheetDrawing = drawingsPart.WorksheetDrawing ?? new WorksheetDrawing();
            if (drawingsPart.WorksheetDrawing == null)
            {
                drawingsPart.WorksheetDrawing = worksheetDrawing;
            }

            var imageFiles = Directory.GetFiles(imagesFolder, "*.bmp");
            Console.WriteLine("Procesando " + imageFiles.Length + " imágenes en la hoja IMAGEN...");

            // Configuración mejorada
            const int imageWidth = 50;  // Ancho en píxeles
            const int imageHeight = 50; // Alto en píxeles
            const int startRow = 2;      // Fila inicial (1-based)
            const int codeColumn = 1;    // Columna para códigos (A)
            const int imageColumn = 2;   // Columna para imágenes (B)

            for (int i = 0; i < imageFiles.Length; i++)
            {
                string imagePath = imageFiles[i];
                string code = Path.GetFileNameWithoutExtension(imagePath);
                int currentRow = startRow + i;

                try
                { 
                    // Insertar código en columna A
                    UpdateValueInWorksheet(worksheetPart, "" + (char)('A' + codeColumn - 1) + currentRow, code, 11, CellValues.String);

                    // Insertar imagen
                    using (FileStream imageStream = new FileStream(imagePath, FileMode.Open))
                    {
                        ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Bmp);
                        imagePart.FeedData(imageStream);

                        // Cálculo de posición exacta
                        TwoCellAnchor twoCellAnchor = new TwoCellAnchor(
                            // new FromMarker
                            new FromMarker2 // Verificar ------------------------------------
                            {
                                ColumnId = new UInt32Value((uint)(imageColumn - 1)),
                                RowId = new UInt32Value((uint)(currentRow - 1)),
                                ColumnOffset = new UInt64Value(0),
                                RowOffset = new UInt64Value(0)
                            },
                            new ToMarker
                            {
                                ColumnId = new UInt32Value((uint)(imageColumn - 1)),
                                RowId = new UInt32Value((uint)(currentRow - 1)),
                                ColumnOffset = new UInt64Value((ulong)(imageWidth * 9525)),
                                RowOffset = new UInt64Value((ulong)(imageHeight * 9525))
                            },
                            new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture(
                                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties
                                    {
                                        Id = (uint)i + 2,
                                        Name = "Imagen " + code
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip { Embed = workbookPart.GetIdOfPart(imagePart) },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                        new DocumentFormat.OpenXml.Drawing.Extents
                                        {
                                            Cx = imageWidth * 9525,
                                            Cy = imageHeight * 9525
                                        }),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry
                                    {
                                        Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                                    }))
                            )
                        {
                            EditAs = EditAsValues.OneCell
                        };

                        worksheetDrawing.Append(twoCellAnchor);
                    }

                    // Guardar progreso periódicamente
                    if ((i + 1) % 100 == 0)
                    {
                        drawingsPart.WorksheetDrawing.Save();
                        Console.WriteLine("Procesadas: " + (i + 1) + "/" + imageFiles.Length);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error procesando imagen " + imagePath + ":" + ex.Message);
                }
            }

            // Guardar todos los cambios
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
            Console.WriteLine("¡Imágenes insertadas correctamente en la hoja IMAGEN!");
        }

        // Método auxiliar completo para actualizar celdas en una WorksheetPart específica
        private static void UpdateValueInWorksheet(WorksheetPart worksheetPart, string address, object value, uint styleIndex, CellValues type)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            Cell cell = InsertCellInWorksheet(worksheet, address);
            string strValue = string.Empty;

            switch (type)
            {
                case CellValues.String:
                    if (string.IsNullOrEmpty(Convert.ToString(value)))
                        value = " ";

                    SharedStringTablePart shareStringPart;
                    if (worksheetPart.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                    {
                        shareStringPart = worksheetPart.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    }
                    else
                    {
                        shareStringPart = worksheetPart.WorkbookPart.AddNewPart<SharedStringTablePart>();
                    }

                    int index = 0;
                    bool found = false;
                    foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                    {
                        if (item.InnerText == value.ToString())
                        {
                            found = true;
                            break;
                        }
                        index++;
                    }

                    if (!found)
                    {
                        shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(value.ToString())));
                        shareStringPart.SharedStringTable.Save();
                    }

                    strValue = index.ToString();
                    type = CellValues.SharedString;
                    break;

                case CellValues.Number:
                    if (string.IsNullOrEmpty(Convert.ToString(value)))
                        value = 0;

                    CultureInfo culture = new CultureInfo(CultureInfo.CurrentCulture.LCID);
                    culture.NumberFormat.NumberDecimalSeparator = ".";
                    culture.NumberFormat.NumberGroupSeparator = "";
                    strValue = Convert.ToString(value, culture);
                    break;

                case CellValues.Date:
                    if (string.IsNullOrEmpty(Convert.ToString(value)))
                        value = DateTime.MinValue;
                    strValue = Convert.ToString(((DateTime)value).ToOADate());
                    break;
            }

            cell.CellValue = new CellValue(strValue);
            cell.DataType = new EnumValue<CellValues>(type);

            if (styleIndex > 0)
            {
                cell.StyleIndex = styleIndex;
            }

            worksheet.Save();
        }
        /* ************************ INSERTAR IMAGENES CON OPENXML ************************ */

        #region UpdateValue
        public static SpreadsheetDocument UpdateValue(string addressName, object value, UInt32Value styleIndex, CellValues tipo, SpreadsheetDocument documento)
        {
            WorkbookPart wbPart = documento.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().First();

            if (sheet != null)
            {
                Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;
                Cell cell = InsertCellInWorksheet(ws, addressName);
                String strValor = String.Empty;
                int stringIndex;

                // ------------------ HEAD ------------------ //
                switch (tipo)
                {
                    case CellValues.String:
                        if (String.IsNullOrEmpty(Convert.ToString(value)))
                            value = " ";
                        stringIndex = InsertSharedStringItem(wbPart, value);
                        strValor = stringIndex.ToString();
                        tipo = CellValues.SharedString;
                        //styleIndex = 10;
                        break;
                    case CellValues.Number:
                        if (String.IsNullOrEmpty(Convert.ToString(value)))
                            value = Decimal.Zero;
                        CultureInfo innerCulture = new CultureInfo(CultureInfo.CurrentCulture.LCID);
                        innerCulture.NumberFormat.NumberDecimalSeparator = ".";
                        innerCulture.NumberFormat.NumberGroupSeparator = "";
                        strValor = Convert.ToString((Decimal)value, innerCulture);
                        //styleIndex = 5;
                        break;
                    case CellValues.Date:
                        if (String.IsNullOrEmpty(Convert.ToString(value)))
                            value = DateTime.MinValue;
                        strValor = Convert.ToString(((DateTime)value).ToOADate());
                        break;
                } 

                cell.CellValue = new CellValue(strValor);
                if (tipo != CellValues.Date)
                    cell.DataType = new EnumValue<CellValues>(tipo);

                if (styleIndex > 0)
                    cell.StyleIndex = styleIndex;

                ws.Save();
            }

            return documento;
        }
        #endregion
    }
}