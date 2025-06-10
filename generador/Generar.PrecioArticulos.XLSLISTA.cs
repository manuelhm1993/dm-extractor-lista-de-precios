using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using System.Linq;

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

            // Validar la insersión de imágenes
            // InsertarImagenesEnHoja(rutaArchivo, @"C:\EXPORTA");

            using (SpreadsheetDocument documento = SpreadsheetDocument.Open(rutaArchivo, true))
            {
                int ultimaFila = 3;

                Imprimir_02_datos_0(datos, rutaArchivo, nombreArchivo, documento, ref ultimaFila);
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
                /*documento = UpdateValue("G" + fila, Decimal.Zero, 11, CellValues.Number, documento);
                documento = UpdateValue("H" + fila, r["co_art"], 11, CellValues.String, documento);
                */
                documento = UpdateValue("G" + fila, r["co_art"], 11, CellValues.String, documento);
                documento = UpdateValue("H" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }
        #endregion

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