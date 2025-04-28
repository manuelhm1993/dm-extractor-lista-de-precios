using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

            if (string.IsNullOrWhiteSpace(nombreArchivo))
                throw new ArgumentException("Debes indicar un nombre de archivo válido.");

            if (!nombreArchivo.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                nombreArchivo += ".xlsm";

            DataTable tabla = datos.Tables["XLSLISTA"];

            string rutaArchivo = Path.Combine(rutaSalida, nombreArchivo);
            string plantilla = Path.Combine(Environment.CurrentDirectory, "Generadores", "PlantillaXLSLISTAP.xlsx");

            if (File.Exists(rutaArchivo))
                File.Delete(rutaArchivo);

            File.Copy(plantilla, rutaArchivo);

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(rutaArchivo, true))
            {
                WorksheetPart wsPart = ObtenerHoja(doc, "LISTA");
                SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

                // Elimina solo filas a partir de la fila 3 (sin borrar encabezados ni banners)
                foreach (var row in sheetData.Elements<Row>())
                {
                    if (row.RowIndex.Value >= 3)
                        row.Remove();
                }

                uint filaActual = 3;

                foreach (DataRow row in tabla.Rows)
                {
                    Row nuevaFila = new Row() { RowIndex = filaActual };

                    nuevaFila.Append(
                        CrearCeldaTexto("A", filaActual, row["co_art"].ToString()),
                        CrearCeldaTexto("B", filaActual, row["art_des"].ToString()),
                        CrearCeldaTexto("C", filaActual, ObtenerReferencia(row)),
                        CrearCeldaTexto("D", filaActual, row["cat_des"].ToString()),
                        CrearCeldaNumero("E", filaActual, row["Precio01"]),
                        CrearCeldaNumero("F", filaActual, row["StockActual"]),
                        CrearCeldaNumero("G", filaActual, 0),
                        CrearCeldaTexto("H", filaActual, string.Empty) // Evita fórmula vacía
                    );

                    sheetData.AppendChild(nuevaFila);
                    filaActual++;
                }

                wsPart.Worksheet.Save();
            }
        }

        private static WorksheetPart ObtenerHoja(SpreadsheetDocument doc, string nombreHoja)
        {
            foreach (Sheet hoja in doc.WorkbookPart.Workbook.Sheets)
                if (hoja.Name == nombreHoja)
                    return (WorksheetPart)doc.WorkbookPart.GetPartById(hoja.Id);

            throw new Exception("La hoja " + nombreHoja + " no fue encontrada en el archivo.");
        }

        private static string ObtenerReferencia(DataRow row)
        {
            if (row.Table.Columns.Contains("modelo") && row["modelo"] != DBNull.Value)
                return row["modelo"].ToString();
            return row["co_uni"].ToString();
        }

        private static Cell CrearCeldaTexto(string columna, uint fila, string valor)
        {
            return new Cell
            {
                CellReference = columna + fila,
                DataType = CellValues.String,
                CellValue = new CellValue(valor ?? string.Empty)
            };
        }

        private static Cell CrearCeldaNumero(string columna, uint fila, object valor)
        {
            decimal numero = 0;
            if (valor != DBNull.Value)
                decimal.TryParse(valor.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out numero);

            return new Cell
            {
                CellReference = columna + fila,
                DataType = CellValues.Number,
                CellValue = new CellValue(numero.ToString(CultureInfo.InvariantCulture))
            };
        }
    }
}
