using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using System.Linq;

namespace Softech.Administrativo.Generacion
{
    public static class XLSLIS
    {
        #region Generar
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

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(rutaArchivo, true))
            {
                WorkbookPart wbPart = doc.WorkbookPart;
                WorksheetPart wsPart = ObtenerHoja(doc, "LISTA");
                SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

                // Elimina solo filas a partir de la fila 3 (sin borrar encabezados ni banners)
                /*foreach (var row in sheetData.Elements<Row>())
                {
                    if (row.RowIndex.Value >= 3)
                        row.Remove();
                }*/

                sheetData.RemoveAllChildren<Row>(); // Borra todas las filas
                uint filaActual = 3;

                // Obtener el número de registros
                //int numeroRegistros = ObtenerNumeroRegistros(datos);

                foreach (DataRow row in tabla.Rows)
                {
                    Row nuevaFila = new Row() { RowIndex = filaActual };

                    nuevaFila.Append(
                            CrearCeldaTexto("A", filaActual, row["co_art"].ToString()),
                            CrearCeldaTexto("B", filaActual, row["art_des"].ToString()),
                            //CrearCeldaTexto("C", filaActual, ObtenerReferencia(row)),
                            CrearCeldaTexto("C", filaActual, row["co_cat"].ToString()),
                            CrearCeldaTexto("D", filaActual, row["cat_des"].ToString()),
                            CrearCeldaTexto("E", filaActual, row["Precio01"].ToString()),
                            CrearCeldaTexto("F", filaActual, row["StockActual"].ToString()),
                            CrearCeldaNumero("G", filaActual, 0),
                            CrearCeldaTexto("H", filaActual, row["co_art"].ToString())
                        );

                    sheetData.AppendChild(nuevaFila);
                    filaActual++;
                }

                // Guardar el documento
                wsPart.Worksheet.Save();
            }
        }

        private static WorksheetPart ObtenerHoja(SpreadsheetDocument doc, string nombreHoja)
        {
            foreach (Sheet hoja in doc.WorkbookPart.Workbook.Sheets)
                if (hoja.Name == nombreHoja)
                    return (WorksheetPart)doc.WorkbookPart.GetPartById(hoja.Id);

            //throw new Exception($"La hoja '{nombreHoja}' no fue encontrada en el archivo.");
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

        private static Cell CrearCeldaFormula(string columna, uint fila, string formula)
        {
            return new Cell
            {
                CellReference = columna + fila,
                CellFormula = new CellFormula(formula),
                DataType = CellValues.String
            };
        }
        #endregion

        #region Recuperar número de registros
        private static int ObtenerNumeroRegistros(DataSet datos)
        {
            if (datos.Tables.Contains("MetaData"))
            {
                DataTable metaData = datos.Tables["MetaData"];
                if (metaData.Rows.Count > 0)
                {
                    return Convert.ToInt32(metaData.Rows[0]["Registros"]);
                }
            }
            return 0;
        }
        #endregion
    }
}