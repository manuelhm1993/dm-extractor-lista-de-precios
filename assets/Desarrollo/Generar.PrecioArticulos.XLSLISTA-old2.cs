
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Softech.Administrativo.Generacion
{
    public static class XLSLIS
    {
        public static void Generar(
            DataSet datos,
            Dictionary<string, object> filtros,
            //Dictionary<string, object> parametros,
            string rutaSalida,
            string nombreArchivo)
        {
            if (datos == null || datos.Tables.Count == 0)
                throw new ArgumentException("No hay datos para generar el archivo.");

            if (!datos.Tables.Contains("XLSLISTA"))
                throw new ArgumentException("La tabla 'XLSLISTA' no está presente en el DataSet.");

            DataTable tabla = datos.Tables["XLSLISTA"];
            string rutaArchivo = Path.Combine(rutaSalida, nombreArchivo);

            // Ruta base a la plantilla lista de precios.xlsm
            string rutaPlantilla = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Generadores" + "\\lista de precios.xlsm");
            File.Copy(rutaPlantilla, rutaArchivo, true);

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(rutaArchivo, true))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                WorksheetPart worksheetPart = null;
                foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                {
                    if (sheet.Name == "LISTA")
                    {
                        worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                        break;
                    }
                }

                if (worksheetPart == null)
                    throw new Exception("No se encontró la hoja 'LISTA' en el archivo XLSM.");

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                uint filaInicio = 3;

                foreach (DataRow row in tabla.Rows)
                {
                    Row nuevaFila = new Row() { RowIndex = filaInicio };

                    // CODIGO (A)
                    nuevaFila.AppendChild(CreateTextCell("A", filaInicio, row["co_art"].ToString()));

                    // DESCRIPCION (B)
                    nuevaFila.AppendChild(CreateTextCell("B", filaInicio, row["art_des"].ToString()));

                    // REFERENCIA (C) (usaremos modelo si está disponible, si no, co_uni)
                    string referencia = row.Table.Columns.Contains("modelo") && row["modelo"] != DBNull.Value
                        ? row["modelo"].ToString()
                        : row["co_uni"].ToString();
                    nuevaFila.AppendChild(CreateTextCell("C", filaInicio, referencia));

                    // MARCA (D) (por ahora usaremos cat_des)
                    nuevaFila.AppendChild(CreateTextCell("D", filaInicio, row["cat_des"].ToString()));

                    // PRECIO $ (E)
                    nuevaFila.AppendChild(CreateNumberCell("E", filaInicio, Convert.ToDecimal(row["Precio01"])));

                    // STOCK (F)
                    nuevaFila.AppendChild(CreateNumberCell("F", filaInicio, Convert.ToDecimal(row["StockActual"])));

                    // PEDIDO (G) siempre 0
                    nuevaFila.AppendChild(CreateNumberCell("G", filaInicio, 0));

                    // IMAGEN (H) fórmula =A#
                    //nuevaFila.AppendChild(CreateFormulaCell("H", filaInicio, $"=A{filaInicio}"));
                    nuevaFila.AppendChild(CreateFormulaCell("H", filaInicio, ""));

                    sheetData.AppendChild(nuevaFila);
                    filaInicio++;
                }

                worksheetPart.Worksheet.Save();
            }
        }

        private static Cell CreateTextCell(string columnName, uint rowIndex, string text)
        {
            return new Cell
            {
                CellReference = columnName + rowIndex,
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };
        }

        private static Cell CreateNumberCell(string columnName, uint rowIndex, decimal number)
        {
            return new Cell
            {
                CellReference = columnName + rowIndex,
                DataType = CellValues.Number,
                CellValue = new CellValue(number.ToString(System.Globalization.CultureInfo.InvariantCulture))
            };
        }

        private static Cell CreateFormulaCell(string columnName, uint rowIndex, string formula)
        {
            return new Cell
            {
                CellReference = columnName + rowIndex,
                CellFormula = new CellFormula(formula),
                DataType = CellValues.String
            };
        }
    }
}
