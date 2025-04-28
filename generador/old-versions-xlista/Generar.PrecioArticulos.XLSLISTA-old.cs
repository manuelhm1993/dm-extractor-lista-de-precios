using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Linq;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
//using System.Windows.Forms;

namespace Softech.Administrativo.Generacion
{
    /// <summary>
    /// Clase para generar datos del sistema
    /// </summary>
    public static class XLSLIS
    {
        #region campos

        private static DataSet ds = default(DataSet);
        private static int iLibroCompra;

        #endregion

        #region Generar

        /// <summary>
        /// Genera una cadena de caracteres con la informacion necesaria para generar el archivo de texto plano
        /// </summary>
        /// <param name="datos">DataSet con los datos necesarios para la generación de la cadena de caracteres</param>
        /// <param name="parametros">Parametros necesarios para la extracción de datos</param>
        /// <param name="ruta"></param>
        /// <param name="nombreArchivo"></param>
        public static void Generar(DataSet datos, Dictionary<String, Object> filtros, Dictionary<String, Object> parametros, String ruta, String nombreArchivo)
        {
            StringBuilder texto = new StringBuilder();

            if (!(datos is DataSet))
                throw new Exception("Es necesario que el objeto de datos sea de tipo DataSet.");
            else
                ds = (DataSet)datos;

            if (ds.Tables.Count.Equals(0))
                throw new Exception("No existen tablas en el DataSet.");

            if (!ds.Tables.Contains("XLSLISTA"))
                throw new Exception("No existe la tabla XLSLISTA en el DataSet.");

            //MessageBox.Show("Pasó todas las validaciones", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //iLibroCompra = ObtenerCombinacionLibroCompra(filtros); 

            iLibroCompra = 0;

            imprimir(datos, ruta, nombreArchivo);
        }

        #endregion

        #region Imprimir

        /// <summary>
        /// Crea el documento y lo escribe
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre que se le dará al archivo</param>
        private static void imprimir(DataSet dt, String ruta, String nombreArchivo)
        {
            SpreadsheetDocument documento;

            if (File.Exists(ruta + nombreArchivo))
                File.Delete(ruta + nombreArchivo);

            String _Plantilla = "PlantillaXLSLISTA.xlsx";

            File.Copy(Environment.CurrentDirectory + @"\Generadores\" + _Plantilla, ruta + nombreArchivo);
            //File.Copy(@"C:\Proyectoss\2KDoceAdministrativo\Administrativo\Softech.Administrativo.Main\bin\x86\Desarrollo\Generadores\Plantilla.xlsx",ruta + nombreArchivo);            documento = SpreadsheetDocument.Open(ruta + nombreArchivo, true);
            FileAttributes attributes = File.GetAttributes(ruta + nombreArchivo);
            if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
            {
                attributes = attributes & ~FileAttributes.ReadOnly;
                File.SetAttributes(ruta + nombreArchivo, attributes);
            }

            documento = SpreadsheetDocument.Open(ruta + nombreArchivo, true);
            documento = Imprimir_01_encabezado(dt, ruta, nombreArchivo, documento);

            int ultimaFila = 1;

            switch (iLibroCompra)
            {
                case 0://Sin columnas adicionales
                    documento = Imprimir_02_datos_0(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 1://Columnas de importación
                    documento = Imprimir_02_datos_1(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 2://Columnas de importación + Columnas Art. 33
                    documento = Imprimir_02_datos_2(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 3://Columnas de importación + Columnas Art. 34
                    documento = Imprimir_02_datos_3(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 4://Columnas de importación + Columnas Art. 33 + Columnas Art. 34
                    documento = Imprimir_02_datos_4(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 5://Columnas Art. 33 
                    documento = Imprimir_02_datos_5(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 6://Columnas Art. 33 + Columnas Art. 34 
                    documento = Imprimir_02_datos_6(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;

                case 7://Columnas Art. 34
                    documento = Imprimir_02_datos_7(dt, ruta, nombreArchivo, documento, ref ultimaFila);
                    break;
            }

            documento = Imprimir_Cuadro_Resumen(dt, documento, ultimaFila);
            documento.Close();
            return;



        }

        #region IMPRIMIR ENCABEZADO DE INFORME

        private static SpreadsheetDocument Imprimir_01_encabezado(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            DataRow r = dt.Tables["XLSLISTA"].Rows[0];

            documento = UpdateValue("A1", "Profit Plus Administrativo", 13, CellValues.String, documento);
            documento = UpdateValue("A2", r["empresa"], 13, CellValues.String, documento);
            documento = UpdateValue("A3", r["rif"], 13, CellValues.String, documento);
            documento = UpdateValue("E2", r["titulo"], 13, CellValues.String, documento);

            MergeTwoCells("T1", "U1", "V1", "V1", documento, 13);
            documento = UpdateValue("T1", "Usuario: " + r["usuario"], 13, CellValues.String, documento);

            MergeTwoCells("T3", "U3", "V3", "V3", documento, 13);
            documento = UpdateValue("T3", "Fecha: " + DateTime.Now, 13, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region IMPRIMIR DATOS

        #region SIN COLUMNAS ADICIONALES

        private static SpreadsheetDocument Imprimir_02_datos_0(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_0(dt, ruta, nombreArchivo, documento);
            /*documento = Imprimir_02_02_renglones_0(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_0(dt, ruta, nombreArchivo, documento, ultimaFila);*/
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_0(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 33 y Art 34

            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Codigo.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "Descripcion", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "Referencia", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "Marca", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "Precio $", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "Stock", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "Pedido", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "Imagen", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_0(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                else
                    documento = UpdateValue("B" + fila, String.Empty, 5, CellValues.String, documento);


                documento = UpdateValue("C" + fila, r["r"].ToString(), 4, CellValues.String, documento);


                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("F" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["n_control"], 4, CellValues.String, documento);
                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("H" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("H" + fila, String.Empty, 4, CellValues.String, documento);
                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);


                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("J" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("J" + fila, String.Empty, 4, CellValues.String, documento);
                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("K" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("L" + fila, r["doc_afec"], 4, CellValues.String, documento);


                documento = UpdateValue("M" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("N" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("O" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("P" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("U" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_0(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            documento = InsertarFormula("SUM(M6:M" + (ultimaFila + 4).ToString() + ")", "M" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(O6:O" + (ultimaFila + 4).ToString() + ")", "O" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas de Importación


        private static SpreadsheetDocument Imprimir_02_datos_1(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            //Debugger.Launch();
            documento = Imprimir_02_01_encabezado_1(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_1(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_1(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_1(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 33 y Art 34

            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "Fecha de Nacionalización", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "Num. Planillas de Importación (C-80 ó C-81)", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "Num. de Expediente de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Valor de la Mercancía en Aduana", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "Derechos de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Tasa por Determinación del Régimen Aplicable", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Otros Aduanales (Gravables)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Otros Aduanales (Exentos)", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Y" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("Z" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("AC" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_1(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                else
                    documento = UpdateValue("B" + fila, String.Empty, 5, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["r"].ToString()))
                    documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("C" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["r"].ToString()))
                    documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("E" + fila, String.Empty, 4, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["FTNac_fechaEmis"].ToString()))
                    documento = UpdateValue("F" + fila, ((DateTime)r["FTNac_fechaEmis"]).Date, 5, CellValues.Date, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 5, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["num_plan_impor"], 4, CellValues.String, documento);
                documento = UpdateValue("H" + fila, r["num_exp_impor"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("J" + fila, r["n_control"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("K" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, String.Empty, 4, CellValues.String, documento);
                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("L" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("L" + fila, String.Empty, 4, CellValues.String, documento);


                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("M" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("M" + fila, String.Empty, 4, CellValues.String, documento);
                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("N" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("N" + fila, "01-reg", 4, CellValues.String, documento);

                documento = UpdateValue("O" + fila, r["doc_afec"], 4, CellValues.String, documento);

                //Columnas de Importación
                documento = UpdateValue("P" + fila, r["valorMercancia"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["der_impor"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["tasa_regimenAplic"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["otrosGravables"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["otrosExentos"], 11, CellValues.Number, documento);

                documento = UpdateValue("U" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("W" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("X" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Y" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("Z" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AA" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AB" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("AC" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_1(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            //Columnas de Importación
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Q6:Q" + (ultimaFila + 4).ToString() + ")", "Q" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);

            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(W6:W" + (ultimaFila + 4).ToString() + ")", "W" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(X6:X" + (ultimaFila + 4).ToString() + ")", "X" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Z6:Z" + (ultimaFila + 4).ToString() + ")", "Z" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AA6:AA" + (ultimaFila + 4).ToString() + ")", "AA" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AB6:AB" + (ultimaFila + 4).ToString() + ")", "AB" + (ultimaFila + 5).ToString(), documento);

            // documento = Imprimir_Cuadro_Resumen(dt, documento, ultimaFila);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas de Importación + Columnas Art. 33 

        private static SpreadsheetDocument Imprimir_02_datos_2(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_2(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_2(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_2(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_2(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulo de cuadro de Art 33
            MergeTwoCells("AA" + filaTitulos.ToString(), "AA" + filaTitulos.ToString(), "AB" + filaTitulos.ToString(), "AB" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. NO Deduc. (Art. 33)", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "Fecha de Nacionalización", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "Num. Planillas de Importación (C-80 ó C-81)", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "Num. de Expediente de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Valor de la Mercancía en Aduana", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "Derechos de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Tasa por Determinación del Régimen Aplicable", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Otros Aduanales (Gravables)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Otros Aduanales (Exentos)", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Y" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("Z" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AC" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("AD" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("AE" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_2(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["FTNac_fechaEmis"].ToString()))
                    documento = UpdateValue("F" + fila, ((DateTime)r["FTNac_fechaEmis"]).Date, 5, CellValues.Date, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 5, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["num_plan_impor"], 4, CellValues.String, documento);
                documento = UpdateValue("H" + fila, r["num_exp_impor"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("J" + fila, r["n_control"], 4, CellValues.String, documento);
                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("K" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, String.Empty, 4, CellValues.String, documento);
                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("L" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("L" + fila, String.Empty, 4, CellValues.String, documento);


                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("M" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("M" + fila, String.Empty, 4, CellValues.String, documento);
                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("N" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("N" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("O" + fila, r["doc_afec"], 4, CellValues.String, documento);

                //Columnas de Importación
                documento = UpdateValue("P" + fila, r["valorMercancia"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["der_impor"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["tasa_regimenAplic"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["otrosGravables"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["otrosExentos"], 11, CellValues.Number, documento);

                documento = UpdateValue("U" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("W" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("X" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Y" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("Z" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AA" + fila, r["base_imponible_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("AB" + fila, r["monto_imp_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("AC" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AD" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("AE" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_2(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            //Columnas de Importación
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Q6:Q" + (ultimaFila + 4).ToString() + ")", "Q" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);

            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(W6:W" + (ultimaFila + 4).ToString() + ")", "W" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(X6:X" + (ultimaFila + 4).ToString() + ")", "X" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Z6:Z" + (ultimaFila + 4).ToString() + ")", "Z" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AA6:AA" + (ultimaFila + 4).ToString() + ")", "AA" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AB6:AB" + (ultimaFila + 4).ToString() + ")", "AB" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AC6:AC" + (ultimaFila + 4).ToString() + ")", "AC" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AD6:AD" + (ultimaFila + 4).ToString() + ")", "AD" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas de Importación + Columnas Art. 34

        private static SpreadsheetDocument Imprimir_02_datos_3(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_3(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_3(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_3(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_3(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 33 y Art 34
            MergeTwoCells("AA" + filaTitulos.ToString(), "AA" + filaTitulos.ToString(), "AB" + filaTitulos.ToString(), "AB" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Total. Deduc. (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            MergeTwoCells("AC" + filaTitulos.ToString(), "AC" + filaTitulos.ToString(), "AD" + filaTitulos.ToString(), "AD" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("AC" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Suj. Prorrateo (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("AD" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);


            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "Fecha de Nacionalización", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "Num. Planillas de Importación (C-80 ó C-81)", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "Num. de Expediente de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Valor de la Mercancía en Aduana", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "Derechos de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Tasa por Determinación del Régimen Aplicable", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Otros Aduanales (Gravables)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Otros Aduanales (Exentos)", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Y" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("Z" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AC" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("AD" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AE" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("AF" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("AG" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_3(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["FTNac_fechaEmis"].ToString()))
                    documento = UpdateValue("F" + fila, ((DateTime)r["FTNac_fechaEmis"]).Date, 5, CellValues.Date, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 5, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["num_plan_impor"], 4, CellValues.String, documento);
                documento = UpdateValue("H" + fila, r["num_exp_impor"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("J" + fila, r["n_control"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("K" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, String.Empty, 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("L" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("L" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("M" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("M" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("N" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("N" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("O" + fila, r["doc_afec"], 4, CellValues.String, documento);

                //Columnas de Importación
                documento = UpdateValue("P" + fila, r["valorMercancia"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["der_impor"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["tasa_regimenAplic"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["otrosGravables"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["otrosExentos"], 11, CellValues.Number, documento);

                documento = UpdateValue("U" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("W" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("X" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Y" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("Z" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AA" + fila, r["base_imponible_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("AB" + fila, r["monto_imp_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("AC" + fila, r["base_imponible_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("AD" + fila, r["monto_imp_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("AE" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AF" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("AG" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_3(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            //Columnas de Importación
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Q6:Q" + (ultimaFila + 4).ToString() + ")", "Q" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);

            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(W6:W" + (ultimaFila + 4).ToString() + ")", "W" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(X6:X" + (ultimaFila + 4).ToString() + ")", "X" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Z6:Z" + (ultimaFila + 4).ToString() + ")", "Z" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AA6:AA" + (ultimaFila + 4).ToString() + ")", "AA" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AB6:AB" + (ultimaFila + 4).ToString() + ")", "AB" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AC6:AC" + (ultimaFila + 4).ToString() + ")", "AC" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AD6:AD" + (ultimaFila + 4).ToString() + ")", "AD" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AE6:AE" + (ultimaFila + 4).ToString() + ")", "AE" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AF6:AF" + (ultimaFila + 4).ToString() + ")", "AF" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas de Importación + Columnas Art. 33 + Columnas Art. 34

        private static SpreadsheetDocument Imprimir_02_datos_4(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_4(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_4(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_4(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_4(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 33 y Art 34
            MergeTwoCells("AA" + filaTitulos.ToString(), "AA" + filaTitulos.ToString(), "AB" + filaTitulos.ToString(), "AB" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. NO Deduc. (Art. 33)", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            MergeTwoCells("AC" + filaTitulos.ToString(), "AC" + filaTitulos.ToString(), "AD" + filaTitulos.ToString(), "AD" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("AC" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Total. Deduc. (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("AD" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            MergeTwoCells("AE" + filaTitulos.ToString(), "AE" + filaTitulos.ToString(), "AF" + filaTitulos.ToString(), "AF" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("AE" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Suj. Prorrateo (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("AF" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);


            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "Fecha de Nacionalización", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "Num. Planillas de Importación (C-80 ó C-81)", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "Num. de Expediente de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Valor de la Mercancía en Aduana", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "Derechos de Importación", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Tasa por Determinación del Régimen Aplicable", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Otros Aduanales (Gravables)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Otros Aduanales (Exentos)", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Y" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("Z" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("AB" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AC" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("AD" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AE" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("AF" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("AG" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("AH" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("AI" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_4(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (!String.IsNullOrEmpty(r["FTNac_fechaEmis"].ToString()))
                    documento = UpdateValue("F" + fila, ((DateTime)r["FTNac_fechaEmis"]).Date, 5, CellValues.Date, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 5, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["num_plan_impor"], 4, CellValues.String, documento);
                documento = UpdateValue("H" + fila, r["num_exp_impor"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("J" + fila, r["n_control"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("K" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, String.Empty, 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("L" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("L" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("M" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("M" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("N" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("N" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("O" + fila, r["doc_afec"], 4, CellValues.String, documento);

                //Columnas de Importación
                documento = UpdateValue("P" + fila, r["valorMercancia"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["der_impor"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["tasa_regimenAplic"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["otrosGravables"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["otrosExentos"], 11, CellValues.Number, documento);

                documento = UpdateValue("U" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("W" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("X" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Y" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("Z" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AA" + fila, r["base_imponible_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("AB" + fila, r["monto_imp_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("AC" + fila, r["base_imponible_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("AD" + fila, r["monto_imp_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("AE" + fila, r["base_imponible_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("AF" + fila, r["monto_imp_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("AG" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("AH" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("AI" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_4(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            //Columnas de Importación
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Q6:Q" + (ultimaFila + 4).ToString() + ")", "Q" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);

            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(W6:W" + (ultimaFila + 4).ToString() + ")", "W" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(X6:X" + (ultimaFila + 4).ToString() + ")", "X" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Z6:Z" + (ultimaFila + 4).ToString() + ")", "Z" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AA6:AA" + (ultimaFila + 4).ToString() + ")", "AA" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AB6:AB" + (ultimaFila + 4).ToString() + ")", "AB" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AC6:AC" + (ultimaFila + 4).ToString() + ")", "AC" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AD6:AD" + (ultimaFila + 4).ToString() + ")", "AD" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AE6:AE" + (ultimaFila + 4).ToString() + ")", "AE" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AF6:AF" + (ultimaFila + 4).ToString() + ")", "AF" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AG6:AG" + (ultimaFila + 4).ToString() + ")", "AG" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(AH6:AH" + (ultimaFila + 4).ToString() + ")", "AH" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas Art. 33

        private static SpreadsheetDocument Imprimir_02_datos_5(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_5(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_5(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_5(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_5(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 33
            MergeTwoCells("S" + filaTitulos.ToString(), "S" + filaTitulos.ToString(), "T" + filaTitulos.ToString(), "T" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. NO Deduc. (Art. 33)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);


            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_5(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("F" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["n_control"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("H" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("H" + fila, String.Empty, 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("J" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("J" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("K" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("L" + fila, r["doc_afec"], 4, CellValues.String, documento);

                documento = UpdateValue("M" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("N" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("O" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("P" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["base_imponible_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["monto_imp_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("U" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("W" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_5(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            documento = InsertarFormula("SUM(M6:M" + (ultimaFila + 4).ToString() + ")", "M" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(O6:O" + (ultimaFila + 4).ToString() + ")", "O" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(V6:V" + (ultimaFila + 4).ToString() + ")", "V" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas Art. 33 + Columnas Art. 34

        private static SpreadsheetDocument Imprimir_02_datos_6(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_6(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_6(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_6(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_6(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 33 y Art 34
            MergeTwoCells("S" + filaTitulos.ToString(), "S" + filaTitulos.ToString(), "T" + filaTitulos.ToString(), "T" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. NO Deduc. (Art. 33)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            MergeTwoCells("U" + filaTitulos.ToString(), "U" + filaTitulos.ToString(), "V" + filaTitulos.ToString(), "V" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("U" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Total. Deduc. (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            MergeTwoCells("W" + filaTitulos.ToString(), "W" + filaTitulos.ToString(), "X" + filaTitulos.ToString(), "X" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("W" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Suj. Prorrateo (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);


            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("Y" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("Z" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("AA" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_6(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("F" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["n_control"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("H" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("H" + fila, String.Empty, 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("J" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("J" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("K" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("L" + fila, r["doc_afec"], 4, CellValues.String, documento);

                documento = UpdateValue("M" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("N" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("O" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("P" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["base_imponible_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["monto_imp_scf"], 11, CellValues.Number, documento);
                documento = UpdateValue("U" + fila, r["base_imponible_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, r["monto_imp_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("W" + fila, r["base_imponible_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("X" + fila, r["monto_imp_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("Y" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Z" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("AA" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_6(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            documento = InsertarFormula("SUM(M6:M" + (ultimaFila + 4).ToString() + ")", "M" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(O6:O" + (ultimaFila + 4).ToString() + ")", "O" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(V6:V" + (ultimaFila + 4).ToString() + ")", "V" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(W6:W" + (ultimaFila + 4).ToString() + ")", "W" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(X6:X" + (ultimaFila + 4).ToString() + ")", "X" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Y6:Y" + (ultimaFila + 4).ToString() + ")", "Y" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(Z6:Z" + (ultimaFila + 4).ToString() + ")", "Z" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion

        #endregion


        #region Columnas Art. 34

        private static SpreadsheetDocument Imprimir_02_datos_7(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            documento = Imprimir_02_01_encabezado_7(dt, ruta, nombreArchivo, documento);
            documento = Imprimir_02_02_renglones_7(dt, ruta, nombreArchivo, documento, ref ultimaFila);
            documento = Imprimir_02_03_totales_7(dt, ruta, nombreArchivo, documento, ultimaFila);
            return documento;
        }

        #region Encabezado

        /// <summary>
        /// Se imprime el encabezado
        /// </summary>
        /// <param name="dt">DataSet</param>
        /// <param name="ruta">Ruta</param>
        /// <param name="nombreArchivo">Nombre del archivo</param>
        /// <param name="documento">Documento a modificar</param>
        /// <returns>Documento modificado</returns>
        private static SpreadsheetDocument Imprimir_02_01_encabezado_7(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento)
        {
            int filaTitulos = 4;

            //FILA 4: Titulos de cuadros de Art 34

            MergeTwoCells("S" + filaTitulos.ToString(), "S" + filaTitulos.ToString(), "T" + filaTitulos.ToString(), "T" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("S" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Total. Deduc. (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);

            MergeTwoCells("U" + filaTitulos.ToString(), "U" + filaTitulos.ToString(), "V" + filaTitulos.ToString(), "V" + filaTitulos.ToString(), documento, 7);
            documento = UpdateValue("U" + filaTitulos.ToString(), "Inform. de Compras con Cred. Fisc. Suj. Prorrateo (Art. 34)", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), String.Empty, 7, CellValues.String, documento);


            filaTitulos++;
            //FILA 5: Titulos de columnas
            documento = UpdateValue("E" + filaTitulos.ToString(), "Tipo Prov.", 7, CellValues.String, documento);
            documento = UpdateValue("F" + filaTitulos.ToString(), "N° de Factura", 7, CellValues.String, documento);
            documento = UpdateValue("G" + filaTitulos.ToString(), "N° de Control", 7, CellValues.String, documento);
            documento = UpdateValue("H" + filaTitulos.ToString(), "N° de Nota de Débito", 7, CellValues.String, documento);
            documento = UpdateValue("I" + filaTitulos.ToString(), "N° de Nota de Crédito", 7, CellValues.String, documento);
            documento = UpdateValue("J" + filaTitulos.ToString(), "Número de Comprobante", 7, CellValues.String, documento);
            documento = UpdateValue("K" + filaTitulos.ToString(), "Tipo de Transcc.", 7, CellValues.String, documento);
            documento = UpdateValue("L" + filaTitulos.ToString(), "N° de Factura Afectada", 7, CellValues.String, documento);
            documento = UpdateValue("M" + filaTitulos.ToString(), "Total de Compras Incluye I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("N" + filaTitulos.ToString(), "Compras no Sujetas", 7, CellValues.String, documento);
            documento = UpdateValue("O" + filaTitulos.ToString(), "Compras sin Derecho a Crédito I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("P" + filaTitulos.ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("Q" + filaTitulos.ToString(), "% Alic.", 7, CellValues.String, documento);
            documento = UpdateValue("R" + filaTitulos.ToString(), "Impuesto I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("S" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("T" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("U" + filaTitulos.ToString(), "B. Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("V" + filaTitulos.ToString(), "Imp. I.V.A.", 7, CellValues.String, documento);
            documento = UpdateValue("W" + filaTitulos.ToString(), "I.V.A. Retenido (al vendedor)", 7, CellValues.String, documento);
            documento = UpdateValue("X" + filaTitulos.ToString(), "I.V.A. Retenido (a terceros)", 7, CellValues.String, documento);
            documento = UpdateValue("Y" + filaTitulos.ToString(), "Anticipo de I.V.A. (importación)", 7, CellValues.String, documento);

            return documento;
        }

        #endregion

        #region Renglones

        private static SpreadsheetDocument Imprimir_02_02_renglones_7(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, ref int ultimaFila)
        {
            for (int i = 0; i < dt.Tables["XLSLISTA"].Rows.Count; i++)
            {
                DataRow r = dt.Tables["XLSLISTA"].Rows[i];
                String fila = (i + 6).ToString();
                documento = UpdateValue("A" + fila, (i + 1).ToString(), 11, CellValues.String, documento);
                if (!String.IsNullOrEmpty(r["fecha_emis"].ToString()))
                    documento = UpdateValue("B" + fila, ((DateTime)r["fecha_emis"]).Date, 5, CellValues.Date, documento);
                documento = UpdateValue("C" + fila, r["r"], 4, CellValues.String, documento);
                documento = UpdateValue("D" + fila, r["prov_des"], 4, CellValues.String, documento);
                documento = UpdateValue("E" + fila, r["tipo_prov"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "FACT" && r["doc_afec"].ToString().Trim() == String.Empty)
                    documento = UpdateValue("F" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("F" + fila, String.Empty, 4, CellValues.String, documento);

                documento = UpdateValue("G" + fila, r["n_control"], 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/DB" && (Decimal)r["monto_ret_imp"] + (Decimal)r["monto_ret_imp_tercero"] == 0)
                    documento = UpdateValue("H" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("H" + fila, String.Empty, 4, CellValues.String, documento);

                if (r["co_tipo_doc"].ToString().Trim() == "N/CR" && Convert.ToDecimal(r["monto_ret_imp"]) + Convert.ToDecimal(r["monto_ret_imp_tercero"]) == 0)
                    documento = UpdateValue("I" + fila, r["nro_fact"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("I" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Decimal)r["total_neto"] <= 0)
                    documento = UpdateValue("J" + fila, r["num_comprobante"], 4, CellValues.String, documento);
                else
                    documento = UpdateValue("J" + fila, String.Empty, 4, CellValues.String, documento);

                if ((Boolean)r["anulado"] == true)
                    documento = UpdateValue("K" + fila, "03-anu", 4, CellValues.String, documento);
                else
                    documento = UpdateValue("K" + fila, "01-reg", 4, CellValues.String, documento);
                documento = UpdateValue("L" + fila, r["doc_afec"], 4, CellValues.String, documento);

                documento = UpdateValue("M" + fila, r["total_neto"], 11, CellValues.Number, documento);
                documento = UpdateValue("N" + fila, String.Empty, 11, CellValues.String, documento); //compras no sujetas
                documento = UpdateValue("O" + fila, r["compras_exentas"], 11, CellValues.Number, documento);
                documento = UpdateValue("P" + fila, r["base_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("Q" + fila, r["tasa"], 11, CellValues.Number, documento);
                documento = UpdateValue("R" + fila, r["monto_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("S" + fila, r["base_imponible_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("T" + fila, r["monto_imp_deducible"], 11, CellValues.Number, documento);
                documento = UpdateValue("U" + fila, r["base_imponible_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("V" + fila, r["monto_imp_prorrateo"], 11, CellValues.Number, documento);
                documento = UpdateValue("W" + fila, r["monto_ret_imp"], 11, CellValues.Number, documento);
                documento = UpdateValue("X" + fila, r["monto_ret_imp_tercero"], 11, CellValues.Number, documento);
                documento = UpdateValue("Y" + fila, Decimal.Zero, 11, CellValues.Number, documento);

                ultimaFila = i + 2;
            }
            return documento;
        }

        #endregion

        #region Totales

        private static SpreadsheetDocument Imprimir_02_03_totales_7(DataSet dt, String ruta, String nombreArchivo, SpreadsheetDocument documento, int ultimaFila)
        {

            documento = InsertarFormula("SUM(M6:M" + (ultimaFila + 4).ToString() + ")", "M" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(O6:O" + (ultimaFila + 4).ToString() + ")", "O" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(P6:P" + (ultimaFila + 4).ToString() + ")", "P" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(R6:R" + (ultimaFila + 4).ToString() + ")", "R" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(S6:S" + (ultimaFila + 4).ToString() + ")", "S" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(T6:T" + (ultimaFila + 4).ToString() + ")", "T" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(U6:U" + (ultimaFila + 4).ToString() + ")", "U" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(V6:V" + (ultimaFila + 4).ToString() + ")", "V" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(W6:W" + (ultimaFila + 4).ToString() + ")", "W" + (ultimaFila + 5).ToString(), documento);
            documento = InsertarFormula("SUM(X6:X" + (ultimaFila + 4).ToString() + ")", "X" + (ultimaFila + 5).ToString(), documento);

            return documento;
        }

        #endregion


        #endregion


        #endregion

        #region IMPRIMIR CUADRO RESUMEN

        private static SpreadsheetDocument Imprimir_Cuadro_Resumen(DataSet dt, SpreadsheetDocument documento, int ultimaFila)
        {
            DataRow r = dt.Tables["XLSLISTA"].Rows[0];
            //titulos
            documento = UpdateValue("J" + (ultimaFila + 7).ToString(), "Base Imponible", 7, CellValues.String, documento);
            documento = UpdateValue("K" + (ultimaFila + 7).ToString(), "Crédito fiscal", 7, CellValues.String, documento);
            documento = UpdateValue("L" + (ultimaFila + 7).ToString(), "IVA retenido por el comprador", 7, CellValues.String, documento);
            documento = UpdateValue("M" + (ultimaFila + 7).ToString(), "IVA retenido (a terceros)", 7, CellValues.String, documento);

            int fila = ultimaFila + 8;
            for (int i = 1; i < 52; i++)
            {
                int campoAnterior = (i == 1)
                                    ? i
                                    : (i - 1);

                int campoProximo = (i == 52)
                                    ? i
                                    : (i + 1);

                if (!String.IsNullOrEmpty((r["descrip" + i.ToString()]).ToString())
                    || !String.IsNullOrEmpty((r["descrip" + campoAnterior.ToString()]).ToString())
                    || !String.IsNullOrEmpty((r["descrip" + campoProximo.ToString()]).ToString()))
                {
                    MergeTwoCells("F" + fila.ToString(), "G" + fila.ToString(), "H" + fila.ToString(), "I" + fila.ToString(), documento, 13);
                    documento = UpdateValue("F" + fila.ToString(), r["descrip" + i.ToString()], 13, CellValues.String, documento);
                    if (!String.IsNullOrEmpty((r["base_imp" + i.ToString()]).ToString()))
                    {
                        documento = UpdateValue("J" + fila.ToString(), r["base_imp" + i.ToString()], 11, CellValues.Number, documento);
                        documento = UpdateValue("K" + fila.ToString(), r["monto_imp" + i.ToString()], 11, CellValues.Number, documento);
                        documento = UpdateValue("L" + fila.ToString(), r["retenido" + i.ToString()], 11, CellValues.Number, documento);
                        documento = UpdateValue("M" + fila.ToString(), r["retenidoter" + i.ToString()], 11, CellValues.Number, documento);
                    }
                    else
                    {
                        documento = UpdateValue("J" + fila.ToString(), String.Empty, 4, CellValues.String, documento);
                        documento = UpdateValue("K" + fila.ToString(), String.Empty, 4, CellValues.String, documento);
                        documento = UpdateValue("L" + fila.ToString(), String.Empty, 4, CellValues.String, documento);
                        documento = UpdateValue("M" + fila.ToString(), String.Empty, 4, CellValues.String, documento);

                    }
                    fila++;
                }
            }

            return documento;
        }

        #endregion

        #endregion

        #region METODOS

        #region ObtenerCombinacionLibroCompra

        /// <summary>
        /// Método empleado para determinar la distribución de las columnas del reporte según los filtros enviados por el usuario
        /// </summary>
        /// <param name="filtros">Filtros enviados por el usuario</param>
        /// <returns>0: Sin columnas adicionales </returns>
        /// <returns>1: Columnas de importación </returns>
        /// <returns>2: Columnas de importación + Columnas Art. 33 </returns>
        /// <returns>3: Columnas de importación + Columnas Art. 34</returns>
        /// <returns>4: Columnas de importación + Columans Art. 33 + Columnas Art. 34</returns>
        /// <returns>5: Columnas Art. 33 </returns>
        /// <returns>6: Columnas Art. 33 + Columnas Art. 34 </returns>
        /// <returns>7: Columnas Art. 34 </returns>

        private static int ObtenerCombinacionLibroCompra(Dictionary<String, Object> filtros)
        {

            bool ColumnasImport = (filtros["f_filtro3"] == "NO" || Equals(filtros["f_filtro3"], null))
                                        ? false
                                        : true;

            bool ColumnasArt33 = (filtros["f_filtro4"] == "NO")
                                        ? false
                                        : true;

            bool ColumnasArt34 = (filtros["f_filtro5"] == "NO")
                                        ? false
                                        : true;

            if (ColumnasImport && !ColumnasArt33 && !ColumnasArt34)
                return 1;

            if (ColumnasImport && ColumnasArt33 && !ColumnasArt34)
                return 2;

            if (ColumnasImport && !ColumnasArt33 && ColumnasArt34)
                return 3;

            if (ColumnasImport && ColumnasArt33 && ColumnasArt34)
                return 4;

            if (!ColumnasImport && ColumnasArt33 && !ColumnasArt34)
                return 5;

            if (!ColumnasImport && ColumnasArt33 && ColumnasArt34)
                return 6;

            if (!ColumnasImport && !ColumnasArt33 && ColumnasArt34)
                return 7;

            return 0;
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

        #region InsertarFormula

        private static SpreadsheetDocument InsertarFormula(String Formula, String celdaResultado, SpreadsheetDocument documento)
        {
            WorkbookPart wbPart = documento.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().First();


            Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;
            Cell cell = InsertCellInWorksheet(ws, celdaResultado);

            //Style 12 = Número con 2 decimales y "." como separador de miles, sin bordes. Si cambian los estilos, aqui tambien hay que modificar
            UInt32Value styleIndex = 12;

            //if (styleIndex > 0)
            cell.StyleIndex = styleIndex;

            cell.CellFormula = new CellFormula(Formula);

            ws.Save();


            return documento;
        }
        #endregion

        #region unir celdas

        private static SpreadsheetDocument MergeTwoCells(string celda1, string celda2, string celda3, string celda4, SpreadsheetDocument documento, UInt32Value styleIndex)
        {


            WorkbookPart wbPart = documento.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().First();
            Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;

            if (ws == null || string.IsNullOrEmpty(celda1) || string.IsNullOrEmpty(celda2) || string.IsNullOrEmpty(celda3) || string.IsNullOrEmpty(celda4))
            {
                return documento;
            }

            documento = CreateSpreadsheetCellIfNotExist(documento, celda1, styleIndex);
            documento = CreateSpreadsheetCellIfNotExist(documento, celda2, styleIndex);
            documento = CreateSpreadsheetCellIfNotExist(documento, celda2, styleIndex);
            documento = CreateSpreadsheetCellIfNotExist(documento, celda2, styleIndex);

            MergeCells mergeCells;
            if (ws.Elements<MergeCells>().Count() > 0)
            {
                mergeCells = ws.Elements<MergeCells>().First();
                MergeCell mergeCell = new MergeCell() { Reference = new StringValue(celda1 + ":" + celda4) };
                mergeCells.Append(mergeCell);

                ws.Save();
            }
            else
            {
                mergeCells = new MergeCells();


                if (ws.Elements<CustomSheetView>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<CustomSheetView>().First());
                }
                else if (ws.Elements<DataConsolidate>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<DataConsolidate>().First());
                }
                else if (ws.Elements<SortState>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<SortState>().First());
                }
                else if (ws.Elements<AutoFilter>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<AutoFilter>().First());
                }
                else if (ws.Elements<Scenarios>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<Scenarios>().First());
                }
                else if (ws.Elements<ProtectedRanges>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<ProtectedRanges>().First());
                }
                else if (ws.Elements<SheetProtection>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<SheetProtection>().First());
                }
                else if (ws.Elements<SheetCalculationProperties>().Count() > 0)
                {
                    ws.InsertAfter(mergeCells, ws.Elements<SheetCalculationProperties>().First());
                }
                else
                {
                    ws.InsertAfter(mergeCells, ws.Elements<SheetData>().First());
                }


                MergeCell mergeCell = new MergeCell() { Reference = new StringValue(celda1 + ":" + celda4) };
                mergeCells.Append(mergeCell);

                ws.Save();

            }
            return documento;
        }

        #endregion

        #region CreateSpreadSheetCellIfNotExist

        private static SpreadsheetDocument CreateSpreadsheetCellIfNotExist(SpreadsheetDocument documento, string cellName, UInt32Value styleIndex)
        {
            WorkbookPart wbPart = documento.WorkbookPart;
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().First();
            Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;

            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = ws.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);


            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                cell.StyleIndex = styleIndex;
                row.Append(cell);
                ws.Descendants<SheetData>().First().Append(row);
                ws.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    cell.StyleIndex = styleIndex;
                    row.Append(cell);
                    ws.Save();
                }
            }
            return documento;
        }

        #endregion

        #region GetWorkSheet

        private static Worksheet GetWorksheet(SpreadsheetDocument documento, string worksheetName)
        {
            IEnumerable<Sheet> sheets = documento.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
            WorksheetPart worksheetPart = (WorksheetPart)documento.WorkbookPart.GetPartById(sheets.First().Id);
            if (sheets.Count() == 0)
                return null;
            else
                return worksheetPart.Worksheet;
        }

        #endregion

        #region GetColumnName

        private static string GetColumnName(string cellName)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        #endregion

        #endregion

    }

}

