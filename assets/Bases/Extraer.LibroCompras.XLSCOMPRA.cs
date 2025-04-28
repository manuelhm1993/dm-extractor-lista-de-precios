// INFORMACIÓN
//
// Nombre:
//      Extraccion de Datos para la ley de políca habitacional del Banco 
//      Mercantil
// Descripción:
//      Archivo de código para la extracción de los datos necesarios para 
//      la generacion del archivo de texto plano para la ley de política
//      del banco Mercantil.
// Empresa:
//      Softech Sistemas
// Autor:
//      Softech Sistemas
// Fecha:
//      12 de Septiembre de 2.008

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Globalization;

namespace Softech.Administrativo.Extraccion
{
    /// <summary>
    /// Clase para extraer datos del sistema
    /// </summary>
    public static class XLSCOM
    {
        /// <summary>
        /// Extrae un DataSet con la informacion necesaria para el proceso de extracción de datos
        /// </summary>
        /// <param name="conexion">Objeto de conexión sql con la conexion abierta</param>
        /// <param name="filtros">Diccionario con los filtros y valores para la extracción de datos</param>
        /// <param name="parametros">Parametros necesarios para la extracción de datos</param>
        /// <returns>Set de datos con la informacion de la extracción</returns>
        public static Object Extraer(SqlConnection conexion,
            Dictionary<String, Object> filtros,
            Dictionary<String, Object> parametros,
            String masterProfit)
        {
            DataSet extraccionDataSet = new DataSet("Extraccion");

            #region Comprobar el estado de la conexion

            if (conexion.State != ConnectionState.Open)
                throw new ArgumentException("Para poder extraer los datos es necesario que la conexion esté abierta.");

            #endregion

            using (conexion)
            {
                #region Armar query

                StringBuilder query = new StringBuilder();
                query.Append("exec RepLibroCompra  ");
                query.Append("@sCo_fecha_d=@f_fecha_i, ");
                query.Append("@sCo_fecha_h=@f_fecha_f, ");
                query.Append("@cCo_Sucursal=@f_filtro2, ");
                query.Append("@bIncluirOrden=default, ");
                query.Append("@bImprimirColumnImport=@f_filtro3, ");
                query.Append("@bImprimirColumnArt33=@f_filtro4, ");
                query.Append("@bImprimirColumnArt34=@f_filtro5");

                SqlCommand comando = new SqlCommand(query.ToString(), conexion);

                #endregion

                #region Crear la coleccion de parametros

                #region Instanciar los parametros


                comando.Parameters.Add(new SqlParameter("@f_fecha_i", SqlDbType.DateTime));
                comando.Parameters.Add(new SqlParameter("@f_fecha_f", SqlDbType.DateTime));
                comando.Parameters.Add(new SqlParameter("@f_filtro2", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro3", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro4", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro5", SqlDbType.Char));

                comando.Parameters["@f_fecha_i"].IsNullable = true;
                comando.Parameters["@f_fecha_i"].SqlValue = DBNull.Value;

                comando.Parameters["@f_fecha_f"].IsNullable = true;
                comando.Parameters["@f_fecha_f"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro2"].IsNullable = true;
                comando.Parameters["@f_filtro2"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro3"].IsNullable = true;
                comando.Parameters["@f_filtro3"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro4"].IsNullable = true;
                comando.Parameters["@f_filtro4"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro5"].IsNullable = true;
                comando.Parameters["@f_filtro5"].SqlValue = DBNull.Value;

                #endregion

                #region Cargar los valores a los parametros

                //var strFechaPatron = "dd/MM/yyyy";
                //var strSeparadorFecha = "/";
                //CultureInfo innerCulture = new CultureInfo(CultureInfo.CurrentCulture.LCID);
                //innerCulture.DateTimeFormat.ShortDatePattern = strFechaPatron;
                //innerCulture.DateTimeFormat.DateSeparator = strSeparadorFecha;

                if (filtros.ContainsKey("fechaDesde"))
                    comando.Parameters["@f_fecha_i"].SqlValue = (DateTime)filtros["fechaDesde"];

                if (filtros.ContainsKey("fechaHasta"))
                    comando.Parameters["@f_fecha_f"].SqlValue = (DateTime)filtros["fechaHasta"];

                if (filtros.ContainsKey("sucursal"))// && !Equals(filtros["f_filtro2"], null))
                    if (String.IsNullOrEmpty((String)filtros["sucursal"]))
                        comando.Parameters["@f_filtro2"].SqlValue = DBNull.Value;
                    else
                        comando.Parameters["@f_filtro2"].SqlValue = filtros["sucursal"];

                if (filtros.ContainsKey("f_filtro3") && !Equals((String)filtros["f_filtro3"], null))
                    comando.Parameters["@f_filtro3"].SqlValue = filtros["f_filtro3"];

                if (filtros.ContainsKey("f_filtro4") && !Equals((String)filtros["f_filtro4"], null))
                    comando.Parameters["@f_filtro4"].SqlValue = filtros["f_filtro4"];

                if (filtros.ContainsKey("f_filtro5") && !Equals((String)filtros["f_filtro5"], null))
                    comando.Parameters["@f_filtro5"].SqlValue = filtros["f_filtro5"];


                #endregion

                #endregion

                #region Ejecutar query ExecuteNonQuery


                SqlDataAdapter da = new SqlDataAdapter(comando);
                da.Fill(extraccionDataSet, "XLSCOMPRA");
                for (int i = 0; i < extraccionDataSet.Tables.Count; i++)
                {
                    for (int j = 0; j < extraccionDataSet.Tables[i].Columns.Count; j++)
                        if (extraccionDataSet.Tables[i].Columns[j].DataType == System.Type.GetType("System.DateTime"))
                            extraccionDataSet.Tables[i].Columns[j].DateTimeMode = DataSetDateTime.Unspecified;
                }
                #endregion
            }

            return extraccionDataSet;
        }
    }
}