using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Globalization;

using System.IO;
using System.Diagnostics;

namespace Softech.Administrativo.Extraccion
{

    public static class XLSLIS 
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
                query.Append("exec ARepArticuloConPrecioLista  ");
                //------------------------ Filtros
                query.Append("@sCo_FechaMRLL_d=@f_fecha_i, ");
                query.Append("@sCo_FechaMRLL_h=@f_fecha_f, ");
                query.Append("@sCo_Art_d=@f_filtro2_i, ");
                query.Append("@sCo_Art_h=@f_filtro2_f, ");
                query.Append("@sCo_Linea_d=@f_filtro3_i, ");
                query.Append("@sCo_Linea_h=@f_filtro3_f, ");
                query.Append("@sCo_SubLinea_d=@f_filtro4_i, ");
                query.Append("@sCo_SubLinea_h=@f_filtro4_f, ");
                query.Append("@sCo_Categoria_d=@f_filtro5_i, ");
                query.Append("@sCo_Categoria_h=@f_filtro5_f, ");
                //------------------------ Parámetros
                query.Append("@sCo_Almacen1=@sCo_Almacen1, ");
                query.Append("@sCo_Almacen2=@sCo_Almacen2, ");
                query.Append("@sCo_Precio01=@sCo_Precio01, ");
                query.Append("@sCo_Color=@sCo_Color, ");
                query.Append("@sCo_NivelStock=@sCo_NivelStock ");

                SqlCommand comando = new SqlCommand(query.ToString(), conexion);

                #endregion

                #region Crear la coleccion de parametros

                #region Instanciar los parametros

                //------------------------ Filtros
                comando.Parameters.Add(new SqlParameter("@f_fecha_i", SqlDbType.DateTime));
                comando.Parameters.Add(new SqlParameter("@f_fecha_f", SqlDbType.DateTime));
                comando.Parameters.Add(new SqlParameter("@f_filtro2_i", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro2_f", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro3_i", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro3_f", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro4_i", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro4_f", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro5_i", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@f_filtro5_f", SqlDbType.Char));
                //------------------------ Parámetros
                comando.Parameters.Add(new SqlParameter("@sCo_Almacen1", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@sCo_Almacen2", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@sCo_Precio01", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@sCo_Color", SqlDbType.Char));
                comando.Parameters.Add(new SqlParameter("@sCo_NivelStock", SqlDbType.Char));

                //------------------------ Filtros
                comando.Parameters["@f_fecha_i"].IsNullable = true;
                comando.Parameters["@f_fecha_i"].SqlValue = DBNull.Value;

                comando.Parameters["@f_fecha_f"].IsNullable = true;
                comando.Parameters["@f_fecha_f"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro2_i"].IsNullable = true;
                comando.Parameters["@f_filtro2_i"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro2_f"].IsNullable = true;
                comando.Parameters["@f_filtro2_f"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro3_i"].IsNullable = true;
                comando.Parameters["@f_filtro3_i"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro3_f"].IsNullable = true;
                comando.Parameters["@f_filtro3_f"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro4_i"].IsNullable = true;
                comando.Parameters["@f_filtro4_i"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro4_f"].IsNullable = true;
                comando.Parameters["@f_filtro4_f"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro5_i"].IsNullable = true;
                comando.Parameters["@f_filtro5_i"].SqlValue = DBNull.Value;

                comando.Parameters["@f_filtro5_f"].IsNullable = true;
                comando.Parameters["@f_filtro5_f"].SqlValue = DBNull.Value;
                //------------------------ Parámetros
                comando.Parameters["@sCo_Almacen1"].IsNullable = true;
                comando.Parameters["@sCo_Almacen1"].SqlValue = DBNull.Value;

                comando.Parameters["@sCo_Almacen2"].IsNullable = true;
                comando.Parameters["@sCo_Almacen2"].SqlValue = DBNull.Value;

                comando.Parameters["@sCo_Precio01"].IsNullable = true;
                comando.Parameters["@sCo_Precio01"].SqlValue = DBNull.Value;

                comando.Parameters["@sCo_Color"].IsNullable = true;
                comando.Parameters["@sCo_Color"].SqlValue = DBNull.Value;

                comando.Parameters["@sCo_NivelStock"].IsNullable = true;
                comando.Parameters["@sCo_NivelStock"].SqlValue = DBNull.Value;

                #endregion

                #region Cargar los valores a los parametros

                //var strFechaPatron = "dd/MM/yyyy";
                //var strSeparadorFecha = "/";
                //CultureInfo innerCulture = new CultureInfo(CultureInfo.CurrentCulture.LCID);
                //innerCulture.DateTimeFormat.ShortDatePattern = strFechaPatron;
                //innerCulture.DateTimeFormat.DateSeparator = strSeparadorFecha;

                //------------------------ Filtros
                if (filtros.ContainsKey("fechaDesde"))
                    comando.Parameters["@f_fecha_i"].SqlValue = (DateTime)filtros["fechaDesde"];

                if (filtros.ContainsKey("fechaHasta"))
                    comando.Parameters["@f_fecha_f"].SqlValue = (DateTime)filtros["fechaHasta"];

                if (filtros.ContainsKey("f_filtro2_i") && !Equals((String)filtros["f_filtro2_i"], null))
                    comando.Parameters["@f_filtro2_i"].SqlValue = filtros["f_filtro2_i"];

                if (filtros.ContainsKey("f_filtro2_f") && !Equals((String)filtros["f_filtro2_f"], null))
                    comando.Parameters["@f_filtro2_f"].SqlValue = filtros["f_filtro2_f"];

                if (filtros.ContainsKey("f_filtro3_i") && !Equals((String)filtros["f_filtro3_i"], null))
                    comando.Parameters["@f_filtro3_i"].SqlValue = filtros["f_filtro3_i"];

                if (filtros.ContainsKey("f_filtro3_f") && !Equals((String)filtros["f_filtro3_f"], null))
                    comando.Parameters["@f_filtro3_f"].SqlValue = filtros["f_filtro3_f"];

                if (filtros.ContainsKey("f_filtro4_i") && !Equals((String)filtros["f_filtro4_i"], null))
                    comando.Parameters["@f_filtro4_i"].SqlValue = filtros["f_filtro4_i"];

                if (filtros.ContainsKey("f_filtro4_f") && !Equals((String)filtros["f_filtro4_f"], null))
                    comando.Parameters["@f_filtro4_f"].SqlValue = filtros["f_filtro4_f"];

                if (filtros.ContainsKey("f_filtro5_i") && !Equals((String)filtros["f_filtro5_i"], null))
                    comando.Parameters["@f_filtro5_i"].SqlValue = filtros["f_filtro5_i"];

                if (filtros.ContainsKey("f_filtro5_f") && !Equals((String)filtros["f_filtro5_f"], null))
                    comando.Parameters["@f_filtro5_f"].SqlValue = filtros["f_filtro5_f"];
                //------------------------ Parámetros
                if (filtros.ContainsKey("sCo_Almacen1") && !Equals((String)filtros["sCo_Almacen1"], null))
                    comando.Parameters["@sCo_Almacen1"].SqlValue = filtros["sCo_Almacen1"];

                if (filtros.ContainsKey("sCo_Almacen2") && !Equals((String)filtros["sCo_Almacen2"], null))
                    comando.Parameters["@sCo_Almacen2"].SqlValue = filtros["sCo_Almacen2"];

                if (filtros.ContainsKey("sCo_Precio01") && !Equals((String)filtros["sCo_Precio01"], null))
                    comando.Parameters["@sCo_Precio01"].SqlValue = filtros["sCo_Precio01"];

                if (filtros.ContainsKey("sCo_Color") && !Equals((String)filtros["sCo_Color"], null))
                    comando.Parameters["@sCo_Color"].SqlValue = filtros["sCo_Color"];

                if (filtros.ContainsKey("sCo_NivelStock") && !Equals((String)filtros["sCo_NivelStock"], null))
                    comando.Parameters["@sCo_NivelStock"].SqlValue = filtros["sCo_NivelStock"];

                #endregion

                #endregion

                #region Ejecutar query ExecuteNonQuery


                SqlDataAdapter da = new SqlDataAdapter(comando);
                da.Fill(extraccionDataSet, "XLSLISTA");

                // Contar los registros totales
                int registros = ContarRegistros(extraccionDataSet);

                // Agregar los registros al extractorDataSet
                AgregarRegistrosMetadata(extraccionDataSet, registros);

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

        private static int ContarRegistros(DataSet dataSet)
        {
            if (dataSet.Tables.Contains("XLSLISTA"))
            {
                int count = dataSet.Tables["XLSLISTA"].Rows.Count;
                return count;
            }
            return 0;
        }

        private static void AgregarRegistrosMetadata(DataSet dataSet, int registros)
        {
            // Crear una tabla para almacenar el número de registros
            DataTable metaDataTable = new DataTable("MetaData");
            metaDataTable.Columns.Add("Registros", typeof(int));

            // Crear una fila con el conteo
            DataRow row = metaDataTable.NewRow();
            row["Registros"] = registros;
            metaDataTable.Rows.Add(row);

            // Añadir la tabla al DataSet
            dataSet.Tables.Add(metaDataTable);
        }

    }
}
