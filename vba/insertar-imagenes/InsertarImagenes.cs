public static void InsertarImagenesEnHoja(string rutaArchivoExcel, string carpetaImagenes)
{
    Excel.Application excelApp = new Excel.Application();
    Excel.Workbook workbook = excelApp.Workbooks.Open(rutaArchivoExcel, ReadOnly: false);
    Excel._Worksheet hojaImagen = null;

    try
    {
        hojaImagen = workbook.Sheets["IMAGEN"];

        // Validar si ya hay contenido (columna A fila 2)
        if (hojaImagen.Cells[2, 1].Value != null)
        {
            Console.WriteLine("Las imágenes ya están insertadas. No se realiza ninguna acción.");
            return;
        }

        var archivosBmp = Directory.GetFiles(carpetaImagenes, "*.bmp");
        int fila = 2;

        foreach (var imagenPath in archivosBmp)
        {
            if (!File.Exists(imagenPath))
                continue;

            string codigo = Path.GetFileNameWithoutExtension(imagenPath);

            hojaImagen.Cells[fila, 1].Value = codigo;

            Excel.Range celda = hojaImagen.Cells[fila, 2];
            float left = (float)(double)celda.Left;
            float top = (float)(double)celda.Top;

            hojaImagen.Shapes.AddPicture(
                imagenPath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoCTrue,
                left, top, 100, 100
            );

            fila++;
        }

        workbook.Save();
    }
    catch (Exception ex)
    {
        Console.WriteLine("Error al insertar imágenes: " + ex.Message);
    }
    finally
    {
        workbook.Close(SaveChanges: false);
        excelApp.Quit();

        Marshal.ReleaseComObject(hojaImagen);
        Marshal.ReleaseComObject(workbook);
        Marshal.ReleaseComObject(excelApp);
    }
}
