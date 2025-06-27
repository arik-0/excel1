using OfficeOpenXml;
using System;
using System.IO;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        string[] archivosRelativos = new string[]
        {
            @"excels\AsientosContables_FrescaTech.xlsx",
            @"excels\Balance_Sumas_y_Saldos_FrescaTech.xlsx",
            @"excels\AsientosContables_FrescaTech.xlsx"
        };

        foreach (var archivoRel in archivosRelativos)
        {
            string rutaCompleta = Path.Combine(baseDir, archivoRel);

            if (!File.Exists(rutaCompleta))
            {
                Console.WriteLine($"❌ No se encontró el archivo: {rutaCompleta}\n");
                continue;
            }

            Console.WriteLine($"📄 Leyendo archivo: {rutaCompleta} ");

            using (var package = new ExcelPackage(new FileInfo(rutaCompleta)))
            {
                foreach (var hoja in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"📑 Hoja: {hoja.Name}");

                    if (hoja.Dimension == null)
                    {
                        Console.WriteLine("La hoja está vacía.\n");
                        continue;
                    }

                    int filaInicio = hoja.Dimension.Start.Row;
                    int filaFin = hoja.Dimension.End.Row;

                    int colInicio = hoja.Dimension.Start.Column;
                    int colFin = hoja.Dimension.End.Column;

                    int anchoColumna = 20; // ancho fijo para cada columna

                    for (int fila = filaInicio; fila <= filaFin; fila++)
                    {
                        for (int col = colInicio; col <= colFin; col++)
                        {
                            var valor = hoja.Cells[fila, col].Value;
                            string texto = valor == null ? "" : valor.ToString();
                            Console.Write(texto.PadRight(anchoColumna));
                        }
                        Console.WriteLine();
                    }

                    Console.WriteLine("\n-----------------------------\n");
                }
            }
        }

        Console.WriteLine("✅ Lectura finalizada.");
        Console.ReadKey();
    }
}
