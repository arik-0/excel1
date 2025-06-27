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

        Console.Write("¿Qué deseas hacer? Leer (L) o Editar (E): ");
        string modo = Console.ReadLine().Trim().ToUpper();

        if (modo != "L" && modo != "E")
        {
            Console.WriteLine("❌ Opción inválida.");
            return;
        }

        // Mostrar archivos disponibles
        for (int i = 0; i < archivosRelativos.Length; i++)
        {
            Console.WriteLine($"{i + 1}. {archivosRelativos[i]}");
        }

        Console.Write("\nSeleccione el número del archivo: ");
        int seleccion = int.Parse(Console.ReadLine());

        if (seleccion < 1 || seleccion > archivosRelativos.Length)
        {
            Console.WriteLine("❌ Selección inválida.");
            return;
        }

        string rutaCompleta = Path.Combine(baseDir, archivosRelativos[seleccion - 1]);

        if (!File.Exists(rutaCompleta))
        {
            Console.WriteLine($"❌ No se encontró el archivo: {rutaCompleta}");
            return;
        }

        using (var package = new ExcelPackage(new FileInfo(rutaCompleta)))
        {
            var workbook = package.Workbook;

            // Mostrar hojas
            Console.WriteLine("\n📑 Hojas disponibles:");
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {workbook.Worksheets[i].Name}");
            }

            Console.Write("\nSeleccione el número de la hoja: ");
            int hojaSeleccionada = int.Parse(Console.ReadLine());

            if (hojaSeleccionada < 1 || hojaSeleccionada > workbook.Worksheets.Count)
            {
                Console.WriteLine("❌ Hoja inválida.");
                return;
            }

            var hoja = workbook.Worksheets[hojaSeleccionada - 1];

            if (hoja.Dimension == null)
            {
                Console.WriteLine("❌ La hoja está vacía.");
                return;
            }

            int filaInicio = hoja.Dimension.Start.Row;
            int filaFin = hoja.Dimension.End.Row;
            int colInicio = hoja.Dimension.Start.Column;
            int colFin = hoja.Dimension.End.Column;
            int anchoColumna = 20;

            if (modo == "L")
            {
                Console.WriteLine($"\n📄 Contenido de la hoja: {hoja.Name}\n");

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

                Console.WriteLine("\n✅ Lectura finalizada.");
            }
            else if (modo == "E")
            {
                Console.WriteLine($"\n📄 Contenido actual (primera parte): {hoja.Name}\n");

                // Mostrar las primeras 10 filas como referencia
                for (int fila = filaInicio; fila <= Math.Min(filaInicio + 9, filaFin); fila++)
                {
                    for (int col = colInicio; col <= colFin; col++)
                    {
                        var valor = hoja.Cells[fila, col].Value;
                        string texto = valor == null ? "" : valor.ToString();
                        Console.Write(texto.PadRight(anchoColumna));
                    }
                    Console.WriteLine();
                }

                Console.Write("\nIngrese número de fila a editar: ");
                int filaEdit = int.Parse(Console.ReadLine());

                Console.Write("Ingrese número de columna a editar: ");
                int colEdit = int.Parse(Console.ReadLine());

                Console.Write("Ingrese nuevo valor: ");
                string nuevoValor = Console.ReadLine();

                hoja.Cells[filaEdit, colEdit].Value = nuevoValor;

                package.Save();

                Console.WriteLine("✅ Celda actualizada y archivo guardado.");
            }
        }

        Console.ReadKey();
    }
}
