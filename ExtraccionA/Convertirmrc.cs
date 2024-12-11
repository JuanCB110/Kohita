using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExtraccionA
{
    internal class Convertirmrc
    {
        public string UbicacionArchivosMRK { private get; set; }
        public string UbicacionSalida { private get; set; }
        private List<string> Errores = new List<string>();

        public void ConvertirMRKaMRC()
        {
            string marceditPath = @"C:\Users\JuanCB\AppData\Roaming\MarcEdit 7.6 (User)\cmarcedit.exe"; // Ruta a cmarcedit.exe

            try
            {
                // Crear la carpeta de salida si no existe
                if (!Directory.Exists(UbicacionSalida))
                {
                    Directory.CreateDirectory(UbicacionSalida);
                }

                // Obtener todos los archivos MRK en la carpeta
                string[] mrkFiles = Directory.GetFiles(UbicacionArchivosMRK, "*.mrk");

                foreach (string mrkFile in mrkFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(mrkFile); // Nombre del archivo sin extensión
                    string outputFile = Path.Combine(UbicacionSalida, $"{fileName}.mrc"); // Archivo de salida MRC

                    // Crear el proceso
                    Process process = new Process();
                    process.StartInfo.FileName = marceditPath;
                    process.StartInfo.Arguments = $"-s \"{mrkFile}\" -d \"{outputFile}\" -make";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.RedirectStandardError = true;
                    /*process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;*/ // Ocultar la ventana

                    // Ejecutar MarcEdit
                    Console.WriteLine($"Convirtiendo: {mrkFile} -> {outputFile}");
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    // Verificar resultados
                    if (!string.IsNullOrEmpty(error))
                    {
                        Console.WriteLine($"Error al convertir {mrkFile}: {error}");
                        // Agregar el error a la lista
                        Errores.Add($"Error al convertir {mrkFile}: {error}");
                    }
                    else
                    {
                        Console.WriteLine($"Conversión completada: {outputFile}");
                    }
                }
                MessageBox.Show("Conversión de archivos completada.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al procesar archivos: {ex.Message}");
            }
            // Verificar si hay errores
            if (Errores.Count > 0)
            {
                // Mostrar todos los errores concatenados
                MessageBox.Show(string.Join("\n", Errores));
            }
        }
    }
}
