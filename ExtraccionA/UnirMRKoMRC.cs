using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtraccionA
{
    internal class UnirMRKoMRC
    {
        public void UnirArchivosMRC(string carpetaEntrada, string archivoSalida)
        {
            string marceditPath = @"C:\\Users\\JuanCB\\AppData\\Roaming\\MarcEdit 7.6 (User)\\cmarcedit.exe"; // Ruta a cmarcedit.exe

            if (!Directory.Exists(carpetaEntrada))
            {
                MessageBox.Show("La carpeta seleccionada no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verificar que la carpeta contiene archivos .mrc
            string[] archivosMrc = Directory.GetFiles(carpetaEntrada, "*.mrc");

            if (archivosMrc.Length < 2)
            {
                MessageBox.Show("La carpeta debe contener al menos dos archivos MRC para unirlos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Generar una lista separada por espacios para los archivos de entrada
            string archivosEntrada = string.Join(";", archivosMrc.Select(a => $"\"{a}\""));

            Process process = new Process();
            process.StartInfo.FileName = marceditPath; // Ruta a cmarcedit.exe
            process.StartInfo.Arguments = $"-join -s \"{archivosEntrada}\" -d \"{archivoSalida}\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;

            try
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode == 0)
                {
                    MessageBox.Show("Archivos unidos exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Error al unir archivos: {error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void UnirArchivosMRK(string carpetaEntrada, string archivoSalida)
        {
            string marceditPath = @"C:\\Users\\JuanCB\\AppData\\Roaming\\MarcEdit 7.6 (User)\\cmarcedit.exe"; // Ruta a cmarcedit.exe

            if (!Directory.Exists(carpetaEntrada))
            {
                MessageBox.Show("La carpeta seleccionada no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Verificar que la carpeta contiene archivos .mrc
            string[] archivosMrc = Directory.GetFiles(carpetaEntrada, "*.mrk");

            if (archivosMrc.Length < 2)
            {
                MessageBox.Show("La carpeta debe contener al menos dos archivos MRK para unirlos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Generar una lista separada por espacios para los archivos de entrada
            string archivosEntrada = string.Join(";", archivosMrc.Select(a => $"\"{a}\""));

            Process process = new Process();
            process.StartInfo.FileName = marceditPath; // Ruta a cmarcedit.exe
            process.StartInfo.Arguments = $"-join -s \"{archivosEntrada}\" -d \"{archivoSalida}\"";
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;

            try
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode == 0)
                {
                    MessageBox.Show("Archivos unidos exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Error al unir archivos: {error}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
