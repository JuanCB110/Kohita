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
            string marceditPath = @"C:\Users\JuanCB\AppData\Roaming\MarcEdit 7.6 (User)\cmarcedit.exe"; // Ruta a cmarcedit.exe

            // Verificar si la ruta de MarcEdit existe
            if (!File.Exists(marceditPath))
            {
                MessageBox.Show("No se encontró MarcEdit en la ruta especificada.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(carpetaEntrada))
            {
                MessageBox.Show("La carpeta seleccionada no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Obtener archivos .mrc de la carpeta
            string[] archivosMrc = Directory.GetFiles(carpetaEntrada, "*.mrc");

            if (archivosMrc.Length < 2)
            {
                MessageBox.Show("La carpeta debe contener al menos dos archivos MRC para unirlos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Construir la lista de archivos separados por ";"
            string archivosEntrada = string.Join(";", archivosMrc);

            // Mensaje de depuración
            //MessageBox.Show($"Archivos de entrada: {archivosEntrada}");

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = marceditPath,
                    Arguments = $"-join -s \"{archivosEntrada}\" -d \"{archivoSalida}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process proceso = new Process { StartInfo = psi })
                {
                    proceso.Start();
                    string salida = proceso.StandardOutput.ReadToEnd();
                    string error = proceso.StandardError.ReadToEnd();
                    proceso.WaitForExit();

                    if (!string.IsNullOrWhiteSpace(salida))
                        MessageBox.Show("Proceso finalizado correctamente: \n" + salida, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (!string.IsNullOrWhiteSpace(error))
                        MessageBox.Show("Error durante la ejecución: \n" + error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al ejecutar MarcEdit: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void UnirArchivosMRK(string carpetaEntrada, string archivoSalida)
        {
            string marceditPath = @"C:\Users\JuanCB\AppData\Roaming\MarcEdit 7.6 (User)\cmarcedit.exe"; // Ruta a cmarcedit.exe

            // Verificar si la ruta de MarcEdit existe
            if (!File.Exists(marceditPath))
            {
                MessageBox.Show("No se encontró MarcEdit en la ruta especificada.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(carpetaEntrada))
            {
                MessageBox.Show("La carpeta seleccionada no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Obtener archivos .mrk de la carpeta
            string[] archivosMrk = Directory.GetFiles(carpetaEntrada, "*.mrk");

            if (archivosMrk.Length < 2)
            {
                MessageBox.Show("La carpeta debe contener al menos dos archivos MRC para unirlos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Construir la lista de archivos separados por ";"
            string archivosEntrada = string.Join(";", archivosMrk);

            // Mensaje de depuración
            //MessageBox.Show($"Archivos de entrada: {archivosEntrada}");

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = marceditPath,
                    Arguments = $"-join -s \"{archivosEntrada}\" -d \"{archivoSalida}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process proceso = new Process { StartInfo = psi })
                {
                    proceso.Start();
                    string salida = proceso.StandardOutput.ReadToEnd();
                    string error = proceso.StandardError.ReadToEnd();
                    proceso.WaitForExit();

                    if (!string.IsNullOrWhiteSpace(salida))
                        MessageBox.Show("Proceso finalizado correctamente: \n" + salida, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (!string.IsNullOrWhiteSpace(error))
                        MessageBox.Show("Error durante la ejecución: \n" + error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al ejecutar MarcEdit: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
