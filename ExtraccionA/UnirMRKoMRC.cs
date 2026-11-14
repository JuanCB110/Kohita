using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExtraccionA
{
    /// <summary>
    /// Une varios archivos MRC o MRK en uno solo.
    /// La unión se hace directamente (sin cmarcedit -join) para no depender del límite
    /// de 32,767 caracteres de la línea de comandos de Windows, que se superaba con
    /// unos 700 archivos parteN.
    /// </summary>
    internal class UnirMRKoMRC
    {
        private const byte FinDeRegistro = 0x1D;

        public void UnirArchivosMRC(string carpetaEntrada, string archivoSalida)
        {
            List<string> archivos = ObtenerArchivos(carpetaEntrada, "*.mrc", archivoSalida);
            if (archivos == null) return;

            try
            {
                long registros = 0;
                using (FileStream salida = new FileStream(archivoSalida, FileMode.Create, FileAccess.Write))
                {
                    foreach (string archivo in archivos)
                    {
                        byte[] contenido = File.ReadAllBytes(archivo);
                        int inicio = TieneBomUtf8(contenido) ? 3 : 0;
                        salida.Write(contenido, inicio, contenido.Length - inicio);

                        for (int i = inicio; i < contenido.Length; i++)
                        {
                            if (contenido[i] == FinDeRegistro) registros++;
                        }
                    }
                }

                MostrarResumen(archivos.Count, registros, archivoSalida);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al unir los archivos MRC: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void UnirArchivosMRK(string carpetaEntrada, string archivoSalida)
        {
            List<string> archivos = ObtenerArchivos(carpetaEntrada, "*.mrk", archivoSalida);
            if (archivos == null) return;

            try
            {
                long registros = 0;
                using (StreamWriter salida = new StreamWriter(archivoSalida, false, new UTF8Encoding(false)))
                {
                    foreach (string archivo in archivos)
                    {
                        // ReadAllText detecta y quita el BOM si existe
                        string contenido = File.ReadAllText(archivo, Encoding.UTF8).TrimEnd('\r', '\n', ' ');
                        if (contenido.Length == 0) continue;

                        salida.Write(contenido);
                        salida.WriteLine();
                        salida.WriteLine(); // línea en blanco entre registros de archivos distintos

                        registros += Regex.Matches(contenido, @"^=LDR", RegexOptions.Multiline).Count;
                    }
                }

                MostrarResumen(archivos.Count, registros, archivoSalida);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrió un error al unir los archivos MRK: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Devuelve los archivos a unir en orden natural (parte2 antes que parte10),
        /// excluyendo el archivo de salida, o null si no se puede continuar
        /// </summary>
        private List<string> ObtenerArchivos(string carpetaEntrada, string patron, string archivoSalida)
        {
            if (!Directory.Exists(carpetaEntrada))
            {
                MessageBox.Show("La carpeta seleccionada no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            string salidaCompleta = Path.GetFullPath(archivoSalida);
            List<string> archivos = Directory.GetFiles(carpetaEntrada, patron)
                .Where(a => !string.Equals(Path.GetFullPath(a), salidaCompleta, StringComparison.OrdinalIgnoreCase))
                .OrderBy(a => ClaveOrden(a).Item1, StringComparer.OrdinalIgnoreCase)
                .ThenBy(a => ClaveOrden(a).Item2)
                .ThenBy(a => a, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (archivos.Count < 2)
            {
                MessageBox.Show("La carpeta debe contener al menos dos archivos para unirlos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

            return archivos;
        }

        /// <summary>
        /// Separa el nombre en texto y número final: "parte12" → ("parte", 12)
        /// </summary>
        private static Tuple<string, long> ClaveOrden(string ruta)
        {
            string nombre = Path.GetFileNameWithoutExtension(ruta);
            Match m = Regex.Match(nombre, @"^(.*?)(\d+)$");
            if (m.Success && long.TryParse(m.Groups[2].Value, out long numero))
            {
                return Tuple.Create(m.Groups[1].Value, numero);
            }
            return Tuple.Create(nombre, -1L);
        }

        private static bool TieneBomUtf8(byte[] datos)
        {
            return datos.Length >= 3 && datos[0] == 0xEF && datos[1] == 0xBB && datos[2] == 0xBF;
        }

        private static void MostrarResumen(int archivos, long registros, string archivoSalida)
        {
            MessageBox.Show(
                $"Unión completada.\n\nArchivos unidos: {archivos}\nRegistros: {registros}\nArchivo: {archivoSalida}",
                "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
