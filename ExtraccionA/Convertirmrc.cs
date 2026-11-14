using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace ExtraccionA
{
    /// <summary>
    /// Avance de la conversión que se envía a la ventana de progreso
    /// </summary>
    internal class ProgresoConversion
    {
        public int Actual { get; set; }
        public int Total { get; set; }
        public string Archivo { get; set; }
    }

    /// <summary>
    /// Resultado de convertir una carpeta de archivos MRK a MRC
    /// </summary>
    internal class ResultadoConversion
    {
        public int Total { get; set; }
        public int Convertidos { get; set; }
        public List<KeyValuePair<string, string>> Fallidos { get; } = new List<KeyValuePair<string, string>>();
        public bool Cancelado { get; set; }
        public bool DetenidoEnPrimerArchivo { get; set; }
        public string ArchivoErrores { get; set; }
    }

    internal class Convertirmrc
    {
        public const string NombreArchivoErrores = "errores_conversion.txt";

        // Tiempo máximo por archivo antes de dar por colgado a cmarcedit
        private static readonly TimeSpan TiempoMaximoPorArchivo = TimeSpan.FromMinutes(5);

        public string UbicacionArchivosMRK { private get; set; }
        public string UbicacionSalida { private get; set; }

        /// <summary>
        /// Convierte cada .mrk de la carpeta a .mrc con cmarcedit. Pensado para correr fuera del hilo de la interfaz.
        /// Un archivo cuenta como convertido solo si cmarcedit genera el .mrc y no está vacío.
        /// Si falla el primer archivo se detiene, porque casi siempre significa que MarcEdit no puede ejecutarse.
        /// </summary>
        public ResultadoConversion ConvertirMRKaMRC(string marceditPath, IProgress<ProgresoConversion> progreso, CancellationToken cancelacion)
        {
            ResultadoConversion resultado = new ResultadoConversion();

            Directory.CreateDirectory(UbicacionSalida);

            List<string> mrkFiles = UnirMRKoMRC.OrdenarArchivos(Directory.GetFiles(UbicacionArchivosMRK, "*.mrk"));
            resultado.Total = mrkFiles.Count;

            // Detalle completo de cada fallo para errores_conversion.txt
            StringBuilder detalle = new StringBuilder();

            for (int i = 0; i < mrkFiles.Count; i++)
            {
                if (cancelacion.IsCancellationRequested)
                {
                    resultado.Cancelado = true;
                    break;
                }

                string mrkFile = mrkFiles[i];
                string nombre = Path.GetFileName(mrkFile);
                progreso?.Report(new ProgresoConversion { Actual = i + 1, Total = mrkFiles.Count, Archivo = nombre });

                string outputFile = Path.Combine(UbicacionSalida, Path.GetFileNameWithoutExtension(mrkFile) + ".mrc");

                string motivo;
                string textoCompleto;
                if (ConvertirArchivo(marceditPath, mrkFile, outputFile, out motivo, out textoCompleto))
                {
                    resultado.Convertidos++;
                    continue;
                }

                resultado.Fallidos.Add(new KeyValuePair<string, string>(nombre, motivo));
                detalle.AppendLine($"Archivo: {mrkFile}");
                detalle.AppendLine($"Motivo: {motivo}");
                if (!string.IsNullOrWhiteSpace(textoCompleto))
                {
                    detalle.AppendLine("Salida de cmarcedit:");
                    detalle.AppendLine(textoCompleto.Trim());
                }
                detalle.AppendLine();

                if (i == 0)
                {
                    resultado.DetenidoEnPrimerArchivo = true;
                    break;
                }
            }

            EscribirArchivoErrores(resultado, detalle, marceditPath);
            return resultado;
        }

        /// <summary>
        /// Ejecuta cmarcedit para un archivo. Devuelve true si se generó el .mrc.
        /// </summary>
        private static bool ConvertirArchivo(string marceditPath, string mrkFile, string outputFile, out string motivo, out string textoCompleto)
        {
            motivo = null;
            textoCompleto = null;

            try
            {
                // Se borra el .mrc anterior para no contar como convertido un archivo de otra corrida
                if (File.Exists(outputFile))
                    File.Delete(outputFile);

                using (Process process = new Process())
                {
                    process.StartInfo.FileName = marceditPath;
                    process.StartInfo.Arguments = $"-s \"{mrkFile}\" -d \"{outputFile}\" -make";
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.RedirectStandardError = true;
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

                    process.Start();

                    // Leer ambas salidas al mismo tiempo evita que cmarcedit se bloquee con un búfer lleno
                    var salida = process.StandardOutput.ReadToEndAsync();
                    var errores = process.StandardError.ReadToEndAsync();

                    if (!process.WaitForExit((int)TiempoMaximoPorArchivo.TotalMilliseconds))
                    {
                        try { process.Kill(); } catch { }
                        motivo = $"cmarcedit no terminó en {TiempoMaximoPorArchivo.TotalMinutes} minutos y se detuvo.";
                        return false;
                    }
                    process.WaitForExit();

                    string textoSalida = salida.Result;
                    string textoError = errores.Result;
                    textoCompleto = (textoError + Environment.NewLine + textoSalida).Trim();

                    FileInfo mrc = new FileInfo(outputFile);
                    if (mrc.Exists && mrc.Length > 0)
                        return true;

                    motivo = mrc.Exists && process.ExitCode == 0 && textoError.Trim().Length == 0
                        ? "cmarcedit terminó sin generar registros (el .mrc quedó vacío). Revise que el .mrk tenga un formato MARC válido."
                        : ExplicarFallo(process.ExitCode, textoError, textoSalida);
                    return false;
                }
            }
            catch (Exception ex)
            {
                motivo = "No se pudo ejecutar cmarcedit: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Convierte la salida de cmarcedit en un motivo corto y entendible
        /// </summary>
        private static string ExplicarFallo(int codigoSalida, string textoError, string textoSalida)
        {
            string todo = (textoError ?? "") + " " + (textoSalida ?? "");

            if (todo.IndexOf("0x800711C7", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Windows bloqueó MarcEdit (Control inteligente de aplicaciones, error 0x800711C7). " +
                       "MarcEdit no tiene firma digital y Windows 11 no le permite ejecutarse.";
            }

            string primeraLinea = todo
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(l => l.Trim())
                .FirstOrDefault(l => l.Length > 0);

            if (primeraLinea != null && primeraLinea.Length > 200)
                primeraLinea = primeraLinea.Substring(0, 200) + "…";

            return primeraLinea != null
                ? $"cmarcedit no generó el archivo (código {codigoSalida}): {primeraLinea}"
                : $"cmarcedit no generó el archivo (código {codigoSalida}) y no dio ningún mensaje.";
        }

        /// <summary>
        /// Escribe errores_conversion.txt en la carpeta de salida si hubo fallos; si no, borra el de una corrida anterior
        /// </summary>
        private void EscribirArchivoErrores(ResultadoConversion resultado, StringBuilder detalle, string marceditPath)
        {
            string ruta = Path.Combine(UbicacionSalida, NombreArchivoErrores);

            try
            {
                if (resultado.Fallidos.Count == 0)
                {
                    if (File.Exists(ruta))
                        File.Delete(ruta);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Errores de conversión MRK a MRC");
                sb.AppendLine($"Fecha: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"MarcEdit: {marceditPath}");
                sb.AppendLine($"Carpeta MRK: {UbicacionArchivosMRK}");
                sb.AppendLine($"Carpeta de salida: {UbicacionSalida}");
                sb.AppendLine($"Convertidos: {resultado.Convertidos} de {resultado.Total}. Fallidos: {resultado.Fallidos.Count}.");
                if (resultado.DetenidoEnPrimerArchivo)
                    sb.AppendLine("La conversión se detuvo porque falló el primer archivo.");
                if (resultado.Cancelado)
                    sb.AppendLine("La conversión fue cancelada por el usuario.");
                sb.AppendLine();
                sb.Append(detalle);

                File.WriteAllText(ruta, sb.ToString(), new UTF8Encoding(true));
                resultado.ArchivoErrores = ruta;
            }
            catch
            {
                // Si no se puede escribir el archivo de errores, el resumen en pantalla sigue mostrando los fallos
                resultado.ArchivoErrores = null;
            }
        }
    }
}
