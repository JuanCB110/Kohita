using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio encargado de la exportación de archivos en formato MARC
    /// </summary>
    public class FileExportService
    {
        private readonly MarcFormatterService _formatter;
        private const int REGISTROS_POR_ARCHIVO = 20;

        public FileExportService()
        {
            _formatter = new MarcFormatterService();
        }

        /// <summary>
        /// Exporta filas seleccionadas del DataGridView a archivos .mrk
        /// </summary>
        public int ExportarFilasAMRK(DataGridView dataGridView, string outputFolder, string codigoBiblioteca)
        {
            if (dataGridView.SelectedRows.Count == 0)
            {
                throw new InvalidOperationException("No hay filas seleccionadas para exportar.");
            }

            if (string.IsNullOrEmpty(outputFolder))
            {
                throw new InvalidOperationException("No se ha especificado la carpeta de salida.");
            }

            StringBuilder formattedText = new StringBuilder();
            int registrosProcesados = 0;
            int archivoNumero = 0;
            HashSet<string> processedControlNumbers = new HashSet<string>();
            List<string> columnasExcluidas = new List<string> { "Numeros_de_Items" };

            // SelectedRows viene en orden inverso al de selección (con Sel. Todas, la última ficha primero);
            // se exporta en el orden en que se ven en la tabla
            IEnumerable<DataGridViewRow> filasEnOrden = dataGridView.SelectedRows
                .Cast<DataGridViewRow>()
                .OrderBy(r => r.Index);

            foreach (DataGridViewRow selectedRow in filasEnOrden)
            {
                string controlNumber = selectedRow.Cells["003"].Value?.ToString();
                string nitems = selectedRow.Cells["Numeros_de_Items"].Value?.ToString();

                if (!string.IsNullOrEmpty(controlNumber))
                {
                    processedControlNumbers.Add(controlNumber);

                    // Agregar líder y campos de control
                    formattedText.AppendLine("=LDR  00000nam a2200000 a 4500");
                    formattedText.AppendLine($"=003  {controlNumber}");
                    formattedText.AppendLine($"=005  {selectedRow.Cells["005"].Value}");
                    formattedText.AppendLine($"=008  {selectedRow.Cells["008"].Value}");

                    // Procesar cada celda de la fila
                    foreach (DataGridViewCell cell in selectedRow.Cells)
                    {
                        string nombreColumna = dataGridView.Columns[cell.ColumnIndex].HeaderText;

                        if (columnasExcluidas.Contains(nombreColumna) || 
                            cell.Value == null || 
                            string.IsNullOrEmpty(cell.Value.ToString().Trim()))
                        {
                            continue;
                        }

                        string lineaFormateada = FormatearCelda(nombreColumna, cell.Value.ToString());
                        if (!string.IsNullOrEmpty(lineaFormateada))
                        {
                            formattedText.Append(lineaFormateada);
                        }
                    }

                    // Agregar items si existen
                    if (!string.IsNullOrEmpty(nitems))
                    {
                        string[] items = nitems.Split(',');
                        foreach (string item in items)
                        {
                            formattedText.AppendLine($"=952  $p{item.Trim()}$b{codigoBiblioteca}$o{selectedRow.Cells["050"].Value}");
                        }
                    }

                    formattedText.AppendLine(); // Línea en blanco entre registros
                    registrosProcesados++;

                    // Guardar archivo cada 20 registros
                    if (registrosProcesados >= REGISTROS_POR_ARCHIVO)
                    {
                        GuardarArchivo(outputFolder, archivoNumero, formattedText.ToString());
                        formattedText.Clear();
                        archivoNumero++;
                        registrosProcesados = 0;
                    }
                }
            }

            // Guardar registros restantes
            if (registrosProcesados > 0)
            {
                GuardarArchivo(outputFolder, archivoNumero, formattedText.ToString());
                archivoNumero++;
            }

            return archivoNumero;
        }

        /// <summary>
        /// Formatea una celda según su etiqueta MARC
        /// </summary>
        private string FormatearCelda(string etiqueta, string valor)
        {
            switch (etiqueta)
            {
                case "003":
                case "005":
                case "008":
                case "Numeros_de_Items":
                    return string.Empty; // Ya fueron procesados

                case "020":
                    return _formatter.FormatearEtiqueta020(valor) + Environment.NewLine;

                case "040":
                    return _formatter.FormatearEtiqueta040(valor) + Environment.NewLine;

                case "050":
                    return _formatter.FormatearEtiqueta050(valor) + Environment.NewLine;

                case "100":
                case "110":
                case "130":
                    return _formatter.FormatearEtiquetaAutor(etiqueta, valor) + Environment.NewLine;

                case "240":
                    return _formatter.FormatearEtiqueta240(valor) + Environment.NewLine;

                case "245":
                    return _formatter.FormatearEtiqueta245(valor) + Environment.NewLine;

                case "250":
                    return _formatter.FormatearEtiqueta250(valor) + Environment.NewLine;

                case "260":
                    return _formatter.FormatearEtiqueta260(valor) + Environment.NewLine;

                case "300":
                    return _formatter.FormatearEtiqueta300(valor) + Environment.NewLine;

                case "440":
                case "490":
                    return _formatter.FormatearEtiqueta490(valor) + Environment.NewLine;

                case "500":
                    return _formatter.FormatearEtiqueta500(valor) + Environment.NewLine;

                case "502":
                case "503":
                case "504":
                    return _formatter.FormatearEtiqueta504(valor) + Environment.NewLine;

                case "505":
                    return _formatter.FormatearEtiqueta505(valor) + Environment.NewLine;

                case "600":
                case "610":
                case "650":
                case "651":
                    return _formatter.FormatearEtiquetaMateria(etiqueta, valor) + Environment.NewLine;

                case "700":
                    return _formatter.FormatearEtiqueta700(valor) + Environment.NewLine;

                case "942":
                    return $"=942  $c{valor}" + Environment.NewLine;

                default:
                    // Etiqueta genérica
                    if (etiqueta.Length == 3 && int.TryParse(etiqueta, out _))
                    {
                        return _formatter.FormatearEtiquetaGenerica(etiqueta, valor) + Environment.NewLine;
                    }
                    return string.Empty;
            }
        }

        /// <summary>
        /// Guarda el contenido en un archivo .mrk
        /// </summary>
        private void GuardarArchivo(string outputFolder, int archivoNumero, string contenido)
        {
            string filePath = Path.Combine(outputFolder, $"parte{archivoNumero + 1}.mrk");
            File.WriteAllText(filePath, contenido);
        }

        /// <summary>
        /// Busca los archivos parte&lt;número&gt; con la extensión indicada (".mrk" o ".mrc") que dejó una corrida anterior.
        /// Solo cuenta nombres exactos como parte12.mrk; cualquier otro archivo de la carpeta se ignora.
        /// </summary>
        public static List<string> BuscarArchivosParte(string carpeta, string extension)
        {
            if (string.IsNullOrEmpty(carpeta) || !Directory.Exists(carpeta))
                return new List<string>();

            Regex patron = new Regex(@"^parte\d+" + Regex.Escape(extension) + "$", RegexOptions.IgnoreCase);
            return Directory.GetFiles(carpeta, "parte*" + extension)
                .Where(a => patron.IsMatch(Path.GetFileName(a)))
                .ToList();
        }

        /// <summary>
        /// Si la carpeta tiene archivos parte&lt;número&gt; de una corrida anterior, pregunta si se borran.
        /// Devuelve false si el usuario no quiere borrarlos o si no se pudieron borrar: en ese caso no se debe continuar.
        /// </summary>
        public static bool ConfirmarYBorrarPartesAnteriores(string carpeta, string extension, DialogService dialogos)
        {
            List<string> anteriores = BuscarArchivosParte(carpeta, extension);
            if (anteriores.Count == 0)
                return true;

            bool borrar = dialogos.MostrarConfirmacion(
                $"La carpeta ya tiene {anteriores.Count:N0} archivos parteN{extension} de una exportación anterior.\n\n" +
                $"Si se quedan, se mezclarán con los nuevos al convertir y unir, y los registros se duplicarán en Koha.\n\n" +
                $"¿Borrarlos y generar los nuevos?\n\n{carpeta}",
                "Archivos anteriores");
            if (!borrar)
                return false;

            foreach (string archivo in anteriores)
            {
                try
                {
                    File.Delete(archivo);
                }
                catch (Exception ex)
                {
                    dialogos.MostrarError(
                        $"No se pudo borrar {Path.GetFileName(archivo)}:\n{ex.Message}\n\n" +
                        "Si está abierto en otro programa, ciérrelo e intente de nuevo. No se generó ningún archivo nuevo.",
                        "Archivos anteriores");
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Obtiene estadísticas de la exportación
        /// </summary>
        public string ObtenerEstadisticas(DataGridView dataGridView)
        {
            if (dataGridView.SelectedRows.Count == 0)
                return "No hay filas seleccionadas";

            HashSet<string> uniqueFichaNumbers = new HashSet<string>();

            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                var fichaNumber = row.Cells[0].Value?.ToString();
                if (!string.IsNullOrEmpty(fichaNumber))
                {
                    uniqueFichaNumbers.Add(fichaNumber);
                }
            }

            return $"Fichas únicas: {uniqueFichaNumbers.Count}, Registros totales: {dataGridView.SelectedRows.Count}";
        }
    }
}
