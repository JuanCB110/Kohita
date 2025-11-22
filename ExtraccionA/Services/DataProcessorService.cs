using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio encargado del procesamiento y transformación de datos bibliográficos
    /// </summary>
    public class DataProcessorService
    {
        /// <summary>
        /// Obtiene el código de tipo de material para Koha basado en el tipo de SIABUC
        /// </summary>
        public string ObtenerCodigoTipoMaterial(string tipoMaterial)
        {
            switch (tipoMaterial)
            {
                case "0": return "OR";  // Tipo No especificado
                case "1": return "BK";  // Tipo Libro
                case "2": return "TS";  // Tipo Tesis
                case "3": return "MP";  // Tipo Mapas
                case "4": return "VD";  // Tipo Video
                case "5": return "VM";  // Tipo CD-ROM, DVD
                case "6": return "PS";  // Tipo Fotografías
                case "7": return "DS";  // Tipo Diapositivas
                case "8": return "DT";  // Tipo Disquetes
                case "9": return "MU";  // Tipo Música
                case "10": return "MM"; // Tipo Microfilm
                case "11": return "CF"; // Tipo Archivos de cómputo
                case "12": return "PN"; // Tipo Pintura
                case "13": return "TS"; // Tipo Escultura
                case "14": return "OR"; // Tipo Otro
                default: return "OR";
            }
        }

        /// <summary>
        /// Procesa la tabla original y la transforma en formato MARC
        /// </summary>
        public DataTable ProcesarTablaParaMARC(DataTable originalTable, string nombreBiblioteca)
        {
            if (originalTable == null) return null;

            DataTable processedTable = CrearEstructuraTablaMARC();
            string fechaActual = DateTime.Now.ToString("yyyyMMddHHmmss");

            // Agrupar por Numero_De_Control
            var groupedData = originalTable.AsEnumerable()
                .GroupBy(row => row["Numero_De_Control"].ToString())
                .Select(group => new
                {
                    NumeroDeControl = string.IsNullOrEmpty(group.Key) ? null : group.Key,
                    NumeroDeItems = string.Join(",", group
                        .Select(row => row["Numero_de_Item"]?.ToString())
                        .Where(item => !string.IsNullOrEmpty(item))
                        .Distinct()),
                    DatosFijos = group.FirstOrDefault()?["DatosFijos"]?.ToString(),
                    MarcCompleto = group.FirstOrDefault()?["MARC_Completo"]?.ToString(),
                    TipoMaterial = group.FirstOrDefault()?["TipoMaterial"]?.ToString(),
                    TituloAux = group.FirstOrDefault()?["Titulo_Auxiliar"]?.ToString(),
                    ISBNAux = group.FirstOrDefault()?["ISBN_Auxiliar"]?.ToString(),
                    AutorAux = group.FirstOrDefault()?["Autor_Auxiliar"]?.ToString(),
                    LCCAux = group.FirstOrDefault()?["Clasificacion_LC"]?.ToString()
                }).ToList();

            var addedColumns = new HashSet<string>();
            var datosFijosService = new DatosFijosCorrecto();

            // Procesar cada grupo
            foreach (var group in groupedData)
            {
                DataRow newRow = processedTable.NewRow();

                // Asignar campos básicos
                newRow["003"] = group.NumeroDeControl;
                newRow["005"] = fechaActual;
                newRow["Numeros_de_Items"] = group.NumeroDeItems;
                newRow["008"] = datosFijosService.Codificar(group.DatosFijos);
                newRow["040"] = nombreBiblioteca;
                newRow["942"] = ObtenerCodigoTipoMaterial(group.TipoMaterial);
                newRow["Titulo_Auxiliar"] = group.TituloAux;
                newRow["ISBN_Auxiliar"] = group.ISBNAux;
                newRow["Autor_Auxiliar"] = group.AutorAux;
                newRow["LCC_Auxiliar"] = group.LCCAux;

                // Procesar MARC_Completo
                if (!string.IsNullOrEmpty(group.MarcCompleto))
                {
                    ProcesarEtiquetasMARC(group.MarcCompleto, newRow, processedTable, addedColumns);
                }

                processedTable.Rows.Add(newRow);
            }

            // Completar campos vacíos con auxiliares
            CompletarCamposConAuxiliares(processedTable);

            // Eliminar columnas auxiliares
            EliminarColumnasAuxiliares(processedTable);

            return processedTable;
        }

        /// <summary>
        /// Crea la estructura de la tabla procesada
        /// </summary>
        private DataTable CrearEstructuraTablaMARC()
        {
            DataTable table = new DataTable();
            table.Columns.Add("003", typeof(string)); // Identificador de control
            table.Columns.Add("005", typeof(string)); // Fecha y hora de la última modificación
            table.Columns.Add("Numeros_de_Items", typeof(string)); // Números de Items
            table.Columns.Add("008", typeof(string)); // Datos fijos
            table.Columns.Add("040", typeof(string)); // Fuente de catálogo
            table.Columns.Add("942", typeof(string)); // Tipo de material
            table.Columns.Add("Titulo_Auxiliar", typeof(string)); // Título Auxiliar
            table.Columns.Add("ISBN_Auxiliar", typeof(string)); // ISBN Auxiliar
            table.Columns.Add("Autor_Auxiliar", typeof(string)); // Autor Auxiliar
            table.Columns.Add("LCC_Auxiliar", typeof(string)); // Clasificación LC Auxiliar
            return table;
        }

        /// <summary>
        /// Procesa las etiquetas MARC y las agrega a la fila
        /// </summary>
        private void ProcesarEtiquetasMARC(string marcCompleto, DataRow row, DataTable table, HashSet<string> addedColumns)
        {
            string sinultimo = marcCompleto.Substring(0, marcCompleto.Length - 1);
            var marcParts = sinultimo.Split('¦')
                .Where(s => s.Length > 3)
                .OrderBy(part => int.TryParse(part.Substring(0, 3), out var res) ? res : int.MaxValue);

            foreach (string part in marcParts)
            {
                string etiqueta = part.Substring(0, 3);
                string valor = part.Substring(3);

                // Añadir la columna si no está ya presente
                if (!addedColumns.Contains(etiqueta))
                {
                    table.Columns.Add(etiqueta, typeof(string));
                    addedColumns.Add(etiqueta);
                }
                row[etiqueta] = valor;
            }
        }

        /// <summary>
        /// Completa campos principales con valores auxiliares si están vacíos
        /// </summary>
        private void CompletarCamposConAuxiliares(DataTable table)
        {
            Dictionary<string, string> campos = new Dictionary<string, string>
            {
                { "020", "ISBN_Auxiliar" },
                { "245", "Titulo_Auxiliar" },
                { "100", "Autor_Auxiliar" },
                { "050", "LCC_Auxiliar" }
            };

            foreach (DataRow row in table.Rows)
            {
                foreach (var campo in campos)
                {
                    string principal = campo.Key;
                    string auxiliar = campo.Value;

                    string valorPrincipal = table.Columns.Contains(principal) && row[principal] != DBNull.Value 
                        ? row[principal].ToString().Trim() 
                        : string.Empty;
                    
                    string valorAuxiliar = table.Columns.Contains(auxiliar) && row[auxiliar] != DBNull.Value 
                        ? row[auxiliar].ToString().Trim() 
                        : string.Empty;

                    // Si la columna principal está vacía, asignar el valor del auxiliar
                    if (string.IsNullOrEmpty(valorPrincipal) && !string.IsNullOrEmpty(valorAuxiliar))
                    {
                        if (!table.Columns.Contains(principal))
                        {
                            table.Columns.Add(principal, typeof(string));
                        }
                        row[principal] = valorAuxiliar;
                    }
                }
            }
        }

        /// <summary>
        /// Elimina las columnas auxiliares de la tabla
        /// </summary>
        private void EliminarColumnasAuxiliares(DataTable table)
        {
            string[] columnasAuxiliares = { "ISBN_Auxiliar", "Titulo_Auxiliar", "Autor_Auxiliar", "###", "LCC_Auxiliar", "082" };

            foreach (string columna in columnasAuxiliares)
            {
                if (table.Columns.Contains(columna))
                {
                    table.Columns.Remove(columna);
                }
            }
        }
    }
}
