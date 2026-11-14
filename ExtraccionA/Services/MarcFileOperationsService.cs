using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para operaciones con archivos MARC (unión y conversión)
    /// </summary>
    public class MarcFileOperationsService
    {
        private readonly DialogService _dialogService;
        private readonly ValidationService _validationService;

        public MarcFileOperationsService(DialogService dialogService, ValidationService validationService)
        {
            _dialogService = dialogService;
            _validationService = validationService;
        }

        /// <summary>
        /// Procesa la conversión de archivos MRK a MRC
        /// </summary>
        public async Task ConvertirMRKaMRCAsync(string marcEditPath, Form ventanaPrincipal)
        {
            string carpetaMRK = _dialogService.SeleccionarCarpeta("Seleccionar Carpeta de Archivos MRK");
            if (string.IsNullOrEmpty(carpetaMRK))
            {
                _dialogService.MostrarAdvertencia("No se seleccionó ninguna carpeta.");
                return;
            }

            int totalMRK = Directory.GetFiles(carpetaMRK, "*.mrk").Length;
            if (totalMRK == 0)
            {
                _dialogService.MostrarAdvertencia("La carpeta seleccionada no contiene archivos .mrk.");
                return;
            }

            string carpetaSalida = _dialogService.SeleccionarCarpeta("Seleccionar Carpeta de Salida");
            if (string.IsNullOrEmpty(carpetaSalida))
            {
                _dialogService.MostrarAdvertencia("No se seleccionó ninguna carpeta.");
                return;
            }

            Convertirmrc cnmrc = new Convertirmrc
            {
                UbicacionArchivosMRK = carpetaMRK,
                UbicacionSalida = carpetaSalida
            };

            ResultadoConversion resultado;
            using (CancellationTokenSource cancelacion = new CancellationTokenSource())
            using (VentanaProgreso ventana = new VentanaProgreso("Conversión MRK a MRC", totalMRK, cancelacion))
            {
                ventana.Show(ventanaPrincipal);
                var progreso = new Progress<ProgresoConversion>(ventana.Actualizar);

                try
                {
                    resultado = await Task.Run(() => cnmrc.ConvertirMRKaMRC(marcEditPath, progreso, cancelacion.Token));
                }
                catch (Exception ex)
                {
                    ventana.Terminar();
                    _dialogService.MostrarError($"Error al procesar archivos: {ex.Message}");
                    return;
                }

                ventana.Terminar();
            }

            MostrarResumenConversion(resultado, carpetaSalida);
        }

        /// <summary>
        /// Muestra cuántos archivos se convirtieron y por qué fallaron los demás
        /// </summary>
        private void MostrarResumenConversion(ResultadoConversion r, string carpetaSalida)
        {
            string archivoErrores = r.ArchivoErrores != null
                ? $"\n\nDetalle completo en:\n{r.ArchivoErrores}"
                : "";

            if (r.DetenidoEnPrimerArchivo)
            {
                var primero = r.Fallidos[0];
                _dialogService.MostrarError(
                    $"No se pudo convertir el primer archivo ({primero.Key}), así que la conversión se detuvo.\n\n" +
                    $"Motivo: {primero.Value}{archivoErrores}",
                    "Conversión detenida");
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(r.Cancelado ? "Conversión cancelada.\n\n" : "");
            sb.Append($"Convertidos: {r.Convertidos:N0} de {r.Total:N0} archivos.");

            if (r.Fallidos.Count == 0)
            {
                sb.Append($"\n\nCarpeta de salida:\n{carpetaSalida}");
                if (r.Cancelado)
                    _dialogService.MostrarAdvertencia(sb.ToString(), "Conversión cancelada");
                else
                    _dialogService.MostrarInformacion(sb.ToString(), "Conversión completada");
                return;
            }

            const int maximoEnPantalla = 10;
            sb.Append($"\nFallidos: {r.Fallidos.Count:N0}\n");
            foreach (var fallo in r.Fallidos.Take(maximoEnPantalla))
                sb.Append($"\n• {fallo.Key}: {fallo.Value}");
            if (r.Fallidos.Count > maximoEnPantalla)
                sb.Append($"\n… y {r.Fallidos.Count - maximoEnPantalla:N0} más.");
            sb.Append(archivoErrores);

            _dialogService.MostrarAdvertencia(sb.ToString(), r.Cancelado ? "Conversión cancelada" : "Conversión con errores");
        }

        /// <summary>
        /// Une múltiples archivos MRC en uno solo
        /// </summary>
        public void UnirArchivosMRC()
        {
            string carpetaSeleccionada = _dialogService.SeleccionarCarpeta("Seleccionar carpeta que contiene los archivos MRC");

            if (string.IsNullOrEmpty(carpetaSeleccionada))
                return;

            if (!Directory.Exists(carpetaSeleccionada))
            {
                _dialogService.MostrarError("La carpeta seleccionada no es válida.");
                return;
            }

            List<string> archivosMRC = Directory.GetFiles(carpetaSeleccionada, "*.mrc").ToList();

            if (!_validationService.ValidarCantidadMinima(archivosMRC.Count, 2, "archivos MRC"))
            {
                _dialogService.MostrarError("La carpeta debe contener al menos dos archivos MRC para unirlos.");
                return;
            }

            string archivoSalida = SolicitarArchivoSalida("MRC");
            if (string.IsNullOrEmpty(archivoSalida))
                return;

            UnirMRKoMRC umrc = new UnirMRKoMRC();
            umrc.UnirArchivosMRC(carpetaSeleccionada, archivoSalida);
        }

        /// <summary>
        /// Une múltiples archivos MRK en uno solo
        /// </summary>
        public void UnirArchivosMRK()
        {
            string carpetaSeleccionada = _dialogService.SeleccionarCarpeta("Seleccionar carpeta que contiene los archivos MRK");

            if (string.IsNullOrEmpty(carpetaSeleccionada))
                return;

            List<string> archivosMRK = Directory.GetFiles(carpetaSeleccionada, "*.mrk").ToList();

            if (!_validationService.ValidarCantidadMinima(archivosMRK.Count, 2, "archivos MRK"))
            {
                _dialogService.MostrarError("La carpeta debe contener al menos dos archivos MRK para unirlos.");
                return;
            }

            string archivoSalida = SolicitarArchivoSalida("MRK");
            if (string.IsNullOrEmpty(archivoSalida))
                return;

            UnirMRKoMRC umrc = new UnirMRKoMRC();
            umrc.UnirArchivosMRK(carpetaSeleccionada, archivoSalida);
        }

        /// <summary>
        /// Solicita al usuario la ubicación del archivo de salida
        /// </summary>
        private string SolicitarArchivoSalida(string extension)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = $"Archivo {extension} (*.{extension.ToLower()})|*.{extension.ToLower()}";
                saveFileDialog.Title = "Guardar Archivo Unido";
                saveFileDialog.FileName = $"Archivo_Unificado.{extension.ToLower()}";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string archivoSalida = Path.GetFullPath(saveFileDialog.FileName);

                    if (string.IsNullOrWhiteSpace(archivoSalida))
                    {
                        _dialogService.MostrarError("El archivo de salida no es válido.");
                        return null;
                    }

                    return archivoSalida;
                }
            }

            return null;
        }
    }
}
