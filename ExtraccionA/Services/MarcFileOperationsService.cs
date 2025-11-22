using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        public void ConvertirMRKaMRC(string marcEditPath)
        {
            string carpetaMRK = _dialogService.SeleccionarCarpeta("Seleccionar Carpeta de Archivos MRK");
            if (string.IsNullOrEmpty(carpetaMRK))
            {
                _dialogService.MostrarAdvertencia("No se seleccionó ninguna carpeta.");
                return;
            }

            _dialogService.MostrarInformacion($"Ruta seleccionada: {carpetaMRK}");

            string carpetaSalida = _dialogService.SeleccionarCarpeta("Seleccionar Carpeta de Salida");
            if (string.IsNullOrEmpty(carpetaSalida))
            {
                _dialogService.MostrarAdvertencia("No se seleccionó ninguna carpeta.");
                return;
            }

            _dialogService.MostrarInformacion($"Ruta seleccionada: {carpetaSalida}");

            Convertirmrc cnmrc = new Convertirmrc
            {
                UbicacionArchivosMRK = carpetaMRK,
                UbicacionSalida = carpetaSalida
            };

            cnmrc.ConvertirMRKaMRC(marcEditPath);
        }

        /// <summary>
        /// Une múltiples archivos MRC en uno solo
        /// </summary>
        public void UnirArchivosMRC(string marcEditPath)
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

            _dialogService.MostrarInformacion($"Archivo de salida: {archivoSalida}", "Depuración");

            UnirMRKoMRC umrc = new UnirMRKoMRC();
            umrc.UnirArchivosMRC(marcEditPath, carpetaSeleccionada, archivoSalida);
        }

        /// <summary>
        /// Une múltiples archivos MRK en uno solo
        /// </summary>
        public void UnirArchivosMRK(string marcEditPath)
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
            umrc.UnirArchivosMRK(marcEditPath, carpetaSeleccionada, archivoSalida);
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
