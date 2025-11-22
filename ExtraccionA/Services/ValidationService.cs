using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para validaciones de datos y estado de la aplicación
    /// </summary>
    public class ValidationService
    {
        private readonly DialogService _dialogService;

        public ValidationService(DialogService dialogService = null)
        {
            _dialogService = dialogService ?? new DialogService();
        }

        /// <summary>
        /// Valida que haya una conexión establecida
        /// </summary>
        public bool ValidarConexion(ConfigurationManager config, bool mostrarMensaje = true)
        {
            if (string.IsNullOrEmpty(config.ConnectionString))
            {
                if (mostrarMensaje)
                    _dialogService.MostrarAdvertencia("Debe seleccionar primero un archivo Access (SIABUC8.mdb)");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Valida que haya una carpeta de salida configurada
        /// </summary>
        public bool ValidarCarpetaSalida(ConfigurationManager config, bool mostrarMensaje = true)
        {
            if (string.IsNullOrEmpty(config.OutputFolder))
            {
                if (mostrarMensaje)
                    _dialogService.MostrarAdvertencia("Debe seleccionar una carpeta de salida");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Valida que MarcEdit esté configurado
        /// </summary>
        public bool ValidarMarcEdit(ConfigurationManager config, bool mostrarMensaje = true)
        {
            if (string.IsNullOrEmpty(config.MarcEditPath))
            {
                if (mostrarMensaje)
                    _dialogService.MostrarAdvertencia("Debe configurar la ruta de MarcEdit");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Valida que haya filas seleccionadas en un DataGridView
        /// </summary>
        public bool ValidarFilasSeleccionadas(DataGridView grid, bool mostrarMensaje = true)
        {
            if (grid.SelectedRows.Count == 0)
            {
                if (mostrarMensaje)
                    _dialogService.MostrarAdvertencia("Debe seleccionar al menos una fila");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Valida que una tabla tenga datos
        /// </summary>
        public bool ValidarTablaTieneDatos(DataGridView grid, bool mostrarMensaje = true)
        {
            if (grid.DataSource == null || grid.Rows.Count == 0)
            {
                if (mostrarMensaje)
                    _dialogService.MostrarAdvertencia("No hay datos para procesar");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Valida que se haya configurado el código de biblioteca
        /// </summary>
        public bool ValidarCodigoBiblioteca(ConfigurationManager config, bool mostrarMensaje = true)
        {
            if (string.IsNullOrEmpty(config.CodigoBiblioteca))
            {
                if (mostrarMensaje)
                    _dialogService.MostrarAdvertencia("Debe configurar el código de biblioteca");
                return false;
            }
            return true;
        }

        /// <summary>
        /// Valida que una cantidad sea mayor o igual al mínimo requerido
        /// </summary>
        public bool ValidarCantidadMinima(int cantidad, int minimo, string nombreItem)
        {
            return cantidad >= minimo;
        }

        /// <summary>
        /// Valida configuración completa para exportación
        /// </summary>
        public bool ValidarConfiguracionCompleta(ConfigurationManager config, bool mostrarMensaje = true)
        {
            var errores = new List<string>();

            if (string.IsNullOrEmpty(config.ConnectionString))
                errores.Add("- Archivo Access no seleccionado");

            if (string.IsNullOrEmpty(config.OutputFolder))
                errores.Add("- Carpeta de salida no configurada");

            if (string.IsNullOrEmpty(config.MarcEditPath))
                errores.Add("- MarcEdit no configurado");

            if (string.IsNullOrEmpty(config.CodigoBiblioteca))
                errores.Add("- Código de biblioteca no configurado");

            if (errores.Any())
            {
                if (mostrarMensaje)
                {
                    string mensaje = "Configuración incompleta:\n" + string.Join("\n", errores);
                    _dialogService.MostrarAdvertencia(mensaje);
                }
                return false;
            }

            return true;
        }

        /// <summary>
        /// Valida que un código tenga el formato correcto (3 dígitos)
        /// </summary>
        public string FormatearCodigoBiblioteca(string codigo)
        {
            if (string.IsNullOrEmpty(codigo))
                return "000";

            // Asegurar que tenga 3 dígitos
            return codigo.PadLeft(3, '0');
        }

        /// <summary>
        /// Valida que los valores de entrada no estén vacíos
        /// </summary>
        public bool ValidarCamposNoVacios(bool mostrarMensaje, params (string valor, string nombreCampo)[] campos)
        {
            var camposVacios = campos.Where(c => string.IsNullOrWhiteSpace(c.valor)).ToList();

            if (camposVacios.Any())
            {
                if (mostrarMensaje)
                {
                    string mensaje = "Los siguientes campos son requeridos:\n" +
                                   string.Join("\n", camposVacios.Select(c => $"- {c.nombreCampo}"));
                    _dialogService.MostrarAdvertencia(mensaje);
                }
                return false;
            }

            return true;
        }
    }
}
