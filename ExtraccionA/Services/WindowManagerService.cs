using System;
using System.Data;
using System.Drawing;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para gestionar ventanas dinámicas de la aplicación
    /// </summary>
    public class WindowManagerService
    {
        private readonly DatabaseService _databaseService;
        private readonly DialogService _dialogService;
        private Form formSinEjemplares = null;

        public WindowManagerService(DatabaseService databaseService, DialogService dialogService)
        {
            _databaseService = databaseService;
            _dialogService = dialogService;
        }

        /// <summary>
        /// Crea y muestra una ventana con las fichas que no tienen ejemplares
        /// </summary>
        public async Task MostrarVentanaFichasSinEjemplares()
        {
            // Si ya hay una ventana abierta, cerrarla
            if (formSinEjemplares != null && !formSinEjemplares.IsDisposed)
            {
                formSinEjemplares.Close();
            }

            // Crear nueva ventana
            formSinEjemplares = new Form
            {
                Text = "Fichas sin ejemplares",
                Size = new Size(800, 500),
                StartPosition = FormStartPosition.CenterScreen
            };

            // Cargar el icono si está disponible
            try
            {
                formSinEjemplares.Icon = new Icon(Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("ExtraccionA.zombie-removebg-preview.ico"));
            }
            catch
            {
                // Si no se puede cargar el icono, continuar sin él
            }

            // Crear DataGridView
            DataGridView tablaSinEjemplares = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            };

            formSinEjemplares.Controls.Add(tablaSinEjemplares);

            // Cargar datos
            await CargarFichasSinEjemplares(tablaSinEjemplares);

            // Mostrar ventana
            formSinEjemplares.Show();
        }

        /// <summary>
        /// Carga las fichas sin ejemplares en el DataGridView
        /// </summary>
        private async Task CargarFichasSinEjemplares(DataGridView tabla)
        {
            try
            {
                DataTable fichasSinEjemplares = await _databaseService.ObtenerFichasSinEjemplaresAsync();
                tabla.DataSource = fichasSinEjemplares;
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError("Error al cargar datos: " + ex.Message);
            }
        }

        /// <summary>
        /// Cierra la ventana de fichas sin ejemplares si está abierta
        /// </summary>
        public void CerrarVentanaFichasSinEjemplares()
        {
            if (formSinEjemplares != null && !formSinEjemplares.IsDisposed)
            {
                formSinEjemplares.Close();
            }
        }
    }
}
