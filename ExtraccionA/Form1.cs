using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Reflection;
using System.Diagnostics;
using ExtraccionA.Properties;
using ExtraccionA.Services; // Importar los nuevos servicios

namespace ExtraccionA
{
    public partial class Form1 : Form
    {
        // Servicios
        private DatabaseService _databaseService;
        private DataProcessorService _dataProcessor;
        private FileExportService _fileExporter;
        private ConfigurationManager _config;
        private DialogService _dialogService;
        private UIStateManager _uiManager;
        private ValidationService _validationService;
        private LibraryManagerService _libraryManager;
        private WindowManagerService _windowManager;
        private MarcFileOperationsService _marcFileOps;

        // Ruta del .mdb abierto, solo para mostrarla en la barra de estado
        private string _rutaAccess;

        public Form1()
        {
            InitializeComponent();
            InicializarServicios();
        }

        /// <summary>
        /// Muestra en la barra de estado el resultado del último paso o lo que sigue
        /// </summary>
        private void MostrarEstado(string mensaje)
        {
            estadoMensaje.Text = mensaje;
            estadoMensaje.ToolTipText = mensaje;
        }

        /// <summary>
        /// Actualiza en la barra de estado lo que ya está configurado: MarcEdit, base y carpeta de salida.
        /// Se muestra solo el nombre; la ruta completa aparece al pasar el mouse.
        /// </summary>
        private void ActualizarBarraEstado()
        {
            PonerRutaEnEstado(estadoMarcEdit, "MarcEdit", _config.MarcEditPath, "(sin elegir; se pedirá al convertir)");
            PonerRutaEnEstado(estadoBase, "Base", _rutaAccess, "(sin abrir)");
            PonerRutaEnEstado(estadoSalida, "Salida", _config.OutputFolder, "(sin elegir; se pedirá al formatear)");
        }

        private static void PonerRutaEnEstado(ToolStripStatusLabel etiqueta, string titulo, string ruta, string sinValor)
        {
            if (string.IsNullOrEmpty(ruta))
            {
                etiqueta.Text = $"{titulo}: {sinValor}";
                etiqueta.ToolTipText = etiqueta.Text;
                return;
            }

            string nombre = System.IO.Path.GetFileName(ruta.TrimEnd('\\', '/'));
            etiqueta.Text = $"{titulo}: {(string.IsNullOrEmpty(nombre) ? ruta : nombre)}";
            etiqueta.ToolTipText = ruta;
        }

        /// <summary>
        /// Inicializa los servicios de la aplicación
        /// </summary>
        private void InicializarServicios()
        {
            _config = ConfigurationManager.Instance;
            _dataProcessor = new DataProcessorService();
            _fileExporter = new FileExportService();
            _dialogService = new DialogService();
            _uiManager = new UIStateManager();
            _validationService = new ValidationService(_dialogService);
            _libraryManager = new LibraryManagerService(_dialogService, _validationService);
            _marcFileOps = new MarcFileOperationsService(_dialogService, _validationService);
            // WindowManagerService se inicializa después cuando _databaseService esté disponible
        }

        //Primera Tabla
        private async void fichas_Click(object sender, EventArgs e) => await FetchData(TipoDato.Fichas, fichas);
        private async void usuarios_Click(object sender, EventArgs e) => await FetchData(TipoDato.Usuarios, usuarios);
        private async void ejemplares_Click(object sender, EventArgs e) => await FetchData(TipoDato.Ejemplares, ejemplares);

        private enum TipoDato { Fichas, Usuarios, Ejemplares }

        private async Task FetchData(TipoDato tipo, Button button)
        {
            Cursor = Cursors.WaitCursor;
            _uiManager.SetControlsEnabled(this, false);
            tabla_1.ReadOnly = true;
            // Fichas, Usuarios y Ejemplares quedan siempre habilitados: cada clic reemplaza la tabla 1

            try
            {
                DataTable dataTable = null;
                
                if (tipo == TipoDato.Fichas)
                    dataTable = await _databaseService.ObtenerFichasAsync();
                else if (tipo == TipoDato.Usuarios)
                    dataTable = await _databaseService.ObtenerUsuariosAsync();
                else if (tipo == TipoDato.Ejemplares)
                    dataTable = await _databaseService.ObtenerEjemplaresAsync();

                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    _dialogService.MostrarInformacion("No se encontraron resultados para la consulta.");
                    return;
                }

                tabla_1.DataSource = dataTable;
                tabla_1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

                // El botón Limpiar ya no se muestra; no hace falta para cambiar de tabla
                _uiManager.MostrarGrupoControles(true, fee);
                fee.Text = $"{button.Text} encontradas: " + tabla_1.RowCount;
                _uiManager.MostrarGrupoControles(tipo == TipoDato.Usuarios, exusers);
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                _uiManager.SetControlsEnabled(this, true);
            }
        }



        //Segunda Tabla
        private async void Llenado_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            _uiManager.SetControlsEnabled(this, false);

            try
            {
                DataTable dataTable = await _databaseService.ObtenerFichasConEjemplaresAsync();

                if (dataTable == null || dataTable.Rows.Count == 0)
                {
                    _dialogService.MostrarInformacion("No se encontraron resultados.");
                    return;
                }

                tabla_2.DataSource = dataTable;
                _uiManager.MostrarGrupoControles(true, split, reg);
                reg.Text = "Registros encontrados: " + tabla_2.RowCount;
                MostrarEstado($"{tabla_2.RowCount:N0} ejemplares unidos con sus fichas. Las fichas sin ejemplares se muestran en otra ventana. Siguiente: Dividir.");

                LlenadoSinEjemplares();
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
                _uiManager.SetControlsEnabled(this, true);
            }
        }

        private void limpiar_Click(object sender, EventArgs e)
        {
            tabla_1.ReadOnly = false;
            tabla_1.DataSource = null;
            _uiManager.MostrarGrupoControles(false, limpiar, exusers, fee);
            _uiManager.SetButtonState(null, fichas, usuarios, ejemplares);
            fee.Text = ": :";
        }

        private async void split_Click(object sender, EventArgs e)
        {
            // Dividir solo arma la tabla 3 en memoria; la carpeta de salida se pide al Formatear.
            // El diálogo de biblioteca muestra nombre y código, que se pueden cambiar ahí mismo.
            codigo cd = new codigo();
            await CargarDatosBiblioteca(cd);

            if (cd.ShowDialog() == DialogResult.OK)
            {
                string nombrebiblio = cd.Nom;
                _config.CodigoBiblioteca = cd.Codi;
                _config.NombreBiblioteca = nombrebiblio;

                if (_validationService.ValidarCamposNoVacios(true, (nombrebiblio, "Nombre de biblioteca"), (cd.Codi, "Código de biblioteca")))
                {
                    Cursor = Cursors.WaitCursor;
                    _uiManager.SetControlsEnabled(this, false);

                    try
                    {
                        SplitTable(nombrebiblio);
                        marc.Visible = true;
                        marc.Text = "Registros formateados: " + tabla_3.RowCount;
                    }
                    catch (Exception ex)
                    {
                        _dialogService.MostrarError($"Error al procesar datos: {ex.Message}");
                    }
                    finally
                    {
                        Cursor = Cursors.Default;
                        _uiManager.SetControlsEnabled(this, true);
                    }
                }
            }
            else
            {
                MostrarEstado("División cancelada.");
            }
        }

        private async Task CargarDatosBiblioteca(codigo cd)
        {
            // Usar la biblioteca elegida al abrir la base (o la confirmada en el último Dividir)
            if (!string.IsNullOrEmpty(_config.NombreBiblioteca))
            {
                cd.datos(_config.NombreBiblioteca, _config.CodigoBiblioteca);
                return;
            }

            try
            {
                DataTable bibliotecas = await _databaseService.ObtenerBibliotecasAsync();

                if (bibliotecas != null && bibliotecas.Rows.Count > 0)
                {
                    DataRow row = bibliotecas.Rows[0];
                    string nombre = row["Nombre"].ToString();
                    string code = _validationService.FormatearCodigoBiblioteca(row["No"].ToString());
                    cd.datos(nombre, code);
                }
                else
                {
                    _dialogService.MostrarAdvertencia("No se encontraron bibliotecas.");
                }
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError($"Error al cargar bibliotecas: {ex.Message}");
            }
        }

        private void SplitTable(string nombrebiblio)
        {
            DataTable originalTable = (DataTable)tabla_2.DataSource;
            if (originalTable == null) return;

            DataTable processedTable = _dataProcessor.ProcesarTablaParaMARC(originalTable, nombrebiblio);
            
            if (processedTable != null)
            {
                tabla_3.DataSource = processedTable;
                // Que el usuario elija las filas; al cargar no queda ninguna seleccionada por accidente
                tabla_3.ClearSelection();
                ActualizarControlesSeleccion();
                MostrarEstado($"{processedTable.Rows.Count:N0} registros listos para formatear. Las fichas sin ejemplares no se incluyen. Seleccione filas y pulse Formatear.");
                Console.WriteLine($"Filas procesadas: {processedTable.Rows.Count}");
            }
        }

        private void format_Click(object sender, EventArgs e)
        {
            // Sin carpeta de salida se abre el selector y, si se elige una, se sigue formateando sin volver a pulsar
            if (!_validationService.ValidarCarpetaSalida(_config, mostrarMensaje: false) &&
                !PathRutadeSalida("Seleccione la carpeta donde se guardarán los archivos .mrk", mostrarMensajes: false))
                return;

            if (!_validationService.ValidarFilasSeleccionadas(tabla_3))
                return;

            // Partes de una corrida anterior se mezclarían con las nuevas al convertir y unir
            if (!FileExportService.ConfirmarYBorrarPartesAnteriores(_config.OutputFolder, ".mrk", _dialogService))
                return;

            try
            {
                int archivosGenerados = _fileExporter.ExportarFilasAMRK(tabla_3, _config.OutputFolder, _config.CodigoBiblioteca);
                _dialogService.MostrarInformacion($"Se generaron {archivosGenerados:N0} archivos .mrk en:\n{_config.OutputFolder}", "Formateo completado");
                MostrarEstado($"{archivosGenerados:N0} archivos .mrk generados. Siguiente: Conversión MRK a MRC.");
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError($"Error al exportar: {ex.Message}");
            }
        }

        private void tabla_3_SelectionChanged(object sender, EventArgs e)
        {
            ActualizarControlesSeleccion();
        }

        /// <summary>
        /// Muestra Sel. Todas y los contadores en cuanto la tabla 3 tiene datos;
        /// Formatear se ve siempre, pero solo se habilita con al menos una fila seleccionada
        /// </summary>
        private void ActualizarControlesSeleccion()
        {
            bool hayDatos = tabla_3.DataSource != null && tabla_3.Rows.Count > 0;
            _uiManager.MostrarGrupoControles(hayDatos, format, selall, label2, regsel);
            if (!hayDatos)
                return;

            int seleccionadas = tabla_3.SelectedRows.Count;
            format.Enabled = seleccionadas > 0;

            if (seleccionadas > 0)
                _uiManager.ActualizarLabelFichasSeleccionadas(tabla_3, label2);
            else
                label2.Text = "Fichas seleccionadas: 0";

            regsel.Text = "Registros Seleccionados: " + seleccionadas;
        }

        private void selall_Click(object sender, EventArgs e)
        {
            // SelectAll selecciona todas las filas con un solo evento SelectionChanged.
            // Seleccionarlas una por una disparaba un evento por fila, y cada evento
            // recorría todas las filas seleccionadas (tiempo cuadrático: minutos con 40,000 filas).
            tabla_3.SelectAll();
        }

        private void seleccionarAccess_Click(object sender, EventArgs e)
        {
            string archivoSeleccionado = _dialogService.SeleccionarArchivoAccess();
            
            if (!string.IsNullOrEmpty(archivoSeleccionado))
            {
                _config.SetAccessConnection(archivoSeleccionado);
                // Base nueva: la biblioteca se vuelve a elegir en NameBiblio()
                _config.NombreBiblioteca = string.Empty;
                _config.CodigoBiblioteca = string.Empty;
                _databaseService = new DatabaseService(_config.ConnectionString);
                _windowManager = new WindowManagerService(_databaseService, _dialogService);

                _rutaAccess = archivoSeleccionado;
                ActualizarBarraEstado();
                MostrarEstado("Base abierta. Pulse Llenar para unir fichas y ejemplares.");

                _uiManager.LimpiarDataGridViews(tabla_1, tabla_2, tabla_3);
                _uiManager.LimpiarLabels(": :", fee, reg, marc, biblioteca, codebiblio);
                ActualizarControlesSeleccion();

                fichas.Enabled = usuarios.Enabled = ejemplares.Enabled = llenado.Enabled = archivosDeSalidaToolStripMenuItem.Enabled = archivosDeSalidaToolStripMenuItem.Visible = split.Enabled = true;
                fee.Visible = reg.Visible = marc.Visible = split.Visible = biblioteca.Visible = codebiblio.Visible = exusers.Visible = false;

                // La carpeta de salida ya no se recuerda aquí: Formatear la pide si falta
                NameBiblio();
            }
            else
            {
                _dialogService.MostrarAdvertencia("No se seleccionó ningún archivo.");
            }
        }

        private void NameBiblio()
        {
            _libraryManager.CargarYSeleccionarBiblioteca(_config.ConnectionString, biblioteca, codebiblio);
        }

        /// <summary>
        /// Pide la ruta de cmarcedit.exe. Si el usuario cancela, el programa sigue funcionando
        /// y la ruta se volverá a pedir cuando se necesite (conversión a MRC).
        /// </summary>
        /// <returns>true si hay una ruta de MarcEdit configurada al terminar</returns>
        private bool PathMarc()
        {
            while (string.IsNullOrEmpty(_config.MarcEditPath))
            {
                string rutaMarcEdit = _dialogService.SeleccionarMarcEdit();

                if (string.IsNullOrEmpty(rutaMarcEdit))
                {
                    _dialogService.MostrarAdvertencia("No se seleccionó MarcEdit. Podrá seleccionarlo más tarde; se le pedirá al convertir archivos MRK a MRC.");
                    return false;
                }

                bool esCmarcedit = string.Equals(System.IO.Path.GetFileName(rutaMarcEdit), "cmarcedit.exe", StringComparison.OrdinalIgnoreCase);
                if (!esCmarcedit && !_dialogService.MostrarConfirmacion(
                        $"El archivo seleccionado no es cmarcedit.exe:\n{rutaMarcEdit}\n\nLa conversión necesita la versión de consola de MarcEdit (cmarcedit.exe). ¿Desea usar este archivo de todos modos?"))
                {
                    continue;
                }

                _config.MarcEditPath = rutaMarcEdit;
                ActualizarBarraEstado();
            }

            return true;
        }

        private void seleccionarRutaDeSalidaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PathRutadeSalida();
        }

        /// <summary>
        /// Pide la carpeta de salida. Devuelve true si se eligió una.
        /// Los mensajes de confirmación solo se muestran cuando se usa desde el menú.
        /// </summary>
        private bool PathRutadeSalida(string descripcion = "Seleccionar Carpeta de Salida", bool mostrarMensajes = true)
        {
            string carpetaSalida = _dialogService.SeleccionarCarpeta(descripcion);

            if (!string.IsNullOrEmpty(carpetaSalida))
            {
                _config.OutputFolder = carpetaSalida;
                ActualizarBarraEstado();
                if (mostrarMensajes)
                    _dialogService.MostrarInformacion($"Ruta de salida seleccionada: {_config.OutputFolder}");
                return true;
            }

            if (mostrarMensajes)
                _dialogService.MostrarAdvertencia("No se seleccionó ninguna carpeta.");
            return false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Un solo mensaje de bienvenida; lo demás se informa en la barra de estado
            _dialogService.MostrarInformacion(
                "Bienvenido al Sistema de Migración de Datos SIABUC8 → Koha.\n\n" +
                "Pasos para comenzar:\n" +
                "1) Seleccione MarcEdit (cmarcedit.exe).\n" +
                "2) Abra la base SIABUC8.mdb en el menú Archivo de entrada.",
                "Bienvenido");
            PathMarc();

            ActualizarBarraEstado();
            MostrarEstado("Abra la base SIABUC8.mdb en el menú Archivo de entrada.");

            toolTip1.SetToolTip(tabla1, "Haz click para mas informacion.");
            toolTip2.SetToolTip(tabla2, "Haz click para mas informacion.");
            toolTip3.SetToolTip(tabla3, "Haz click para mas informacion.");
        }

        private void tabla1_Click(object sender, EventArgs e)
        {
            _dialogService.MostrarInformacion("Información directa del Access, tabla de fichas, ejemplares y usuarios.", "Información");
        }

        private void tabla2_Click(object sender, EventArgs e)
        {
            _dialogService.MostrarInformacion("Tabla de informacion procesada para formato MARC.", "Información");
        }

        private void tabla3_Click(object sender, EventArgs e)
        {
            _dialogService.MostrarInformacion("Tabla formateada para exportacion por formato MARC para posterior conversion.", "Información");
        }

        private async void seleccionarArchivosMRK_Click(object sender, EventArgs e)
        {
            if (!PathMarc())
                return;

            Cursor = Cursors.WaitCursor;
            _uiManager.SetControlsEnabled(this, false);

            try
            {
                await _marcFileOps.ConvertirMRKaMRCAsync(_config.MarcEditPath, this);
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError($"Error al convertir archivos: {ex.Message}");
            }
            finally
            {
                Cursor = Cursors.Default;
                _uiManager.SetControlsEnabled(this, true);
            }
        }

        private void selMRC_Click(object sender, EventArgs e)
        {
            _marcFileOps.UnirArchivosMRC();
        }

        private void selMRK_Click(object sender, EventArgs e)
        {
            _marcFileOps.UnirArchivosMRK();
        }

        //Fichas sin ejemplares
        private async void LlenadoSinEjemplares()
        {
            if (_windowManager != null)
            {
                await _windowManager.MostrarVentanaFichasSinEjemplares();
            }
            else
            {
                _dialogService.MostrarAdvertencia("Debe conectarse a la base de datos primero");
            }
        }

        private void exusers_Click(object sender, EventArgs e)
        {
            EUsuarios extraer = new EUsuarios();
            string codigo = extraer.CodigoBiblio(_config.CodigoBiblioteca);
            if (codigo == null)
                return; // Cancelado: no se genera el CSV

            extraer.Extraer(_config.ConnectionString, codigo);
        }
    }
}
