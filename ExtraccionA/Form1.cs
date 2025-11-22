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

        public Form1()
        {
            InitializeComponent();
            InicializarServicios();
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
            _uiManager.SetButtonState(button, fichas, usuarios, ejemplares);

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
                tabla_1.AutoResizeColumns();
                tabla_1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                _uiManager.MostrarGrupoControles(true, limpiar, fee);
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
            if (!_validationService.ValidarCarpetaSalida(_config))
            {
                _dialogService.MostrarAdvertencia("Por favor, seleccione una ruta de salida en el menu superior");
                PathRutadeSalida();
                return;
            }

            _dialogService.MostrarInformacion("Nota:\nTanto codigo de biblioteca como nombre de la misma, son tomados\ndirectamente del access, asi que si es necesario cambiarlos, hagalo");
            
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
                _dialogService.MostrarInformacion("Operación cancelada.");
            }
        }

        private async Task CargarDatosBiblioteca(codigo cd)
        {
            try
            {
                DataTable bibliotecas = await _databaseService.ObtenerBibliotecasAsync();
                
                if (bibliotecas != null && bibliotecas.Rows.Count > 0)
                {
                    foreach (DataRow row in bibliotecas.Rows)
                    {
                        string nombre = row["Nombre"].ToString();
                        string code = _validationService.FormatearCodigoBiblioteca(row["No"].ToString());
                        cd.datos(nombre, code);
                    }
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
                _dialogService.MostrarInformacion("Nota:\nSi los Registros Formateados son menores que las Fichas Encontradas, significa que esas fichas no contienen ejemplares");
                Console.WriteLine($"Filas procesadas: {processedTable.Rows.Count}");
            }
        }

        private void format_Click(object sender, EventArgs e)
        {
            if (!_validationService.ValidarCarpetaSalida(_config))
            {
                PathRutadeSalida();
                return;
            }

            if (!_validationService.ValidarFilasSeleccionadas(tabla_3))
                return;

            try
            {
                int archivosGenerados = _fileExporter.ExportarFilasAMRK(tabla_3, _config.OutputFolder, _config.CodigoBiblioteca);
                _dialogService.MostrarInformacion($"Archivos guardados correctamente en {archivosGenerados} partes.", "Éxito");
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError($"Error al exportar: {ex.Message}");
            }
        }

        private void tabla_3_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            _uiManager.MostrarGrupoControles(true, format, selall, label2, regsel);
        }

        private void tabla_3_SelectionChanged(object sender, EventArgs e)
        {
            _uiManager.ActualizarLabelFichasSeleccionadas(tabla_3, label2, format, selall, label2, regsel);
            regsel.Text = "Registros Seleccionados: " + tabla_3.SelectedRows.Count.ToString();
        }

        private void selall_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in tabla_3.Rows)
            {
                row.Selected = true;
            }
        }

        private void seleccionarAccess_Click(object sender, EventArgs e)
        {
            string archivoSeleccionado = _dialogService.SeleccionarArchivoAccess();
            
            if (!string.IsNullOrEmpty(archivoSeleccionado))
            {
                _config.SetAccessConnection(archivoSeleccionado);
                _databaseService = new DatabaseService(_config.ConnectionString);
                _windowManager = new WindowManagerService(_databaseService, _dialogService);

                _dialogService.MostrarInformacion($"Conexión establecida con la base de datos: {archivoSeleccionado}");

                _uiManager.LimpiarDataGridViews(tabla_1, tabla_2, tabla_3);
                _uiManager.LimpiarLabels(": :", fee, reg, marc, biblioteca, codebiblio);

                fichas.Enabled = usuarios.Enabled = ejemplares.Enabled = llenado.Enabled = archivosDeSalidaToolStripMenuItem.Enabled = archivosDeSalidaToolStripMenuItem.Visible = split.Enabled = true;
                fee.Visible = reg.Visible = marc.Visible = split.Visible = biblioteca.Visible = codebiblio.Visible = exusers.Visible = false;

                _dialogService.MostrarInformacion("Recuerde seleccionar la ruta de salida en el menu desplegable superior (Archivo de salida)");
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

        private void PathMarc()
        {
            while (string.IsNullOrEmpty(_config.MarcEditPath))
            {
                string rutaMarcEdit = _dialogService.SeleccionarMarcEdit();
                
                if (!string.IsNullOrEmpty(rutaMarcEdit))
                {
                    _config.MarcEditPath = rutaMarcEdit;
                    _dialogService.MostrarInformacion($"Herramienta en uso: {_config.MarcEditPath}");
                }
                else
                {
                    _dialogService.MostrarAdvertencia("Por favor, seleccione el archivo ejecutable");
                }
            }
        }

        private void seleccionarRutaDeSalidaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PathRutadeSalida();
        }

        private void PathRutadeSalida()
        {
            string carpetaSalida = _dialogService.SeleccionarCarpeta("Seleccionar Carpeta de Salida");
            
            if (!string.IsNullOrEmpty(carpetaSalida))
            {
                _config.OutputFolder = carpetaSalida;
                _dialogService.MostrarInformacion($"Ruta de salida seleccionada: {_config.OutputFolder}");
            }
            else
            {
                _dialogService.MostrarAdvertencia("No se seleccionó ninguna carpeta.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _dialogService.MostrarInformacion("Bienvenido al Sistema de Migracion de Datos SIABUC8 => Koha");

            _dialogService.MostrarInformacion("Para comenzar, por favor busque y seleccione la herramienta necesaria MarcEdit en su sistema");
            PathMarc();

            _dialogService.MostrarInformacion("Luego, busque el archivo Access (SIABUC8.mdb) y seleccionelo en el menu superior (Archivo de entrada)");

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

        private void seleccionarArchivosMRK_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            _uiManager.SetControlsEnabled(this, false);
            
            try
            {
                _marcFileOps.ConvertirMRKaMRC(_config.MarcEditPath);
            }
            finally
            {
                Cursor = Cursors.Default;
                _uiManager.SetControlsEnabled(this, true);
            }
        }

        private void selMRC_Click(object sender, EventArgs e)
        {
            _marcFileOps.UnirArchivosMRC(_config.MarcEditPath);
        }

        private void selMRK_Click(object sender, EventArgs e)
        {
            _marcFileOps.UnirArchivosMRK(_config.MarcEditPath);
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
            extraer.CodigoBiblio();
            extraer.Extraer(_config.ConnectionString);
        }
    }
}
