using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para manejar diálogos de archivos y carpetas
    /// </summary>
    public class DialogService
    {
        /// <summary>
        /// Muestra un diálogo para seleccionar un archivo Access (.mdb)
        /// </summary>
        public string SeleccionarArchivoAccess()
        {
            using (OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Archivos de Base de Datos Access (*.mdb)|*.mdb|Todos los Archivos (*.*)|*.*",
                Title = "Seleccionar Archivo MDB"
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Muestra un diálogo para seleccionar el ejecutable de MarcEdit
        /// </summary>
        public string SeleccionarMarcEdit()
        {
            using (OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Herramienta MarcEdit cmarcedit (*.exe)|*.exe|Todos los Archivos (*.*)|*.*",
                Title = "Seleccionar el Archivo Ejecutable EXE"
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Muestra un diálogo para seleccionar una carpeta
        /// </summary>
        public string SeleccionarCarpeta(string descripcion = "Seleccionar Carpeta")
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog
            {
                Description = descripcion
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.SelectedPath : null;
            }
        }

        /// <summary>
        /// Muestra un diálogo para guardar un archivo CSV
        /// </summary>
        public string SeleccionarArchivoGuardarCSV(string nombrePorDefecto = "usuarios_koha.csv")
        {
            using (SaveFileDialog dialog = new SaveFileDialog
            {
                Filter = "Archivo CSV (*.csv)|*.csv",
                Title = "Guardar Archivo CSV",
                FileName = nombrePorDefecto
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Muestra un diálogo para guardar un archivo MRC
        /// </summary>
        public string SeleccionarArchivoGuardarMRC(string nombrePorDefecto = "Archivo_Unificado.mrc")
        {
            using (SaveFileDialog dialog = new SaveFileDialog
            {
                Filter = "Archivo MRC (*.mrc)|*.mrc",
                Title = "Guardar Archivo Unido",
                FileName = nombrePorDefecto
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Muestra un diálogo para guardar un archivo MRK
        /// </summary>
        public string SeleccionarArchivoGuardarMRK(string nombrePorDefecto = "Archivo_Unificado.mrk")
        {
            using (SaveFileDialog dialog = new SaveFileDialog
            {
                Filter = "Archivo MRK (*.mrk)|*.mrk",
                Title = "Guardar Archivo Unido",
                FileName = nombrePorDefecto
            })
            {
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>
        /// Muestra un mensaje de información
        /// </summary>
        public void MostrarInformacion(string mensaje, string titulo = "Información")
        {
            MessageBox.Show(mensaje, titulo, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Muestra un mensaje de error
        /// </summary>
        public void MostrarError(string mensaje, string titulo = "Error")
        {
            MessageBox.Show(mensaje, titulo, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// Muestra un mensaje de advertencia
        /// </summary>
        public void MostrarAdvertencia(string mensaje, string titulo = "Advertencia")
        {
            MessageBox.Show(mensaje, titulo, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Muestra un mensaje de confirmación
        /// </summary>
        public bool MostrarConfirmacion(string mensaje, string titulo = "Confirmación")
        {
            return MessageBox.Show(mensaje, titulo, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
        }
    }
}
