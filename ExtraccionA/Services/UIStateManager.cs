using System.Collections.Generic;
using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para manejar el estado y visibilidad de controles de UI
    /// </summary>
    public class UIStateManager
    {
        /// <summary>
        /// Habilita o deshabilita todos los controles de un formulario
        /// </summary>
        public void SetControlsEnabled(Form form, bool enabled)
        {
            foreach (Control control in form.Controls)
            {
                control.Enabled = enabled;
            }
        }

        /// <summary>
        /// Actualiza la visibilidad y habilitación de múltiples controles
        /// </summary>
        public void ActualizarEstadoControles(params (Control control, bool visible, bool enabled)[] estados)
        {
            foreach (var (control, visible, enabled) in estados)
            {
                control.Visible = visible;
                control.Enabled = enabled;
            }
        }

        /// <summary>
        /// Limpia el contenido de múltiples labels
        /// </summary>
        public void LimpiarLabels(string textoDefault, params Label[] labels)
        {
            foreach (var label in labels)
            {
                label.Text = textoDefault;
                label.Visible = false;
            }
        }

        /// <summary>
        /// Limpia los DataSource de múltiples DataGridViews
        /// </summary>
        public void LimpiarDataGridViews(params DataGridView[] grids)
        {
            foreach (var grid in grids)
            {
                grid.DataSource = null;
            }
        }

        /// <summary>
        /// Configura el estado de botones mutuamente excluyentes
        /// </summary>
        public void SetButtonState(Button activeButton, params Button[] allButtons)
        {
            foreach (var button in allButtons)
            {
                button.Enabled = (button == activeButton);
            }
        }

        /// <summary>
        /// Muestra u oculta un grupo de controles relacionados
        /// </summary>
        public void MostrarGrupoControles(bool mostrar, params Control[] controles)
        {
            foreach (var control in controles)
            {
                control.Visible = mostrar;
                control.Enabled = mostrar;
            }
        }

        /// <summary>
        /// Actualiza un label con el número de fichas únicas seleccionadas en un DataGridView
        /// </summary>
        public void ActualizarLabelFichasSeleccionadas(DataGridView grid, Label label, params Control[] controlesAsociados)
        {
            var selectedRows = grid.SelectedRows;

            if (selectedRows.Count == 0)
            {
                MostrarGrupoControles(false, controlesAsociados);
                return;
            }

            // Usar HashSet para almacenar solo números de ficha únicos
            HashSet<string> uniqueFichaNumbers = new HashSet<string>();

            foreach (DataGridViewRow row in selectedRows)
            {
                var fichaNumber = row.Cells[0].Value?.ToString();

                if (!string.IsNullOrEmpty(fichaNumber))
                {
                    uniqueFichaNumbers.Add(fichaNumber);
                }
            }

            label.Text = "Fichas seleccionadas: " + uniqueFichaNumbers.Count;
        }
    }
}

