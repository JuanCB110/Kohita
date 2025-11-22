using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para gestionar la selección y configuración de bibliotecas
    /// </summary>
    public class LibraryManagerService
    {
        private readonly DialogService _dialogService;
        private readonly ValidationService _validationService;

        public LibraryManagerService(DialogService dialogService, ValidationService validationService)
        {
            _dialogService = dialogService;
            _validationService = validationService;
        }

        /// <summary>
        /// Carga las bibliotecas desde la base de datos y permite seleccionar una
        /// </summary>
        /// <param name="connectionString">Cadena de conexión a la base de datos</param>
        /// <param name="bibliotecaLabel">Label para mostrar el nombre de la biblioteca</param>
        /// <param name="codigoBiblioLabel">Label para mostrar el código de la biblioteca</param>
        public void CargarYSeleccionarBiblioteca(string connectionString, Label bibliotecaLabel, Label codigoBiblioLabel)
        {
            bibliotecaLabel.Visible = codigoBiblioLabel.Visible = true;
            string query = "SELECT * FROM Bibliotecas";

            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();

                    using (OleDbCommand cmd = new OleDbCommand(query, conn))
                    {
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.HasRows)
                            {
                                List<string> bibliotecas = new List<string>();
                                List<string> codigos = new List<string>();

                                while (reader.Read())
                                {
                                    string nombre = reader["Nombre"].ToString();
                                    string codigo = reader["No"].ToString();
                                    bibliotecas.Add(nombre);
                                    codigos.Add(codigo);
                                }

                                if (bibliotecas.Count > 1)
                                {
                                    MostrarDialogoSeleccion(bibliotecas, codigos, bibliotecaLabel, codigoBiblioLabel);
                                }
                                else if (bibliotecas.Count == 1)
                                {
                                    AsignarBiblioteca(bibliotecas[0], codigos[0], bibliotecaLabel, codigoBiblioLabel);
                                }
                            }
                            else
                            {
                                _dialogService.MostrarAdvertencia("No se encontraron resultados.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _dialogService.MostrarError($"Error al cargar bibliotecas: {ex.Message}");
            }
        }

        /// <summary>
        /// Muestra un diálogo para seleccionar una biblioteca cuando hay múltiples opciones
        /// </summary>
        private void MostrarDialogoSeleccion(List<string> bibliotecas, List<string> codigos, Label bibliotecaLabel, Label codigoBiblioLabel)
        {
            using (Form selectForm = new Form())
            {
                selectForm.Text = "Selecciona una biblioteca";
                selectForm.Size = new Size(400, 100);
                selectForm.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                selectForm.FormBorderStyle = FormBorderStyle.FixedDialog;
                selectForm.StartPosition = FormStartPosition.CenterScreen;

                // Cargar el icono
                try
                {
                    byte[] iconBytes = Properties.Resources.zombie;
                    using (MemoryStream ms = new MemoryStream(iconBytes))
                    {
                        selectForm.Icon = new Icon(ms);
                    }
                }
                catch
                {
                    // Si no se puede cargar el icono, continuar sin él
                }

                // Crear ComboBox
                ComboBox comboBiblio = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Dock = DockStyle.Fill,
                    DataSource = bibliotecas
                };
                selectForm.Controls.Add(comboBiblio);

                // Botón Aceptar
                Button btnAceptar = new Button
                {
                    Text = "Aceptar",
                    Dock = DockStyle.Bottom
                };
                
                btnAceptar.Click += (sender, e) =>
                {
                    int selectedIndex = comboBiblio.SelectedIndex;
                    AsignarBiblioteca(bibliotecas[selectedIndex], codigos[selectedIndex], bibliotecaLabel, codigoBiblioLabel);
                    selectForm.DialogResult = DialogResult.OK;
                    selectForm.Close();
                };
                
                selectForm.Controls.Add(btnAceptar);
                selectForm.ShowDialog();
            }
        }

        /// <summary>
        /// Asigna la biblioteca y código seleccionado a los labels
        /// </summary>
        private void AsignarBiblioteca(string nombreBiblioteca, string codigo, Label bibliotecaLabel, Label codigoBiblioLabel)
        {
            bibliotecaLabel.Text = "BIBLIOTECA: " + nombreBiblioteca;
            string numeroFormateado = _validationService.FormatearCodigoBiblioteca(codigo);
            codigoBiblioLabel.Text = "NUMERO: " + numeroFormateado;
        }
    }
}
