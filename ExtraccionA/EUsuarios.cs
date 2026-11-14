using System;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

class EUsuarios
{
    /// <summary>
    /// Exporta la tabla Usuarios a un CSV para Koha con el código de biblioteca indicado (branchcode)
    /// </summary>
    public void Extraer(string connectionString, string codigo)
    {
        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
        {
            saveFileDialog.Filter = "Archivo CSV (*.csv)|*.csv";
            saveFileDialog.Title = "Guardar Archivo CSV";
            saveFileDialog.FileName = "usuarios_koha.csv";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string outputPath = saveFileDialog.FileName;

                string query = "SELECT NoCuenta, Nombre, NoGrupo, NoEscuela, [E-mail], Domicilio, Colonia, [Ciudad Estado], [Codigo Postal], Telefono, Notas FROM Usuarios";

                try
                {
                    EscribirCsv(connectionString, query, outputPath, codigo);
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"No se pudo escribir el archivo:\n{outputPath}\n\n{ex.Message}\n\nSi el archivo está abierto en Excel u otro programa, ciérrelo e intente de nuevo.",
                        "Error al exportar usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (UnauthorizedAccessException ex)
                {
                    MessageBox.Show($"No tiene permiso para escribir en:\n{outputPath}\n\n{ex.Message}",
                        "Error al exportar usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ocurrió un error al exportar los usuarios:\n{ex.Message}",
                        "Error al exportar usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                MessageBox.Show($"Exportación completada.\nArchivo guardado en:\n{outputPath}");
            }
        }
    }

    /// <summary>
    /// Lee la tabla Usuarios y escribe el CSV para el importador de usuarios de Koha
    /// </summary>
    private void EscribirCsv(string connectionString, string query, string outputPath, string codigo)
    {
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            connection.Open();
            using (OleDbCommand command = new OleDbCommand(query, connection))
            using (OleDbDataReader reader = command.ExecuteReader())
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Escribir la cabecera del CSV
                writer.WriteLine("cardnumber,surname,firstname,email,address,city,zipcode,phone,categorycode,branchcode,userid,password");

                while (reader.Read())
                {
                    string cardnumber = reader["NoCuenta"].ToString().Replace(",", "");
                    string fullName = reader["Nombre"].ToString().Replace(",", "");
                    string email = reader["E-mail"].ToString().Replace(",", "");
                    string address = reader["Domicilio"].ToString().Replace(",", "");
                    string city = reader["Ciudad Estado"].ToString().Replace(",", "");
                    string zipcode = reader["Codigo Postal"].ToString().Replace(",", "");
                    string phone = reader["Telefono"].ToString().Replace(",", "");

                    // Separar el nombre en nombre y apellido (esto depende del formato de los datos)
                    string[] nameParts = fullName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                    // Asignar 'firstname' y 'surname' según la longitud del arreglo
                    string firstname = string.Empty;
                    string surname = string.Empty;

                    // Validar que tenemos al menos 2 elementos para los apellidos
                    if (nameParts.Length >= 2)
                    {
                        surname = nameParts[0] + " " + nameParts[1];  // Primer y segundo apellido

                        // El resto son los nombres
                        firstname = string.Join(" ", nameParts, 2, nameParts.Length - 2);
                    }

                    // Definir valores fijos o inferidos
                    string categorycode = "ST";  // Código de usuario en Koha (por ejemplo, estudiante)
                    string branchcode = codigo;  // Biblioteca principal
                    string userid = cardnumber;  // Puede ser igual al número de cuenta
                    string password = "123";    // Contraseña inicial por defecto

                    writer.WriteLine($"{cardnumber},{surname},{firstname},{email},{address},{city},{zipcode},{phone},{categorycode},{branchcode},{userid},{password}");
                }
            }
        }
    }

    /// <summary>
    /// Pide el código de biblioteca, ya escrito con el de la biblioteca elegida al abrir la base.
    /// Devuelve el código sin espacios sobrantes, o null si el usuario cancela o cierra la ventana.
    /// </summary>
    public string CodigoBiblio(string codigoInicial)
    {
        using (Form ventana = new Form
        {
            Text = "Código de Biblioteca",
            Size = new Size(400, 200),
            StartPosition = FormStartPosition.CenterScreen
        })
        {
            try
            {
                ventana.Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExtraccionA.zombie-removebg-preview.ico"));
            }
            catch
            {
                // Si no se puede cargar el icono, continuar sin él
            }

            Label labelCodigo = new Label
            {
                Text = "Ingrese el código de la biblioteca:",
                Location = new Point(20, 30),
                Size = new Size(250, 30)
            };

            TextBox textBoxCodigo = new TextBox
            {
                Location = new Point(20, 70),
                Size = new Size(200, 25),
                Text = codigoInicial ?? string.Empty
            };

            Button buttonAceptar = new Button
            {
                Text = "Aceptar",
                Location = new Point(20, 110),
                Size = new Size(100, 30)
            };
            buttonAceptar.Click += (sender, e) =>
            {
                // Sin código, Koha rechaza a los usuarios: no se deja continuar con la caja vacía
                if (textBoxCodigo.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Escriba el código de la biblioteca.", "Código de Biblioteca",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBoxCodigo.Focus();
                    return;
                }
                ventana.DialogResult = DialogResult.OK;
            };

            Button buttonCancelar = new Button
            {
                Text = "Cancelar",
                Location = new Point(130, 110),
                Size = new Size(100, 30),
                DialogResult = DialogResult.Cancel
            };

            ventana.Controls.Add(labelCodigo);
            ventana.Controls.Add(textBoxCodigo);
            ventana.Controls.Add(buttonAceptar);
            ventana.Controls.Add(buttonCancelar);

            // Enter = Aceptar, Esc = Cancelar
            ventana.AcceptButton = buttonAceptar;
            ventana.CancelButton = buttonCancelar;

            // Cancelar o cerrar con la X devuelven Cancel: no se exporta nada
            return ventana.ShowDialog() == DialogResult.OK ? textBoxCodigo.Text.Trim() : null;
        }
    }

}
