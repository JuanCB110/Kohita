using System;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

class EUsuarios
{
    private static Form formSinEjemplares = null;
    public static string codigo = string.Empty;

    public void Extraer(string connectionString)
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

                MessageBox.Show($"Exportación completada.\nArchivo guardado en:\n{outputPath}");
            }
        }
    }

    public void CodigoBiblio()
    {
        // Si ya hay una ventana abierta, ciérrala
        if (formSinEjemplares != null && !formSinEjemplares.IsDisposed)
        {
            formSinEjemplares.Close();
        }

        // Crear una nueva ventana en tiempo de ejecución
        formSinEjemplares = new Form
        {
            Text = "Código de Biblioteca",
            Size = new Size(400, 200),
            StartPosition = FormStartPosition.CenterScreen,
            Icon = new Icon(Assembly.GetExecutingAssembly().GetManifestResourceStream("ExtraccionA.zombie-removebg-preview.ico"))  // Acceder al recurso incrustado
        };

        // Crear un Label para mostrar el mensaje
        Label labelCodigo = new Label
        {
            Text = "Ingrese el código de la biblioteca:",
            Location = new Point(20, 30),
            Size = new Size(250, 30)
        };

        // Crear un TextBox para ingresar el código de la biblioteca
        TextBox textBoxCodigo = new TextBox
        {
            Location = new Point(20, 70),
            Size = new Size(200, 25)
        };

        // Crear un botón de Aceptar
        Button buttonAceptar = new Button
        {
            Text = "Aceptar",
            Location = new Point(20, 110),
            Size = new Size(100, 30)
        };
        buttonAceptar.Click += (sender, e) =>
        {
            // Aquí podrías manejar lo que sucede cuando el usuario hace clic en Aceptar
            codigo = textBoxCodigo.Text;
            // Llamar a la función que maneja el código ingresado
            MessageBox.Show($"Código de la biblioteca ingresado: {codigo}");
            formSinEjemplares.Close();
        };

        // Crear un botón de Cancelar
        Button buttonCancelar = new Button
        {
            Text = "Cancelar",
            Location = new Point(130, 110),
            Size = new Size(100, 30)
        };
        buttonCancelar.Click += (sender, e) =>
        {
            // Cerrar la ventana si el usuario hace clic en Cancelar
            formSinEjemplares.Close();
        };

        // Agregar los controles al formulario
        formSinEjemplares.Controls.Add(labelCodigo);
        formSinEjemplares.Controls.Add(textBoxCodigo);
        formSinEjemplares.Controls.Add(buttonAceptar);
        formSinEjemplares.Controls.Add(buttonCancelar);

        // Mostrar la ventana
        formSinEjemplares.ShowDialog();
    }

}
