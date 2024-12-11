using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtraccionA
{
    public partial class Form1 : Form
    {
        // Declaramos una variable que almacenará la cadena de conexión dinámica
        private string connString = string.Empty;
        private string folderout = string.Empty;
        private string codigobiblio = string.Empty;
        private string codei = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        //Primera Tabla
        private void fichas_Click(object sender, EventArgs e) => FetchData("SELECT * FROM Fichas", fichas) ;
        private void editorial_Click(object sender, EventArgs e) => FetchData("SELECT * FROM Editoriales", editorial);
        private void ejemplares_Click(object sender, EventArgs e) => FetchData("SELECT * FROM EJEMPLARES", ejemplares);
        private async void FetchData(string query, Button button)
        {
            Cursor = Cursors.WaitCursor;
            // Deshabilitar los controles
            SetControlsEnabled(false);
            tabla_1.ReadOnly = true;
            SetButtonState(button);
            await FillDataWithoutProgress(query, tabla_1);
            limpiar.Visible = limpiar.Enabled = true;
            fee.Visible = true;
            fee.Text = $"{button.Text} encontradas: " + tabla_1.RowCount;
            Cursor = Cursors.Default;
            // Volver a habilitar los controles y cambiar el cursor
            SetControlsEnabled(true);
        }

        // Método para habilitar o deshabilitar los controles del formulario
        private void SetControlsEnabled(bool enabled)
        {
            foreach (Control control in this.Controls)
            {
                control.Enabled = enabled;
            }
        }

        //Segunda Tabla
        private async void Llenado_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            SetControlsEnabled(false);
            string query = @"SELECT 
                                f.Ficha_No AS Numero_De_Control,
                                e.NumAdqui AS Numero_de_Item,
                                f.Fecha AS Fecha_Entrada_Ficha, 
                                f.FechaMod AS Fecha_Modificacion_Ficha, 
                                f.DatosFijos,
                                f.ISBN,
                                f.Clasificacion AS Clasificacion_LC, 
                                f.TipoMaterial,  
                                f.Titulo, 
                                f.Autor,  
                                f.Estatus,  
                                e.FechaIngreso AS Fecha_Ingreso, 
                                e.Biblioteca, 
                                e.Volumen, 
                                e.Ejemplar, 
                                e.Tomo, 
                                e.Accesible, 
                                e.NoEscuela, 
                                e.FechaMod AS Fecha_Modificacion_Ejemplar, 
                                e.Analista,
                                f.EtiquetasMARC AS MARC_Completo
                            FROM Fichas f
                            INNER JOIN EJEMPLARES e ON f.Ficha_No = e.Ficha_No
                            ORDER BY f.Ficha_No, e.NumAdqui;";
            await FillDataWithoutProgress(query, tabla_2);
            split.Visible = split.Enabled = reg.Visible = true;
            reg.Text = "Registros encontrados: " + tabla_2.RowCount;
            Cursor = Cursors.Default;
            SetControlsEnabled(true);
        }

        private async Task FillDataWithoutProgress(string query, DataGridView gridView)
        {
            try
            {
                using (var conn = new OleDbConnection(connString))
                {
                    await conn.OpenAsync();

                    using (var cmd = new OleDbCommand(query, conn))
                    {
                        // Intentamos ejecutar la consulta
                        using (var reader = await cmd.ExecuteReaderAsync())
                        {
                            // Verificamos si el lector tiene filas
                            if (!reader.HasRows)
                            {
                                MessageBox.Show("No se encontraron resultados para la consulta.");
                                return;
                            }

                            // Crear un DataTable e ir llenando a medida que leemos las filas
                            DataTable dataTable = new DataTable();
                            dataTable.Load(reader);

                            // Asignar el DataTable al DataGridView
                            gridView.DataSource = dataTable;
                            gridView.AutoResizeColumns();
                            gridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                        }
                    }
                }
            }
            catch (OleDbException oleDbEx)
            {
                // Captura errores específicos de la base de datos, como tablas inexistentes
                MessageBox.Show("Error de base de datos: " + oleDbEx.Message);
            }
            catch (Exception ex)
            {
                // Captura errores generales
                MessageBox.Show("Error al conectar con la base de datos: " + ex.Message);
            }
        }

        private void codigos(codigo cd)
        {
            string queryc = "SELECT * FROM Bibliotecas";

            try
            {
                // Crear la conexión
                using (OleDbConnection conn = new OleDbConnection(connString))
                {
                    conn.OpenAsync();

                    // Crear el comando para ejecutar el query
                    using (OleDbCommand cmd = new OleDbCommand(queryc, conn))
                    {
                        // Ejecutar el comando y obtener el reader para leer los resultados
                        using (OleDbDataReader reader = cmd.ExecuteReader())
                        {
                            // Verificar si hay datos
                            if (reader.HasRows)
                            {
                                while (reader.Read())
                                {
                                    // Aquí suponemos que tienes las columnas "Nombre" y "No" en tu tabla
                                    string nombre = reader["Nombre"].ToString();
                                    string code = reader["No"].ToString();

                                    // Llamar al método para actualizar los controles con los nuevos datos
                                    cd.datos(nombre, code);
                                }
                            }
                            else
                            {
                                MessageBox.Show("No se encontraron resultados.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Manejo de excepciones
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void SetButtonState(Button activeButton)
        {
            fichas.Enabled = activeButton == fichas;
            editorial.Enabled = activeButton == editorial;
            ejemplares.Enabled = activeButton == ejemplares;
        }

        private void limpiar_Click(object sender, EventArgs e)
        {
            tabla_1.ReadOnly = false;
            tabla_1.DataSource = null;
            limpiar.Visible = limpiar.Enabled = false;
            fichas.Enabled = editorial.Enabled = ejemplares.Enabled = true;
            fee.Visible = false;
            fee.Text = ": :";
        }

        private void split_Click(object sender, EventArgs e)
        {
            // Crear una instancia del formulario 'codigo'
            codigo cd = new codigo();
            codigos(cd);

            // Mostrar el formulario de entrada y esperar a que se cierre (ShowDialog bloquea la ejecución)
            if (cd.ShowDialog() == DialogResult.OK) // Aseguramos que solo continuemos si el usuario hace clic en "Aceptar"
            {
                // Ahora los valores ya están disponibles en las propiedades Nom y Codi
                string nombrebiblio = cd.Nom; // No es necesario llamar a ToString()
                codigobiblio = cd.Codi;

                // Verificar si los valores no están vacíos
                if (!string.IsNullOrEmpty(nombrebiblio) && !string.IsNullOrEmpty(cd.Codi))
                {
                    // Proceder con el procesamiento como antes
                    Cursor = Cursors.WaitCursor;
                    SetControlsEnabled(false);
                    SplitTable(nombrebiblio);
                    Cursor = Cursors.Default;
                    SetControlsEnabled(true);
                    marc.Visible = true;
                    marc.Text = "Registros formateados: " + tabla_3.RowCount;
                }
                else
                {
                    MessageBox.Show("Por favor, ingresa un nombre y un código.");
                }
            }
            else
            {
                // Si el usuario cancela, puedes manejarlo si es necesario
                MessageBox.Show("Operación cancelada.");
            }
        }

        private void SplitTable(string nombrebiblio)
        {
            DataTable originalTable = (DataTable)tabla_2.DataSource;
            if (originalTable == null) return;

            DataTable processedTable = new DataTable();
            processedTable.Columns.Add("003", typeof(string)); // Identificador de control
            processedTable.Columns.Add("005", typeof(string)); // Fecha y hora de la ultima modificacion
            processedTable.Columns.Add("Numeros_de_Items", typeof(string)); // Agrega la columna "Numero_de_Item"
            processedTable.Columns.Add("008", typeof(string)); // Datos fijos
            processedTable.Columns.Add("040", typeof(string)); // Fuente de catalogo
            processedTable.Columns.Add("942", typeof(string)); // Tipo de material

            string fechaActual = DateTime.Now.ToString("yyyyMMddHHmmss");

            // Usa un diccionario para agrupar por "Numero_De_Control"
            var groupedData = originalTable.AsEnumerable()
                .GroupBy(row => row["Numero_De_Control"].ToString())
                .Select(group => new
                {
                    // Si el "Numero_De_Control" es nulo o vacío, se asigna un valor nulo
                    NumeroDeControl = string.IsNullOrEmpty(group.Key) ? null : group.Key,

                    // Concatenar y evitar duplicados, pero también considerar los vacíos
                    NumeroDeItems = string.Join(",", group
                        .Select(row => row["Numero_de_Item"]?.ToString()) // Usamos el operador de seguridad para evitar nulos
                        .Where(item => !string.IsNullOrEmpty(item)) // Filtramos vacíos
                        .Distinct()),

                    // Convertir a string y dejar que sea null si no hay valor
                    DatosFijos = group.FirstOrDefault()?["DatosFijos"]?.ToString(), // Usamos ?.ToString() para convertir a string, si es null, permanece null

                    // Convertir a string y dejar que sea null si no hay valor
                    MarcCompleto = group.FirstOrDefault()?["MARC_Completo"]?.ToString(), // Igual para MarcCompleto

                    //Pasar el tipo de material
                    TipoMaterial = group.FirstOrDefault()?["TipoMaterial"]?.ToString()
                }).ToList();


            var addedColumns = new HashSet<string>();

            // Procesamos cada grupo
            foreach (var group in groupedData)
            {
                DataRow newRow = processedTable.NewRow();

                // Asigna el "Número de Control" y "Numero_de_Item"
                newRow["003"] = group.NumeroDeControl;
                newRow["005"] = fechaActual;
                newRow["Numeros_de_Items"] = group.NumeroDeItems;
                newRow["008"] = group.DatosFijos.Replace(" ", "");
                newRow["040"] = nombrebiblio;
                switch (group.TipoMaterial)
                {
                    case "0":
                        newRow["942"] = codei = "OR";
                        break;
                    case "1":
                        newRow["942"] = codei = "BK";
                        break;
                    case "2":
                        newRow["942"] = codei = "TS";
                        break;
                    default:
                        break;
                }

                // Procesar "MARC_Completo" si existe
                var marcParts = group.MarcCompleto.Split('¦').Where(s => s.Length > 3)
                    .OrderBy(part => int.TryParse(part.Substring(0, 3), out var res) ? res : int.MaxValue);

                foreach (string part in marcParts)
                {
                    string etiqueta = part.Substring(0, 3);
                    string valor = part.Substring(3);

                    // Añadir la columna si no está ya presente
                    if (!addedColumns.Contains(etiqueta))
                    {
                        processedTable.Columns.Add(etiqueta, typeof(string));
                        addedColumns.Add(etiqueta);
                    }
                    newRow[etiqueta] = valor;
                }
                processedTable.Rows.Add(newRow);
            }
            // Asigna el DataTable procesado al DataGridView
            tabla_3.DataSource = processedTable;
            MessageBox.Show("Nota:\nSi los Registros Formateados son menores que las Fichas Encontradas, significa que esas fichas no contienen ejemplares");
            Console.WriteLine($"Filas procesadas: {processedTable.Rows.Count}");
        }

        private void format_Click(object sender, EventArgs e)
        {
            if (folderout != string.Empty)
            {
                FormatMARC(codigobiblio);
            } else
            {
                MessageBox.Show("Por favor, seleccione una ruta accesible en el menu posterior");
            }
        }

        private void FormatMARC(string codebiblio)
        {
            string fechaActual = DateTime.Now.ToString("yyyyMMddHHmmss");

            // Verifica si hay filas seleccionadas
            if (tabla_3.SelectedRows.Count > 0)
            {
                // StringBuilder para almacenar el contenido de todas las filas seleccionadas
                StringBuilder formattedText = new StringBuilder();

                int registrosProcesados = 0; // Contador de registros procesados
                int archivoNumero = 0; // Contador para los archivos (parte1, parte2, etc.)

                // HashSet para almacenar números de control ya procesados
                HashSet<string> processedControlNumbers = new HashSet<string>();
                List<string> columnasExcluidas = new List<string> { "Numeros_de_Items" };
                List<string> filas = new List<string>();

                // Recorre todas las filas seleccionadas
                foreach (DataGridViewRow selectedRow in tabla_3.SelectedRows)
                {
                    // Obtiene el número de control o identificador único, por ejemplo, de la primera celda
                    string controlNumber = selectedRow.Cells["003"].Value?.ToString();
                    string nitems = selectedRow.Cells["Numeros_de_Items"].Value?.ToString();

                    if (!string.IsNullOrEmpty(controlNumber))
                    {
                        processedControlNumbers.Add(controlNumber);
                        // Recorre cada celda en la fila actual y la agrega al texto formateado
                        foreach (DataGridViewCell cell in selectedRow.Cells)
                        {
                            // Obtiene el nombre de la columna actual
                            string nombreColumna = tabla_3.Columns[cell.ColumnIndex].HeaderText;

                            if (!columnasExcluidas.Contains(nombreColumna))
                            {
                                if (cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString()))
                                {
                                    Console.WriteLine(tabla_3.Columns[cell.ColumnIndex].HeaderText + " Vacia");
                                }
                                else
                                {
                                    switch (tabla_3.Columns[cell.ColumnIndex].HeaderText)
                                    {
                                        case "020":
                                            {

                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "040":
                                            {

                                                string c = "$c" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {c}");
                                                break;
                                            }
                                        case "050":
                                            {
                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "100":
                                            {

                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "110":
                                            {

                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "130":
                                            {

                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "240":
                                            {

                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "245":
                                            {
                                                string div = cell.Value.ToString();

                                                if (div.Contains(":"))
                                                {
                                                    string[] partes = div.Split(':');

                                                    string a = partes[0].Trim();
                                                    string b = partes[1].Trim();

                                                    if (b.Contains("/"))
                                                    {
                                                        string[] partes2 = b.Split('/');

                                                        b = partes2[0].Trim();
                                                        string c = " / $c" + partes2[1].Trim();

                                                        string com = "$a" + a + " : $b" + b + c;

                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {com}");
                                                    }
                                                    else
                                                    {
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{a} : $b{b}");
                                                    }
                                                }
                                                else if (div.Contains("/"))
                                                {
                                                    string[] partes = div.Split('/');

                                                    string a = "$a" + partes[0].Trim() + " / ";
                                                    string c = "$c" + partes[1].Trim();

                                                    string com = a + c;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {com}");
                                                }
                                                else
                                                {
                                                    string a = "$a" + div;
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                }
                                                break; // Salir del case "245"
                                            }
                                        case "250":
                                            {
                                                string div = cell.Value.ToString();

                                                if (div.Contains("/"))
                                                {
                                                    string[] partes = div.Split('/');

                                                    string a = "$a" + partes[0].Trim();
                                                    string b = " $b" + " / " + partes[0].Trim();

                                                    string com = a + b;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {com}");
                                                }
                                                else
                                                {
                                                    string a = "$a" + div;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                }
                                                break;
                                            }
                                        case "260":
                                            {
                                                string div = cell.Value.ToString();

                                                if (div.Contains(":"))
                                                {
                                                    string[] partes = div.Split(':');

                                                    string a = partes[0].Trim();
                                                    string b = partes[1].Trim();

                                                    string com = "$a" + a + " : " + "$b" + b;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {com}");
                                                }
                                                else
                                                {
                                                    string a = "$a" + div;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                }
                                                break;
                                            }
                                        case "300":
                                            {
                                                string div = cell.Value.ToString();

                                                if (div.Contains("/"))
                                                {
                                                    string[] partes = div.Split('/');

                                                    string a = "$a" + partes[0].Trim();
                                                    string c = " / $c" + partes[1].Trim();

                                                    string com = a + c;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {com}");
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{div}");
                                                }
                                                break;
                                            }
                                        case "440":
                                            {
                                                string div = cell.Value.ToString();

                                                if (div.Contains("/"))
                                                {
                                                    string[] partes = div.Split('/');

                                                    string a = partes[0].Trim();
                                                    string b = " / " + partes[1].Trim();

                                                    string com = a + b;

                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {com}");
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {div}");
                                                }
                                                break;
                                            }
                                        case "500":
                                            {
                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"=500  {a}");
                                                break;
                                            }
                                        case "502":
                                        case "503":
                                        case "504":
                                            {
                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"=504  {a}");
                                                break;
                                            }
                                        case "505":
                                            {
                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "600":
                                            {
                                                string div = cell.Value.ToString();

                                                // Dividir la cadena principal en partes
                                                string[] partes = Regex.Split(div, @"(?<=\.)\s*(?=\d+\.)");

                                                if (partes.Length > 1)
                                                {
                                                    // Iterar sobre cada parte
                                                    foreach (var parte in partes)
                                                    {
                                                        // Separar la parte en sus componentes
                                                        Console.WriteLine($"\nParte:\n{parte}");
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{parte}");
                                                    }
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{cell.Value}");
                                                }
                                                break;
                                            }
                                        case "610":
                                            {
                                                string div = cell.Value.ToString();

                                                // Dividir la cadena principal en partes
                                                string[] partes = Regex.Split(div, @"(?<=\.)\s*(?=\d+\.)");

                                                if (partes.Length > 1)
                                                {
                                                    // Iterar sobre cada parte
                                                    foreach (var parte in partes)
                                                    {
                                                        // Separar la parte en sus componentes
                                                        Console.WriteLine($"\nParte:\n{parte}");
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{parte}");
                                                    }
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{div}");
                                                }
                                                break;
                                            }
                                        case "650":
                                            {
                                                string div = cell.Value.ToString();

                                                // Dividir la cadena principal en partes
                                                string[] partes = Regex.Split(div, @"(?<=\.)\s*(?=\d+\.)");

                                                if (partes.Length > 1)
                                                {
                                                    // Iterar sobre cada parte
                                                    foreach (var parte in partes)
                                                    {
                                                        // Separar la parte en sus componentes
                                                        Console.WriteLine($"\nParte:\n{parte}");
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{parte}");
                                                    }
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{div}");
                                                }
                                                break;
                                            }
                                        case "651":
                                            {
                                                string div = cell.Value.ToString();

                                                // Dividir la cadena principal en partes
                                                string[] partes = Regex.Split(div, @"(?<=\.)\s*(?=\d+\.)");

                                                if (partes.Length > 1)
                                                {
                                                    // Iterar sobre cada parte
                                                    foreach (var parte in partes)
                                                    {
                                                        // Separar la parte en sus componentes
                                                        Console.WriteLine($"\nParte:\n{parte}");
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{parte}");
                                                    }
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{div}");
                                                }
                                                break;
                                            }
                                        case "700":
                                            {
                                                string div = cell.Value.ToString();

                                                // Dividir la cadena principal en partes
                                                string[] partes = Regex.Split(div, @"(?<=\\)\s*");

                                                if (partes.Length > 1)
                                                {
                                                    // Iterar sobre cada parte
                                                    foreach (var parte in partes)
                                                    {
                                                        // Separar la parte en sus componentes
                                                        Console.WriteLine($"\nParte:\n{parte}");
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{parte}");
                                                    }
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{div}");
                                                }
                                                break;
                                            }
                                        case "710":
                                            {
                                                string div = cell.Value.ToString();

                                                // Dividir la cadena principal en partes
                                                string[] partes = Regex.Split(div, @"(?<=\\)\s*");

                                                if (partes.Length > 1)
                                                {
                                                    // Iterar sobre cada parte
                                                    foreach (var parte in partes)
                                                    {
                                                        // Separar la parte en sus componentes
                                                        Console.WriteLine($"\nParte:\n{parte}");
                                                        formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{parte.Trim()}");
                                                    }
                                                }
                                                else
                                                {
                                                    formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  $a{div}");
                                                }
                                                break;
                                            }
                                        case "856":
                                            {
                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "490":
                                            {
                                                string a = "$a" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {a}");
                                                break;
                                            }
                                        case "942":
                                            {
                                                string c = "$c" + cell.Value.ToString();

                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {c}");
                                                break;
                                            }
                                        default:
                                            {
                                                formattedText.AppendLine($"={tabla_3.Columns[cell.ColumnIndex].HeaderText}  {cell.Value}"); // Mover aquí
                                                break; // Salir del caso por defecto
                                            }
                                    }
                                }
                            }
                        }
                        string[] items = nitems.Split(',');

                        foreach (var item in items)
                        {
                            // Convierte nitems a una cadena y la agrega al texto formateado
                            formattedText.AppendLine($"=952  $a{codebiblio}$b{codebiblio}$d{fechaActual}$i{item}$y{codei}");
                        }

                        formattedText.AppendLine(); // Añade una línea en blanco entre filas
                        formattedText.AppendLine();
                        registrosProcesados++;
                    }
                    // Cada 20 registros, guarda el archivo y reinicia el StringBuilder para el siguiente archivo

                    if (folderout != string.Empty)
                    {
                        if (registrosProcesados >= 20)
                        {
                            string filePath = $@"{folderout}\parte{archivoNumero + 1}.mrk";
                            File.WriteAllText(filePath, formattedText.ToString());

                            // Reinicia el StringBuilder para el siguiente archivo
                            formattedText.Clear();

                            // Incrementa el número de archivo y reinicia el contador de registros procesados
                            archivoNumero++;
                            registrosProcesados = 0;
                        }
                    }
                }

                // Guarda los registros restantes que no completaron 20 al final del proceso
                if (registrosProcesados > 0)
                {
                    // Guardar el último archivo con los registros restantes

                    if (folderout != string.Empty)
                    {
                        string filePath = $@"{folderout}\parte{archivoNumero + 1}.mrk";
                        File.WriteAllText(filePath, formattedText.ToString());
                        //filas.Add($"parte{archivoNumero}");
                        MessageBox.Show($"Archivos guardados correctamente en {archivoNumero + 1} partes.");
                    }
                } else
                {
                    MessageBox.Show($"Archivos guardados correctamente en {archivoNumero + 1} partes.");
                }
            }
            else
            {
                MessageBox.Show("Por favor, selecciona al menos una fila antes de formatear.");
                return;
            }
        }

        private void tabla_3_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            format.Enabled = format.Visible = true;
            selall.Enabled = selall.Visible = true;
            label2.Enabled = label2.Visible = true;
            regsel.Enabled = regsel.Visible = true;
        }

        private void tabla_3_SelectionChanged(object sender, EventArgs e)
        {
            // Actualizar el label con los índices de las filas seleccionadas
            UpdateLabelWithSelectedRows();
            regsel.Text = "Registros Seleccionados: " + tabla_3.SelectedRows.Count.ToString();
        }

        private void UpdateLabelWithSelectedRows()
        {
            // Obtén las filas seleccionadas
            var selectedRows = tabla_3.SelectedRows;

            // Si no hay filas seleccionadas, limpia el label
            if (selectedRows.Count == 0)
            {
                format.Enabled = format.Visible = false;
                selall.Enabled = selall.Visible = false;
                label2.Enabled = label2.Visible = false;
                regsel.Enabled = regsel.Visible = false;
                return;
            }

            // Usamos un HashSet para almacenar solo los números de ficha únicos
            HashSet<string> uniqueFichaNumbers = new HashSet<string>();

            // Iterar sobre las filas seleccionadas
            foreach (DataGridViewRow row in selectedRows)
            {
                // Asegúrate de que la celda que contiene el número de ficha no sea nula
                var fichaNumber = row.Cells[0].Value?.ToString();

                // Si el número de ficha es válido, añádelo al HashSet
                if (!string.IsNullOrEmpty(fichaNumber))
                {
                    uniqueFichaNumbers.Add(fichaNumber);
                }
            }

            // Actualizar el Label con el número de fichas únicas seleccionadas
            label2.Text = "Fichas seleccionadas: " + uniqueFichaNumbers.Count;
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
            // Crear un nuevo OpenFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // Establecer el filtro para solo mostrar archivos .mdb
            openFileDialog.Filter = "Archivos de Base de Datos Access (*.mdb)|*.mdb|Todos los Archivos (*.*)|*.*";

            // Establecer el título del cuadro de diálogo
            openFileDialog.Title = "Seleccionar Archivo MDB";

            // Verificar si el usuario selecciona un archivo y presiona Aceptar
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Obtener la ruta completa del archivo seleccionado
                string archivoSeleccionado = openFileDialog.FileName;

                // Actualizar la cadena de conexión con la ruta seleccionada
                connString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={archivoSeleccionado};Persist Security Info=False;";

                // Aquí puedes usar la cadena de conexión para abrir la base de datos
                MessageBox.Show($"Conexión establecida con la base de datos: {archivoSeleccionado}");

                //Tablas
                tabla_1.DataSource = tabla_2.DataSource = tabla_3.DataSource = null;

                //Labels
                fee.Text = reg.Text = marc.Text = ": :";
                fee.Visible = reg.Visible = marc.Visible = split.Visible = false;

                fichas.Enabled = editorial.Enabled = ejemplares.Enabled = llenado.Enabled = archivosDeSalidaToolStripMenuItem.Enabled = archivosDeSalidaToolStripMenuItem.Visible = split.Enabled = true;

                //Pedir la ruta de salida para guardar las partes de los datos
                MessageBox.Show("Recuerde seleccionar la ruta de salida en el menu desplegable superior (Archivo de salida)");
            }
            else
            {
                MessageBox.Show("No se seleccionó ningún archivo.");
            }
        }

        private void seleccionarRutaDeSalidaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Crear un nuevo FolderBrowserDialog
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // Establecer el título del cuadro de diálogo
            folderBrowserDialog.Description = "Seleccionar Carpeta de Salida";

            // Mostrar el diálogo y verificar si el usuario selecciona una carpeta
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                // Obtener la ruta de la carpeta seleccionada
                folderout = folderBrowserDialog.SelectedPath;

                // Mostrar la ruta seleccionada
                MessageBox.Show($"Ruta de salida seleccionada: {folderout}");
            }
            else
            {
                MessageBox.Show("No se seleccionó ninguna carpeta.");
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            MessageBox.Show("Bienvenido al Sistema de Migracion de Datos SIABUC8 => Koha");
            MessageBox.Show("Para comenzar, busque el archivo Access (SIABUC8.mdb) y seleccionelo en el menu superior (Archivo de entrada)");

            toolTip1.SetToolTip(tabla1, "Haz click para mas informacion.");
            toolTip2.SetToolTip(tabla2, "Haz click para mas informacion.");
            toolTip3.SetToolTip(tabla3, "Haz click para mas informacion.");
        }

        private void tabla1_Click(object sender, EventArgs e)
        {
            // Mostrar una ventana de información al hacer clic
            MessageBox.Show("Información directa del Access, tabla de fichas, ejemplares y editoriales.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void tabla2_Click(object sender, EventArgs e)
        {
            // Mostrar una ventana de información al hacer clic
            MessageBox.Show("Tabla de informacion procesada para formato MARC.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void tabla3_Click(object sender, EventArgs e)
        {
            // Mostrar una ventana de información al hacer clic
            MessageBox.Show("Tabla formateada para exportacion por formato MARC para posterior conversion.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void seleccionarArchivosMRK_Click(object sender, EventArgs e)
        {
            // Crear un nuevo FolderBrowserDialog
            FolderBrowserDialog folderBrowserDialogcn = new FolderBrowserDialog();
            FolderBrowserDialog folderBrowserDialogcnsalida = new FolderBrowserDialog();

            // Establecer el título del cuadro de diálogo
            folderBrowserDialogcn.Description = "Seleccionar Carpeta de Archivos MRK";
            folderBrowserDialogcnsalida.Description = "Seleccionar Carpeta de Salida";

            Convertirmrc cnmrc = new Convertirmrc();

            // Mostrar el diálogo y verificar si el usuario selecciona una carpeta
            if (folderBrowserDialogcn.ShowDialog() == DialogResult.OK)
            {
                // Obtener la ruta de la carpeta seleccionada
                cnmrc.UbicacionArchivosMRK = folderBrowserDialogcn.SelectedPath;

                // Mostrar la ruta seleccionada
                MessageBox.Show($"Ruta seleccionada: {folderBrowserDialogcn.SelectedPath}");

                if (folderBrowserDialogcnsalida.ShowDialog() == DialogResult.OK)
                {
                    cnmrc.UbicacionSalida = folderBrowserDialogcnsalida.SelectedPath;

                    MessageBox.Show($"Ruta seleccionada: {folderBrowserDialogcnsalida.SelectedPath}");

                    Cursor = Cursors.WaitCursor;
                    SetControlsEnabled(false);
                    cnmrc.ConvertirMRKaMRC();
                    Cursor = Cursors.Default;
                    SetControlsEnabled(true);
                }
            }
            else
            {
                MessageBox.Show("No se seleccionó ninguna carpeta.");
            }
        }

        private void selMRC_Click(object sender, EventArgs e)
        {
            // Crear un FolderBrowserDialog para seleccionar la carpeta
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Seleccionar carpeta que contiene los archivos MRC";

                // Mostrar el cuadro de diálogo
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Obtener la carpeta seleccionada
                    string carpetaSeleccionada = folderBrowserDialog.SelectedPath;

                    // Obtener todos los archivos .mrc en la carpeta
                    List<string> archivosMRC = Directory.GetFiles(carpetaSeleccionada, "*.mrc").ToList();

                    if (archivosMRC.Count < 2)
                    {
                        MessageBox.Show("La carpeta debe contener al menos dos archivos MRC para unirlos.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    // Ruta de salida
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Archivo MRC (*.mrc)|*.mrc";
                        saveFileDialog.Title = "Guardar Archivo Unido";
                        saveFileDialog.FileName = "Archivo_Unificado.mrc";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string archivoSalida = Path.GetFullPath(saveFileDialog.FileName);

                            UnirMRKoMRC umrc = new UnirMRKoMRC();
                            umrc.UnirArchivosDesdeCarpeta(carpetaSeleccionada, archivoSalida);
                        }
                    }
                }
            }
        }

    }
}
