using System;
using System.Windows.Forms;

namespace ExtraccionA
{
    public partial class codigo : Form
    {
        // Propiedades públicas para devolver los valores
        public string Nom { get; private set; }
        public string Codi { get; private set; }

        public codigo()
        {
            InitializeComponent();
        }

        public void datos(string nombre, string codigo)
        {
            namebi.Text = nombre;  // namebi es el TextBox donde se ingresa el nombre
            cobi.Text = codigo;
        }

        // Este es el evento para cuando el usuario haga clic en "Aceptar"
        private void accept_Click(object sender, EventArgs e)
        {
            // Guardar el valor ingresado en las propiedades
            Nom = namebi.Text;  // namebi es el TextBox donde se ingresa el nombre
            Codi = cobi.Text;   // cobi es el TextBox donde se ingresa el código
            //Type = tipos.SelectedItem.ToString();
            this.DialogResult = DialogResult.OK;
            this.Close();       // Cerrar el formulario
        }

        // Este es el evento para cuando el usuario haga clic en "Cancelar"
        private void cancel_Click(object sender, EventArgs e)
        {
            // En caso de que el usuario cancele, no asignar valores y cerrar el formulario
            this.Close();
        }
    }
}
