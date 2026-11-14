using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace ExtraccionA
{
    /// <summary>
    /// Ventana pequeña con barra de avance y botón Cancelar para procesos largos
    /// </summary>
    internal class VentanaProgreso : Form
    {
        private readonly ProgressBar _barra;
        private readonly Label _texto;
        private readonly Button _cancelar;
        private readonly CancellationTokenSource _cancelacion;
        private bool _terminado;

        public VentanaProgreso(string titulo, int total, CancellationTokenSource cancelacion)
        {
            _cancelacion = cancelacion;

            Text = titulo;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(460, 120);

            _texto = new Label
            {
                Location = new Point(12, 12),
                Size = new Size(436, 34),
                Text = "Preparando…"
            };

            _barra = new ProgressBar
            {
                Location = new Point(12, 50),
                Size = new Size(436, 22),
                Minimum = 0,
                Maximum = Math.Max(total, 1)
            };

            _cancelar = new Button
            {
                Text = "Cancelar",
                Location = new Point(373, 82),
                Size = new Size(75, 27)
            };
            _cancelar.Click += (s, e) => Cancelar();

            Controls.Add(_texto);
            Controls.Add(_barra);
            Controls.Add(_cancelar);

            // Cerrar con la X equivale a cancelar
            FormClosing += (s, e) =>
            {
                if (!_terminado)
                {
                    e.Cancel = true;
                    Cancelar();
                }
            };
        }

        public void Actualizar(ProgresoConversion progreso)
        {
            _barra.Value = Math.Min(Math.Max(progreso.Actual, 0), _barra.Maximum);
            if (!_cancelacion.IsCancellationRequested)
                _texto.Text = $"Convirtiendo {progreso.Actual:N0} de {progreso.Total:N0}: {progreso.Archivo}";
        }

        public void Terminar()
        {
            _terminado = true;
            Close();
        }

        private void Cancelar()
        {
            if (_cancelacion.IsCancellationRequested)
                return;

            _cancelacion.Cancel();
            _cancelar.Enabled = false;
            _texto.Text = "Cancelando… se detendrá al terminar el archivo actual.";
        }
    }
}
