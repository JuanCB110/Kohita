using System;
using System.IO;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Recuerda la última carpeta usada para que las ventanas de archivos y carpetas
    /// se abran ahí y no haya que navegar cada vez. Se guarda en la configuración
    /// del usuario, así que se conserva al cerrar el programa.
    /// Solo indica dónde se abre cada ventana; no elige nada por el usuario.
    /// </summary>
    public static class CarpetasRecientes
    {
        /// <summary>Carpeta de la última base SIABUC8 (.mdb) abierta</summary>
        public static string Base
        {
            get { return SiExiste(Properties.Settings.Default.UltimaCarpetaBase); }
        }

        /// <summary>Última carpeta de trabajo: salida, .mrk, .mrc, archivos unidos o CSV de usuarios</summary>
        public static string Salida
        {
            get { return SiExiste(Properties.Settings.Default.UltimaCarpetaSalida); }
        }

        /// <summary>Guarda la carpeta del archivo .mdb elegido</summary>
        public static void RecordarBase(string archivo)
        {
            Guardar(() => Properties.Settings.Default.UltimaCarpetaBase = Path.GetDirectoryName(archivo));
        }

        /// <summary>Guarda una carpeta de trabajo elegida</summary>
        public static void RecordarSalida(string carpeta)
        {
            Guardar(() => Properties.Settings.Default.UltimaCarpetaSalida = carpeta);
        }

        /// <summary>Guarda la carpeta donde se guardó un archivo (CSV o archivo unido)</summary>
        public static void RecordarSalidaDeArchivo(string archivo)
        {
            Guardar(() => Properties.Settings.Default.UltimaCarpetaSalida = Path.GetDirectoryName(archivo));
        }

        // Una carpeta que ya no existe (USB retirada, carpeta borrada) se ignora:
        // la ventana se abre en su lugar de siempre
        private static string SiExiste(string carpeta)
        {
            return !string.IsNullOrEmpty(carpeta) && Directory.Exists(carpeta) ? carpeta : string.Empty;
        }

        // Recordar la carpeta es una comodidad: si no se puede guardar, el programa sigue igual
        private static void Guardar(Action asignar)
        {
            try
            {
                asignar();
                Properties.Settings.Default.Save();
            }
            catch (Exception)
            {
            }
        }
    }
}
