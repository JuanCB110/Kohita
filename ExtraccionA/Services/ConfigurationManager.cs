using System;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio para manejar la configuración de la aplicación
    /// </summary>
    public class ConfigurationManager
    {
        private static ConfigurationManager _instance;
        private static readonly object _lock = new object();

        public string ConnectionString { get; set; }
        public string OutputFolder { get; set; }
        public string MarcEditPath { get; set; }
        public string CodigoBiblioteca { get; set; }
        public string NombreBiblioteca { get; set; }

        private ConfigurationManager()
        {
            ConnectionString = string.Empty;
            OutputFolder = string.Empty;
            MarcEditPath = string.Empty;
            CodigoBiblioteca = string.Empty;
            NombreBiblioteca = string.Empty;
        }

        /// <summary>
        /// Obtiene la instancia única del ConfigurationManager (Singleton)
        /// </summary>
        public static ConfigurationManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ConfigurationManager();
                        }
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// Construye la cadena de conexión para Access
        /// </summary>
        public void SetAccessConnection(string mdbFilePath)
        {
            ConnectionString = $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={mdbFilePath};Persist Security Info=False;";
        }

        /// <summary>
        /// Valida que la configuración esté completa
        /// </summary>
        public bool IsValid()
        {
            return !string.IsNullOrEmpty(ConnectionString) &&
                   !string.IsNullOrEmpty(MarcEditPath);
        }

        /// <summary>
        /// Valida que la configuración de exportación esté completa
        /// </summary>
        public bool IsExportConfigurationValid()
        {
            return IsValid() &&
                   !string.IsNullOrEmpty(OutputFolder) &&
                   !string.IsNullOrEmpty(CodigoBiblioteca);
        }

        /// <summary>
        /// Limpia toda la configuración
        /// </summary>
        public void Reset()
        {
            ConnectionString = string.Empty;
            OutputFolder = string.Empty;
            CodigoBiblioteca = string.Empty;
            NombreBiblioteca = string.Empty;
        }
    }
}
