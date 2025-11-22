using System;
using System.Data;
using System.Data.OleDb;
using System.Threading.Tasks;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio encargado del acceso a la base de datos Access (SIABUC8)
    /// </summary>
    public class DatabaseService
    {
        private readonly string _connectionString;

        public DatabaseService(string connectionString)
        {
            _connectionString = connectionString;
        }

        /// <summary>
        /// Ejecuta una consulta y retorna un DataTable con los resultados
        /// </summary>
        public async Task<DataTable> ExecuteQueryAsync(string query)
        {
            try
            {
                using (var conn = new OleDbConnection(_connectionString))
                {
                    await conn.OpenAsync();

                    using (var cmd = new OleDbCommand(query, conn))
                    using (var reader = await cmd.ExecuteReaderAsync())
                    {
                        if (!reader.HasRows)
                        {
                            return null;
                        }

                        DataTable dataTable = new DataTable();
                        dataTable.Load(reader);
                        return dataTable;
                    }
                }
            }
            catch (OleDbException oleDbEx)
            {
                throw new InvalidOperationException($"Error de base de datos: {oleDbEx.Message}", oleDbEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error al conectar con la base de datos: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Obtiene todas las fichas de la tabla Fichas
        /// </summary>
        public async Task<DataTable> ObtenerFichasAsync()
        {
            string query = "SELECT * FROM Fichas";
            return await ExecuteQueryAsync(query);
        }

        /// <summary>
        /// Obtiene todos los usuarios de la tabla Usuarios
        /// </summary>
        public async Task<DataTable> ObtenerUsuariosAsync()
        {
            string query = "SELECT * FROM Usuarios";
            return await ExecuteQueryAsync(query);
        }

        /// <summary>
        /// Obtiene todos los ejemplares de la tabla EJEMPLARES
        /// </summary>
        public async Task<DataTable> ObtenerEjemplaresAsync()
        {
            string query = "SELECT * FROM EJEMPLARES";
            return await ExecuteQueryAsync(query);
        }

        /// <summary>
        /// Obtiene datos combinados de Fichas y Ejemplares
        /// </summary>
        public async Task<DataTable> ObtenerFichasConEjemplaresAsync()
        {
            string query = @"SELECT 
                                f.Ficha_No AS Numero_De_Control,
                                e.NumAdqui AS Numero_de_Item,
                                f.Fecha AS Fecha_Entrada_Ficha,
                                f.Titulo AS Titulo_Auxiliar,
                                f.FechaMod AS Fecha_Modificacion_Ficha, 
                                f.DatosFijos,
                                f.ISBN AS ISBN_Auxiliar,
                                f.Clasificacion AS Clasificacion_LC, 
                                f.TipoMaterial,   
                                f.Autor AS Autor_Auxiliar,  
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

            return await ExecuteQueryAsync(query);
        }

        /// <summary>
        /// Obtiene fichas que no tienen ejemplares asociados
        /// </summary>
        public async Task<DataTable> ObtenerFichasSinEjemplaresAsync()
        {
            string query = @"SELECT 
                        f.Ficha_No AS Numero_De_Control,
                        f.Fecha AS Fecha_Entrada_Ficha,
                        f.Titulo AS Titulo_Auxiliar,
                        f.FechaMod AS Fecha_Modificacion_Ficha, 
                        f.DatosFijos,
                        f.ISBN AS ISBN_Auxiliar,
                        f.Clasificacion AS Clasificacion_LC, 
                        f.TipoMaterial,   
                        f.Autor AS Autor_Auxiliar,  
                        f.Estatus,  
                        f.EtiquetasMARC AS MARC_Completo
                    FROM Fichas f
                    LEFT JOIN EJEMPLARES e ON f.Ficha_No = e.Ficha_No
                    WHERE e.Ficha_No IS NULL
                    ORDER BY f.Ficha_No;";

            return await ExecuteQueryAsync(query);
        }

        /// <summary>
        /// Obtiene información de bibliotecas
        /// </summary>
        public async Task<DataTable> ObtenerBibliotecasAsync()
        {
            string query = "SELECT * FROM Bibliotecas";
            return await ExecuteQueryAsync(query);
        }
    }
}
