using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExtraccionA
{
    internal class DatosFijosCorrecto
    {
        public string Codificar(string dato)
        {
            // Crear un arreglo para los datos codificados (40 posiciones exactas)
            string[] dfcodificado = new string[40];

            // Inicializar todo el arreglo con "|" (valor por defecto en Koha para campos vacíos)
            for (int i = 0; i < dfcodificado.Length; i++)
            {
                dfcodificado[i] = "|";
            }
            
            // Obtener la fecha actual y formatearla sin guiones para posiciones 0-5
            string fechaActual = DateTime.Now.ToString("yyMMdd");
            for (int i = 0; i < 6 && i < fechaActual.Length; i++)
            {
                dfcodificado[i] = fechaActual[i].ToString();
            }
            
            // Preservar el dato original para análisis de posiciones
            string datoOriginal = dato;
            
            // Normalizar el dato: eliminar espacios múltiples y dividir por espacios
            dato = dato.Trim();
            dato = Regex.Replace(dato, @"\s+", " ");
            string[] partes = dato.Split(' ');
            
            // Si no hay datos suficientes, devolver el arreglo con valores por defecto
            if (partes.Length == 0)
            {
                return string.Join("", dfcodificado);
            }
            
            // Definir variables para seguimiento
            bool primeraFechaEncontrada = false;
            bool segundaFechaEncontrada = false;
            bool paisEncontrado = false;
            bool idiomaEncontrado = false;
            
            // Procesar la primera parte (líder/datos de control)
            string lider = partes[0];
            
            // Extraer el tipo de fecha de publicación (posición 5 en SIABUC, posición 6 en Koha)
            char tipoFecha = '|';
            if (lider.Length > 5)
            {
                char tipoFechaSiabuc = lider[5];
                switch (tipoFechaSiabuc)
                {
                    case '0':
                    case '1':
                        tipoFecha = '|';
                        break;
                    case '2':
                        tipoFecha = 'r'; // Fechas de reimpresión
                        break;
                    case '3':
                        tipoFecha = 'm'; // Fechas múltiples
                        break;
                    case '4':
                        tipoFecha = 'n'; // Fecha desconocida
                        break;
                    case '5':
                        tipoFecha = 'q'; // Fecha cuestionable
                        break;
                    case '6':
                        tipoFecha = 'p'; // Fechas de distribución/lanzamiento
                        break;
                    case '7':
                        tipoFecha = 's'; // Fecha única
                        break;
                    default:
                        tipoFecha = '|';
                        break;
                }
            }
            dfcodificado[6] = tipoFecha.ToString();
            
            // Extraer solo los dígitos para la primera fecha del líder
            if (lider.Length > 6)
            {
                // Extraer todos los dígitos a partir de la posición 6
                string primeraFecha = new string(lider.Substring(6).Where(char.IsDigit).ToArray());
                
                // Copiar hasta 4 dígitos a las posiciones 7-10
                for (int i = 0; i < Math.Min(primeraFecha.Length, 4); i++)
                {
                    dfcodificado[7 + i] = primeraFecha[i].ToString();
                }
                
                // Marcar que ya encontramos la primera fecha
                if (primeraFecha.Length > 0)
                {
                    primeraFechaEncontrada = true;
                }
            }
            
            // Recorrer el resto de partes para buscar datos específicos
            for (int i = 1; i < partes.Length; i++)
            {
                string parte = partes[i].Trim();
                
                // Si está vacío, continuar
                if (string.IsNullOrEmpty(parte))
                {
                    continue;
                }
                
                // Si la parte comienza con caracteres no numéricos seguidos por números,
                // podría ser un formato como "c1982" - extraer solo los dígitos
                if (Regex.IsMatch(parte, @"^[^\d]+\d+$"))
                {
                    string soloDigitos = new string(parte.Where(char.IsDigit).ToArray());
                    
                    // Si es un año de 4 dígitos
                    if (soloDigitos.Length == 4)
                    {
                        // Determinar si es primera o segunda fecha basado en la posición en la cadena original
                        int posicion = datoOriginal.IndexOf(parte);
                        bool esSegundaFecha = posicion > 15; // Si está más adelante en la cadena, probablemente sea la segunda fecha
                        
                        if (esSegundaFecha && !segundaFechaEncontrada)
                        {
                            for (int j = 0; j < soloDigitos.Length; j++)
                            {
                                dfcodificado[11 + j] = soloDigitos[j].ToString();
                            }
                            segundaFechaEncontrada = true;
                        }
                        else if (!primeraFechaEncontrada)
                        {
                            for (int j = 0; j < soloDigitos.Length; j++)
                            {
                                dfcodificado[7 + j] = soloDigitos[j].ToString();
                            }
                            primeraFechaEncontrada = true;
                        }
                    }
                    continue;
                }
                
                // Extraer año de formatos como [YYYY] o "reimp. [YYYY]"
                if (parte.Contains("[") && parte.Contains("]"))
                {
                    string yearPattern = @"\[(\d{4})\]";
                    Match match = Regex.Match(parte, yearPattern);
                    
                    if (match.Success)
                    {
                        string year = match.Groups[1].Value;
                        
                        // Determinar si es primera o segunda fecha basado en la posición en la cadena original
                        int posicion = datoOriginal.IndexOf(parte);
                        bool esSegundaFecha = posicion > 15; // Si está más adelante en la cadena, probablemente sea la segunda fecha
                        
                        if (esSegundaFecha && !segundaFechaEncontrada)
                        {
                            for (int j = 0; j < year.Length && j < 4; j++)
                            {
                                dfcodificado[11 + j] = year[j].ToString();
                            }
                            segundaFechaEncontrada = true;
                        }
                        else if (!primeraFechaEncontrada)
                        {
                            for (int j = 0; j < year.Length && j < 4; j++)
                            {
                                dfcodificado[7 + j] = year[j].ToString();
                            }
                            primeraFechaEncontrada = true;
                        }
                    }
                }
                // Si contiene "reimp.", marcar como reimpresión
                else if (parte.Contains("reimp"))
                {
                    dfcodificado[6] = "r";
                    
                    // Extraer año si existe después de "reimp."
                    string yearPattern = @"(\d{4})";
                    Match match = Regex.Match(parte, yearPattern);
                    
                    if (match.Success)
                    {
                        string year = match.Groups[1].Value;
                        
                        // Si no se ha encontrado la segunda fecha, usarla como segunda
                        if (!segundaFechaEncontrada)
                        {
                            for (int j = 0; j < year.Length && j < 4; j++)
                            {
                                dfcodificado[11 + j] = year[j].ToString();
                            }
                            segundaFechaEncontrada = true;
                        }
                    }
                }
                // Procesar fechas numéricas (solo dígitos)
                else if (Regex.IsMatch(parte, @"^\d+$") && parte.Length <= 4)
                {
                    // Si es un año de 4 dígitos
                    if (parte.Length == 4)
                    {
                        // Determinar si es primera o segunda fecha basado en la posición en la cadena original
                        int posicion = datoOriginal.IndexOf(parte);
                        bool esSegundaFecha = posicion > 15; // Si está más adelante en la cadena, probablemente sea la segunda fecha
                        
                        if (esSegundaFecha && !segundaFechaEncontrada)
                        {
                            for (int j = 0; j < parte.Length; j++)
                            {
                                dfcodificado[11 + j] = parte[j].ToString();
                            }
                            segundaFechaEncontrada = true;
                        }
                        else if (!primeraFechaEncontrada)
                        {
                            for (int j = 0; j < parte.Length; j++)
                            {
                                dfcodificado[7 + j] = parte[j].ToString();
                            }
                            primeraFechaEncontrada = true;
                        }
                        else if (!segundaFechaEncontrada)
                        {
                            for (int j = 0; j < parte.Length; j++)
                            {
                                dfcodificado[11 + j] = parte[j].ToString();
                            }
                            segundaFechaEncontrada = true;
                        }
                    }
                }
                // Procesar códigos de país (2-3 letras)
                else if (!paisEncontrado && parte.Length <= 3 && parte.All(char.IsLetter))
                {
                    string codigoPais = parte.ToLower().PadRight(3);
                    for (int j = 0; j < 3; j++)
                    {
                        dfcodificado[15 + j] = codigoPais[j].ToString();
                    }
                    paisEncontrado = true;
                }
                // Procesar códigos de ilustración (a, b, c, d, etc.)
                else if (parte.Length == 1 && "abcdefghijklmnopqrstuvwxyz".Contains(parte.ToLower()))
                {
                    // Asignar códigos de ilustración (posiciones 18-21 en Koha)
                    for (int j = 0; j < 4; j++)
                    {
                        if (dfcodificado[18 + j] == "|")
                        {
                            dfcodificado[18 + j] = parte.ToLower();
                            break;
                        }
                    }
                }
                // Procesar códigos de idioma (tres letras, como "spa")
                else if (!idiomaEncontrado && (parte.Contains("spa") || (parte.Length == 3 && parte.All(char.IsLetter))))
                {
                    string codigoIdioma = Regex.Match(parte, @"[a-z]{3}").Value;
                    if (!string.IsNullOrEmpty(codigoIdioma))
                    {
                        for (int j = 0; j < 3; j++)
                        {
                            dfcodificado[35 + j] = codigoIdioma[j].ToString().ToLower();
                        }
                        idiomaEncontrado = true;
                    }
                }
            }
            
            // Generar la cadena final con exactamente 40 caracteres
            string datoFijo = string.Join("", dfcodificado);
            
            // Verificar que la longitud sea exactamente 40 caracteres
            if (datoFijo.Length < 40)
            {
                datoFijo = datoFijo.PadRight(40, '|');
            }
            else if (datoFijo.Length > 40)
            {
                datoFijo = datoFijo.Substring(0, 40);
            }
            
            return datoFijo;
        }
    }
}
