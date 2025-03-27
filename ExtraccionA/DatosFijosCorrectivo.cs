using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExtraccionA
{
    internal class DatosFijosCorrectivo
    {
        public string Codificar(string dato)
        {
            //Dato prueba
            string prueba = string.Empty;

            // Limpiar el dato, reemplazar los espacios por '|' y eliminar múltiples '|'
            dato = dato.Trim().Replace(" ", "|");
            dato = Regex.Replace(dato, @"\|+", "|");

            // Dividir el dato en partes usando el separador '|'
            string[] datos = dato.Split('|');

            // Crear un arreglo para los datos codificados (esto puedes ajustar según sea necesario)
            string[] dfcodificado = new string[40];

            // Obtener la fecha actual y formatearla sin guiones
            string fecha = DateTime.Now.ToString("yyMMdd");
            //dfcodificado[0] = fecha;

            // Codificar entre los primeros 6 y 10 dígitos
            string dato1 = datos[0];  // Se asume que los primeros dígitos están en datos[0]
            //Verificar tipo de fecha
            bool fechaaprox = false;
            if (dato1.Contains("?") || dato1.Contains("[") || dato1.Contains("]") || dato1.Contains("-") || dato1.Contains("--"))
            {
                fechaaprox = true;
            }
            if (dato1.Contains("s"))
            {
                fechaaprox = false;
            }
            //Primer dato que abarca de los primeros 6 o 10 digitos, excluyendo cualquier caracter menos los digitos
            char[] digitos = new string(dato1.Where(char.IsDigit).ToArray()).ToCharArray();
            for (int i = 0; i < digitos.Length; i++)
            {
                char d = digitos[i]; // Obtener el carácter en la posición i

                // Usamos un switch para manejar casos específicos según la posición
                switch (i)
                {
                    //Estado del registro
                    case 0:
                        // Caso específico para la primera posición (i == 0)
                        switch (d)
                        {
                            case '0':
                            case '1':
                            //SIABUC8 " ": No se utiliza - Koha: "|"
                            case '2':
                            //SIABUC8 "a":  - Koha: ""
                            case '3':
                            //SIABUC8 "c":  - Koha: ""
                            case '4':
                            //SIABUC8 "d":  - Koha: ""
                            case '5':
                                //SIABUC8 "n":  - Koha: ""
                                //dfcodificado[0] = fecha;
                                char[] fechadiv = new string(fecha.Where(char.IsDigit).ToArray()).ToCharArray();
                                for (int j = 0; j < fechadiv.Length; j++)
                                {
                                    dfcodificado[j] = fechadiv[j].ToString(); // Obtener el carácter en la posición i
                                }
                                break;
                            default:
                                Console.WriteLine($"En la posición 1 hay un carácter distinto: {d}");
                                break;
                        }
                        break;
                    //Tipo de registro
                    case 1:
                        break;
                    //Nivel Bibliografico
                    case 2:
                        break;
                    //Nivel de codificacion
                    case 3:
                        break;
                    //Catalogacion descriptiva
                    case 4:
                        break;
                    //Tipo de fecha de publicacion
                    case 5:
                        switch (d)
                        {
                            case '0':
                            case '1':
                                //SIABUC8 " ":  - Koha: "|"
                                if (fechaaprox == true)
                                {
                                    dfcodificado[6] = "n";
                                }
                                else
                                {
                                    dfcodificado[6] = "|";
                                }
                                break;
                            case '2':
                                //SIABUC8 "c":  - Koha: "r"
                                dfcodificado[6] = "r";
                                break;
                            case '3':
                                //SIABUC8 "m":  - Koha: "m"
                                dfcodificado[6] = "m";
                                break;
                            case '4':
                                //SIABUC8 "n":  - Koha: "n"
                                dfcodificado[6] = "n";
                                break;
                            case '5':
                                //SIABUC8 "q":  - Koha: "q"
                                dfcodificado[6] = "q";
                                break;
                            case '6':
                                //SIABUC8 "r":  - Koha: "p"
                                dfcodificado[6] = "p";
                                break;
                            case '7':
                                //SIABUC8 "s":  - Koha: "s"
                                dfcodificado[6] = "s";
                                break;
                            default:
                                break;
                        }
                        break;
                    //Primera fecha
                    case 6:
                        dfcodificado[7] = d.ToString();
                        break;
                    case 7:
                        dfcodificado[8] = d.ToString();
                        break;
                    case 8:
                        if (fechaaprox)
                        {
                            if (!d.ToString().Any(c => char.IsDigit(c)))
                            {
                                dfcodificado[9] = "?";
                            }
                            else
                            {
                                dfcodificado[9] = d.ToString();
                            }
                        }
                        else
                        {
                            dfcodificado[9] = d.ToString();
                        }
                        break;
                    case 9:
                        if (fechaaprox)
                        {
                            if (!d.ToString().Any(c => char.IsDigit(c)))
                            {
                                dfcodificado[10] = "?";
                            }
                            else
                            {
                                dfcodificado[10] = d.ToString();
                            }
                        }
                        else
                        {
                            dfcodificado[10] = d.ToString();
                        }
                        break;
                    default:
                        // Para las posiciones restantes, manejamos casos generales
                        break;
                }
            }

            //Verificar si la primera fecha existe
            for (int i = 7; i <= 10; i++)
            {
                if (string.IsNullOrWhiteSpace(dfcodificado[i]) || string.IsNullOrEmpty(dfcodificado[i]))
                {
                    dfcodificado[i] = " ";
                }
            }

            //Segundo dato
            string dato2 = datos[1];

            //Identificar si se trata de una posible segunda fecha o del pais de publicacion, pero tambien si se trata de una reimpresion
            if (dato2.Contains("re") || dato2.Contains("Re"))
            {
                string dato3 = datos[2];
                char[] fechados = new string(dato3.Where(char.IsDigit).ToArray()).ToCharArray();
                for (int i = 0; i < fechados.Length; i++)
                {
                    dfcodificado[11 + i] = fechados[i].ToString();
                }
                //Verificar si la segunda fecha existe
                for (int i = 11; i <= 14; i++)
                {
                    if (string.IsNullOrWhiteSpace(dfcodificado[i]) || string.IsNullOrEmpty(dfcodificado[i]))
                    {
                        dfcodificado[i] = " ";
                    }
                }
                //Caso del pais
                string dato4 = datos[3];
                if (dato4.Length < 4)
                {
                    dato4 = dato4.PadRight(4);
                    char[] pais = dato4.ToCharArray();
                    for (int i = 0; i < 2; i++)
                    {
                        dfcodificado[15 + i] = pais[i].ToString();
                    }
                }
            }

            //Manejar en caso de que solamente tenga la fecha y/o lugar de publicacion
            else if (!dato2.Contains("re") || !dato2.Contains("Re"))
            {
                if (dato2.Contains("["))
                {
                    char[] fechados = new string(dato2.Where(char.IsDigit).ToArray()).ToCharArray();
                    for (int i = 0; i < fechados.Length; i++)
                    {
                        dfcodificado[11 + i] = fechados[i].ToString();
                    }
                    //Caso del pais
                    string dato3 = datos[2];
                    if (dato3.Length < 4)
                    {
                        dato3 = dato3.PadRight(4);
                        char[] pais = dato3.ToCharArray();
                        for (int i = 0; i < 2; i++)
                        {
                            dfcodificado[15 + i] = pais[i].ToString();
                        }
                    }
                }
                else
                {
                    char[] fechados = new string(dato2.Where(char.IsDigit).ToArray()).ToCharArray();
                    for (int i = 0; i < fechados.Length; i++)
                    {
                        dfcodificado[11 + i] = fechados[i].ToString();
                    }
                    //Caso del pais
                    string dato3 = datos[2];
                    if (dato3.Length < 4)
                    {
                        dato3 = dato3.PadRight(4);
                        char[] pais = dato3.ToCharArray();
                        for (int i = 0; i < 2; i++)
                        {
                            dfcodificado[15 + i] = pais[i].ToString();
                        }
                    }
                }
                if (dato2.All(Char.IsLetter))
                {
                    //Caso del pais
                    if (dato2.Length < 4)
                    {
                        dato2 = dato2.PadRight(4);
                        char[] pais = dato2.ToCharArray();
                        for (int i = 0; i < 2; i++)
                        {
                            dfcodificado[15 + i] = pais[i].ToString();
                        }
                    }
                }
            }

            // Finalmente, puedes devolver el dato original o hacer algo con los datos procesados.
            string datofijo = string.Join("", dfcodificado); //string.Join("|", datos); // Ejemplo: reconstruir la cadena original
            return datofijo;
        }
    }
}
