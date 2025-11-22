using System;
using System.Text;
using System.Text.RegularExpressions;

namespace ExtraccionA.Services
{
    /// <summary>
    /// Servicio encargado del formateo de etiquetas MARC21
    /// </summary>
    public class MarcFormatterService
    {
        /// <summary>
        /// Normaliza texto eliminando saltos de línea y espacios extra
        /// </summary>
        public string NormalizarTexto(string texto)
        {
            if (string.IsNullOrEmpty(texto))
                return string.Empty;

            // Normalizar texto: eliminar saltos de línea y espacios extra
            string resultado = texto.Trim()
                .Replace("\r\n", " ")
                .Replace("\n", " ");

            // Eliminar espacios múltiples
            while (resultado.Contains("  "))
                resultado = resultado.Replace("  ", " ");

            return resultado;
        }

        /// <summary>
        /// Formatea la etiqueta 020 (ISBN)
        /// </summary>
        public string FormatearEtiqueta020(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=020  $a{normalizado}";
        }

        /// <summary>
        /// Formatea la etiqueta 040 (Fuente de catalogación)
        /// </summary>
        public string FormatearEtiqueta040(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=040  $c{normalizado}";
        }

        /// <summary>
        /// Formatea la etiqueta 050 (Clasificación LC)
        /// </summary>
        public string FormatearEtiqueta050(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=050  $a{normalizado}";
        }

        /// <summary>
        /// Formatea las etiquetas 100, 110, 130 (Autor principal)
        /// </summary>
        public string FormatearEtiquetaAutor(string etiqueta, string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"={etiqueta}  $a{normalizado}";
        }

        /// <summary>
        /// Formatea la etiqueta 240 (Título uniforme)
        /// </summary>
        public string FormatearEtiqueta240(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=240  $a{normalizado}";
        }

        /// <summary>
        /// Formatea la etiqueta 245 (Título)
        /// </summary>
        public string FormatearEtiqueta245(string valor)
        {
            string normalizado = NormalizarTexto(valor);

            if (normalizado.Contains(":"))
            {
                string[] partes = normalizado.Split(':');
                string a = partes[0].Trim();
                string b = partes[1].Trim();

                if (b.Contains("/"))
                {
                    string[] partes2 = b.Split('/');
                    b = partes2[0].Trim();
                    string c = " / $c" + partes2[1].Trim();
                    return $"=245  $a{a} : $b{b}{c}";
                }
                else
                {
                    return $"=245  $a{a} : $b{b}";
                }
            }
            else if (normalizado.Contains("/"))
            {
                string[] partes = normalizado.Split('/');
                string a = partes[0].Trim();
                string c = partes[1].Trim();
                return $"=245  $a{a} / $c{c}";
            }
            else
            {
                return $"=245  $a{normalizado}";
            }
        }

        /// <summary>
        /// Formatea la etiqueta 250 (Edición)
        /// </summary>
        public string FormatearEtiqueta250(string valor)
        {
            string normalizado = NormalizarTexto(valor);

            if (normalizado.Contains("/"))
            {
                string[] partes = normalizado.Split('/');
                string a = partes[0].Trim();
                string b = partes[1].Trim();
                return $"=250  $a{a} $b / {b}";
            }
            else
            {
                return $"=250  $a{normalizado}";
            }
        }

        /// <summary>
        /// Formatea la etiqueta 260 (Publicación)
        /// </summary>
        public string FormatearEtiqueta260(string valor)
        {
            string normalizado = NormalizarTexto(valor);

            if (normalizado.Contains(":"))
            {
                string[] partes = normalizado.Split(':');
                string a = partes[0].Trim();
                string b = partes[1].Trim();
                return $"=260  $a{a} : $b{b}";
            }
            else
            {
                return $"=260  $a{normalizado}";
            }
        }

        /// <summary>
        /// Formatea la etiqueta 300 (Descripción física)
        /// </summary>
        public string FormatearEtiqueta300(string valor)
        {
            string normalizado = NormalizarTexto(valor);

            if (normalizado.Contains("/"))
            {
                string[] partes = normalizado.Split('/');
                string a = partes[0].Trim();
                string c = partes[1].Trim();
                return $"=300  $a{a} / $c{c}";
            }
            else
            {
                return $"=300  $a{normalizado}";
            }
        }

        /// <summary>
        /// Formatea las etiquetas 440/490 (Serie)
        /// </summary>
        public string FormatearEtiqueta490(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=490  $a{normalizado}";
        }

        /// <summary>
        /// Formatea la etiqueta 500 (Nota general)
        /// </summary>
        public string FormatearEtiqueta500(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=500  $a{normalizado}";
        }

        /// <summary>
        /// Formatea las etiquetas 502, 503, 504 (Notas de bibliografía)
        /// </summary>
        public string FormatearEtiqueta504(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=504  $a{normalizado}";
        }

        /// <summary>
        /// Formatea la etiqueta 505 (Contenido)
        /// </summary>
        public string FormatearEtiqueta505(string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"=505  $a{normalizado}";
        }

        /// <summary>
        /// Formatea etiquetas de materia (600, 610, 650, 651) con manejo de múltiples temas
        /// </summary>
        public string FormatearEtiquetaMateria(string etiqueta, string valor)
        {
            StringBuilder resultado = new StringBuilder();
            string normalizado = NormalizarTexto(valor);

            // Eliminar numeraciones iniciales como "1. " o "2. "
            if (Regex.IsMatch(normalizado, @"^\d+\.\s"))
            {
                normalizado = Regex.Replace(normalizado, @"^\d+\.\s", "");
            }

            // Dividir la cadena principal en partes (por si hay múltiples temas numerados)
            string[] partes = Regex.Split(normalizado, @"(?<=\.)\s*(?=\d+\.)");

            if (partes.Length > 1)
            {
                // Iterar sobre cada parte
                foreach (var parte in partes)
                {
                    string parteLimpia = NormalizarTexto(parte);

                    // Eliminar numeraciones iniciales de cada parte
                    if (Regex.IsMatch(parteLimpia, @"^\d+\.\s"))
                    {
                        parteLimpia = Regex.Replace(parteLimpia, @"^\d+\.\s", "");
                    }

                    if (!string.IsNullOrWhiteSpace(parteLimpia))
                    {
                        resultado.AppendLine($"={etiqueta}  $a{parteLimpia}");
                    }
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(normalizado))
                {
                    resultado.AppendLine($"={etiqueta}  $a{normalizado}");
                }
            }

            return resultado.ToString().TrimEnd('\r', '\n');
        }

        /// <summary>
        /// Formatea la etiqueta 700 (Autor adicional)
        /// </summary>
        public string FormatearEtiqueta700(string valor)
        {
            StringBuilder resultado = new StringBuilder();
            string normalizado = NormalizarTexto(valor);

            // Dividir la cadena principal en partes
            string[] partes = Regex.Split(normalizado, @"(?<=\\)\s*");

            if (partes.Length > 1)
            {
                // Iterar sobre cada parte
                foreach (var parte in partes)
                {
                    string parteLimpia = NormalizarTexto(parte).Replace("\\", "").Trim();

                    if (!string.IsNullOrWhiteSpace(parteLimpia))
                    {
                        resultado.AppendLine($"=700  $a{parteLimpia}");
                    }
                }
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(normalizado))
                {
                    resultado.AppendLine($"=700  $a{normalizado}");
                }
            }

            return resultado.ToString().TrimEnd('\r', '\n');
        }

        /// <summary>
        /// Formatea cualquier etiqueta genérica con subcampo $a
        /// </summary>
        public string FormatearEtiquetaGenerica(string etiqueta, string valor)
        {
            string normalizado = NormalizarTexto(valor);
            return $"={etiqueta}  $a{normalizado}";
        }
    }
}
