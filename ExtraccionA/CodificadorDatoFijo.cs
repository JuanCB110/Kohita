using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExtraccionA
{
    internal class CodificadorDatoFijo
    {
        public string Datofijo { private get; set; }
        public string Codificado { get; private set; }
        private string[] DatoCompleto = new string[40]; // Arreglo para almacenar los datos procesados

        public void ProcesarDatosFijos()
        {
            string datosFijos = Datofijo;

            //if (string.IsNullOrWhiteSpace(datosFijos) || datosFijos.Length != 40)
            //{
            //    Console.WriteLine("El dato no es válido. Debe tener exactamente 40 caracteres.");
            //    return;
            //}

            // Lista para almacenar los datos procesados con su significado
            List<string> posiciones = new List<string>();

            // Primera parte: Procesar 0-9
            string primeraParte = datosFijos.Substring(0, 10).Trim();
            if (!string.IsNullOrEmpty(primeraParte))
            {
                if (primeraParte.Contains("["))
                {
                    if (primeraParte.Contains("199-")) // Caso especial: [199-?]
                    {
                        posiciones.Add("Año aproximado: 199?");
                        DatoCompleto[0] = "199?";
                    }
                    else if (primeraParte.Contains("s.n.")) // Caso especial: [s.n.]
                    {
                        posiciones.Add("Sin número: " + primeraParte.Substring(0, 6));
                        DatoCompleto[0] = primeraParte.Substring(0, 6);
                    }
                    else // Caso general de corchetes
                    {
                        string limpio = primeraParte.Replace("[", "").Replace("]", "");
                        posiciones.Add("Primera fecha: " + limpio);
                        DatoCompleto[0] = limpio;
                    }
                }
                else
                {
                    for (int i = 0; i < Math.Min(6, primeraParte.Length); i++)
                    {
                        posiciones.Add($"Posición {i + 1}: {primeraParte[i]}");
                        DatoCompleto[i] = primeraParte[i].ToString();
                    }

                    if (primeraParte.Length > 6) // Procesar fecha primaria si existe
                    {
                        string fechaPrimaria = primeraParte.Substring(6, 4);
                        posiciones.Add("Fecha primaria: " + fechaPrimaria);
                        DatoCompleto[6] = fechaPrimaria;
                    }
                }
            }

            // Segunda parte: Procesar 10-19 (Fecha secundaria)
            string segundaParte = datosFijos.Substring(10, 10).Trim();
            if (!string.IsNullOrWhiteSpace(segundaParte))
            {
                posiciones.Add("Fecha secundaria: " + segundaParte);
                DatoCompleto[10] = segundaParte;
            }

            // Tercera parte: Procesar 20-39
            string terceraParte = datosFijos.Substring(20, 20);

            string pais = terceraParte.Substring(0, 3).Trim();
            posiciones.Add("País: " + pais);
            DatoCompleto[20] = pais;

            string codigosIlustracion = terceraParte.Substring(3, 4).Trim();
            posiciones.Add("Códigos de ilustración: " + codigosIlustracion);
            DatoCompleto[23] = codigosIlustracion;

            string codigosInternos = terceraParte.Substring(7, 10).Trim();
            for (int i = 0; i < codigosInternos.Length; i++)
            {
                posiciones.Add($"Código interno posición {i + 1}: {codigosInternos[i]}");
                DatoCompleto[27 + i] = codigosInternos[i].ToString();
            }

            string idioma = terceraParte.Substring(17, 3).Trim();
            posiciones.Add("Idioma: " + idioma);
            DatoCompleto[37] = idioma;

            string ultimasPosiciones = terceraParte.Substring(20, 2).Trim();
            posiciones.Add("Últimas posiciones: " + ultimasPosiciones);
            DatoCompleto[39] = ultimasPosiciones;

            // Construir el dato codificado
            Codificado = string.Join("", DatoCompleto);

            // Mostrar resultados
            Console.WriteLine("Procesamiento completo:");
            foreach (var posicion in posiciones)
            {
                Console.WriteLine(posicion);
            }

            Console.WriteLine("Dato codificado: " + Codificado);
        }


    }
}
