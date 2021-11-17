using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Correos.SimuladorOfertas.Common
{
    public enum TipoEntidad { PERSONA_FISICA, PERSONA_JURIDICA, NINGUNO }
    public enum TipoDocumento { DNI, NIF_PERSONA, NIE, NIF_JURIDICO, CIF, NINGUNO}

    /// </summary>
    public static class AyudaNIF
    {
        public static TipoEntidad ObtenerTipoEntidad(string cadena)
        {
            try
            {
                switch (ObtenerTipoDocumento(cadena))
                {
                    case TipoDocumento.DNI: return TipoEntidad.PERSONA_FISICA;
                    case TipoDocumento.NIE: return TipoEntidad.PERSONA_FISICA;
                    case TipoDocumento.NIF_JURIDICO: return TipoEntidad.PERSONA_JURIDICA;
                    case TipoDocumento.NIF_PERSONA: return TipoEntidad.PERSONA_FISICA;
                    case TipoDocumento.CIF: return TipoEntidad.PERSONA_JURIDICA;
                    default: return TipoEntidad.NINGUNO;
                }
            }
            catch
            {
                return TipoEntidad.NINGUNO;
            }
        }

        public static TipoDocumento ObtenerTipoDocumento(string cadena)
        {
            try
            {
                cadena = cadena.Trim();

                if (cadena.Length != 9)
                {
                    return TipoDocumento.NINGUNO;
                }

                long n;
                char cIni, cFin;
                char.TryParse(cadena.ElementAt(0).ToString(), out cIni);
                char.TryParse(cadena.ElementAt(cadena.Length - 1).ToString(), out cFin);

                // DNI: El primero es número y el último es letra
                if (long.TryParse(cadena.ElementAt(0).ToString(), out n) && char.IsLetter(cadena.ElementAt(cadena.Length - 1)))
                {
                    return TipoDocumento.DNI;
                }

                // Empiezan por letra distinta de I, O, T, X, Y, Z y cumplen formato CIF
                if (char.IsLetter(cadena.ElementAt(0))
                    && cIni.Minuscula() != "i" && cIni.Minuscula() != "o" && cIni.Minuscula() != "t" && cIni.Minuscula() != "x" && cIni.Minuscula() != "y" && cIni.Minuscula() != "z"
                    && VerificarFormato(cadena, TipoDocumento.CIF))
                {
                    return TipoDocumento.CIF;
                }

                // NIF: De persona. Empiezan por K, L o M y acaban con letra
                if (char.IsLetter(cadena.ElementAt(0)) && char.IsLetter(cadena.ElementAt(cadena.Length - 1)))
                {
                    if (cIni.Minuscula() == "k" || cIni.Minuscula() == "l" || cIni.Minuscula() == "m")
                    {
                        return TipoDocumento.NIF_PERSONA;
                    }
                }

                // NIE: Empiezan por X, Y o Z y acaban con letra
                if (char.IsLetter(cadena.ElementAt(0)) && char.IsLetter(cadena.ElementAt(cadena.Length - 1)))
                {
                    if (cIni.Minuscula() == "x" || cIni.Minuscula() == "y" || cIni.Minuscula() == "z")
                    {
                        return TipoDocumento.NIE;
                    }
                }

                // NIF: Otros tipos de nif. Persona jurídica
                if (char.IsLetter(cadena.ElementAt(0)) && char.IsLetter(cadena.ElementAt(cadena.Length - 1)))
                {
                    return TipoDocumento.NIF_JURIDICO;
                }

                return TipoDocumento.NINGUNO;
            }
            catch
            {
                return TipoDocumento.NINGUNO;
            }
        }

        public static bool VerificarFormato(string cadena, TipoDocumento forzarTipoDocumento = TipoDocumento.NINGUNO)
        {
            try
            {
                cadena = cadena.Trim();

                if (cadena.Length != 9)
                {
                    return false;
                }

                char cIni, cFin;
                char.TryParse(cadena.ElementAt(0).ToString(), out cIni);
                char.TryParse(cadena.ElementAt(cadena.Length - 1).ToString(), out cFin);
                long numero = cadena.ExtraerNumero();

                TipoDocumento tipoDocumento = forzarTipoDocumento;
                if (forzarTipoDocumento == TipoDocumento.NINGUNO)
                {
                    tipoDocumento = ObtenerTipoDocumento(cadena);
                }

                switch (tipoDocumento)
                {
                    case TipoDocumento.DNI:
                        {
                            if (!cadena.NumericoDesdeHasta(0, cadena.Length - 2))
                            {
                                return false;
                            }

                            return numero.LetraDeControl(tipoDocumento) == cFin.Minuscula();
                        }
                    case TipoDocumento.NIF_PERSONA:
                    case TipoDocumento.NIF_JURIDICO:
                        {
                            if (!cadena.NumericoDesdeHasta(1, cadena.Length - 2))
                            {
                                return false;
                            }

                            return numero.LetraDeControl(tipoDocumento) == cFin.Minuscula();
                        }
                    case TipoDocumento.NIE:
                        {
                            if (!cadena.NumericoDesdeHasta(1, cadena.Length - 2))
                            {
                                return false;
                            }

                            long digitoExtra = 0;
                            if (cadena.ElementAt(0).Minuscula() == "y")
                            {
                                digitoExtra = 10000000;
                            }
                            else if (cadena.ElementAt(0).Minuscula() == "z")
                            {
                                digitoExtra = 20000000;
                            }
                            numero += digitoExtra;

                            return numero.LetraDeControl(tipoDocumento) == cFin.Minuscula();
                        }
                    case TipoDocumento.CIF:
                        {
                            if (!cadena.NumericoDesdeHasta(1, cadena.Length - 2))
                            {
                                return false;
                            }

                            return cFin.Minuscula().Equals(cadena.LetraDeControl(tipoDocumento).Item1) || cFin.Minuscula().Equals(cadena.LetraDeControl(tipoDocumento).Item2);
                        }
                    default: return false;
                }
            }
            catch
            {
                return false;
            }
        }
    }

    public static class Extensiones
    {
        public static string Minuscula(this char c)
        {
            return c.ToString().ToLower();
        }


        public static bool NumericoDesdeHasta(this string s, int inicial, int final)
        {
            char[] cadena = s.ToCharArray();
            if(s.Length <= final)
            {
                return false;
            }

            int valor = -1;
            for (int i = inicial; i < final; i++)
            {
                if(!int.TryParse(cadena[i].ToString(), out valor))
                {
                    return false;
                }
            }
            return true;
        }

        public static long ExtraerNumero(this string s)
        {
            try
            {
                string aux = new string(s.ToCharArray().Where(x => x >= '0' && x <= '9').ToArray());
                int valor = -1;
                int.TryParse(aux, out valor);
                return valor;
            }
            catch
            {
                return -1;
            }
        }

        public static int SumaDigitos(this int n)
        {
            int sum = 0;
            while (n != 0)
            {
                sum += n % 10;
                n /= 10;
            }
            return sum;
        }

        public static string LetraDeControl(this long n, TipoDocumento tipoDocumento)
        {
            if (tipoDocumento == TipoDocumento.DNI || tipoDocumento == TipoDocumento.NIE)
            {
                char[] codigosControl = new char[23]
                {'t', 'r', 'w', 'a', 'g', 'm', 'y', 'f', 'p', 'd', 'x', 'b', 'n', 'j', 'z', 's', 'q', 'v', 'h', 'l', 'c', 'k', 'e'};

                return codigosControl[n % 23].ToString();
            }
            else if (tipoDocumento == TipoDocumento.NIF_JURIDICO || tipoDocumento == TipoDocumento.NIF_PERSONA)
            {
                char[] codigosControl = new char[10]
                { 'j', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i' };

                return codigosControl[n % 10].ToString();
            }

            return null;
        }

        public static Tuple<string, string> LetraDeControl(this string n, TipoDocumento tipoDocumento)
        {
            if (tipoDocumento == TipoDocumento.CIF)
            {
                bool puedeSerLetra, puedeSerNumero;
                if (n.ElementAt(0).Minuscula() == "k" || n.ElementAt(0).Minuscula() == "p" || n.ElementAt(0).Minuscula() == "q" || n.ElementAt(0).Minuscula() == "s" || n.ElementAt(0).Minuscula() == "w")
                {
                    puedeSerLetra = true;
                    puedeSerNumero = false;
                }
                else if (n.ElementAt(0).Minuscula() == "a" || n.ElementAt(0).Minuscula() == "b" || n.ElementAt(0).Minuscula() == "e" || n.ElementAt(0).Minuscula() == "h")
                {
                    puedeSerLetra = false;
                    puedeSerNumero = true;
                }
                else
                {
                    puedeSerLetra = true;
                    puedeSerNumero = true;
                }

                string substring = n.Substring(1, n.Length - 2);

                int sumaPares = int.Parse(substring.ElementAt(1).ToString()) + int.Parse(substring.ElementAt(3).ToString()) + int.Parse(substring.ElementAt(5).ToString());
                int sumaImpares =
                    (int.Parse(substring.ElementAt(0).ToString()) * 2).SumaDigitos() +
                    (int.Parse(substring.ElementAt(2).ToString()) * 2).SumaDigitos() +
                    (int.Parse(substring.ElementAt(4).ToString()) * 2).SumaDigitos() +
                    (int.Parse(substring.ElementAt(6).ToString()) * 2).SumaDigitos();

                int control = (sumaPares + sumaImpares) % 10;

                if (control != 0)
                {
                    control = 10 - control;
                }

                char[] codigosControl = new char[10] { 'j', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i'};
                
                if (puedeSerNumero && !puedeSerLetra)
                {
                    return new Tuple<string,string>(control.ToString(), null);
                }
                else if (puedeSerLetra && !puedeSerNumero)
                {
                    return new Tuple<string, string>(codigosControl[control].ToString(), null);
                }
                else if (puedeSerLetra && puedeSerNumero)
                {
                    return new Tuple<string, string>(control.ToString(), codigosControl[control].ToString());
                }
            }

            return new Tuple<string, string>(null, null);
        }
    }
}