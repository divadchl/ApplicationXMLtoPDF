using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ApplicationXMLtoPDF.Services
{
    public class Moneda
    {
        private string strPesosMN;
        public string PesosMN(double tyCantidad)
        {
            tyCantidad = Math.Round(tyCantidad, 2);
            double lyCantidad = (int)tyCantidad;
            double lyCentavos = (tyCantidad - lyCantidad) * 100;
            string[] laUnidades = new string[] { "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiún", "veintidós", "veintitrés", "veinticuatro", "veinticinco", "veintiséis", "veintisiete", "veintiocho", "veintinueve" };
            string[] laDecenas = new string[] { "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa" };
            string[] laCentenas = new string[] { "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos" };
            int lnNumeroBloques = 1;

            byte lnDigito = 0;
            do
            {
                byte lnPrimerDigito = 0;
                byte lnSegundoDigito = 0;
                byte lnTercerDigito = 0;
                string lcBloque = "";
                int lnBloqueCero = 0;
                for (int i = 1; i <= 3; i++)
                {
                    lnDigito = (byte)(lyCantidad % 10);
                    if (lnDigito != 0)
                    {
                        switch (i)
                        {
                            case 1:
                                lcBloque = " " + laUnidades[lnDigito - 1];
                                lnPrimerDigito = lnDigito;
                                break;
                            case 2:
                                if (lnDigito <= 2)
                                {
                                    lcBloque = " " + laUnidades[(lnDigito * 10) + lnPrimerDigito - 1];
                                }
                                else
                                {
                                    if (lnPrimerDigito != 0)
                                    {
                                        lcBloque = " " + laDecenas[lnDigito - 1] + " y" + lcBloque;
                                    }
                                    else
                                    {
                                        lcBloque = " " + laDecenas[lnDigito - 1] + null + lcBloque;
                                    }
                                }
                                lnSegundoDigito = lnDigito;
                                break;
                            case 3:
                                if (lnDigito == 1 && lnPrimerDigito == 0 && lnSegundoDigito == 0)
                                {
                                    lcBloque = " " + "cien" + lcBloque;
                                }
                                else
                                {
                                    lcBloque = " " + laCentenas[lnDigito - 1] + lcBloque;
                                }
                                lnTercerDigito = lnDigito;
                                break;
                        }
                    }
                    else
                    {
                        lnBloqueCero = lnBloqueCero + 1;
                    }
                    lyCantidad = (int)(lyCantidad / 10);
                    if (lyCantidad == 0)
                    {
                        break;
                    }
                }
                switch (lnNumeroBloques)
                {

                    case 1:
                        strPesosMN = lcBloque;
                        break;
                    case 2:
                        if (lnBloqueCero == 3)
                        {
                            strPesosMN = lcBloque + null + strPesosMN;
                        }
                        else
                        {
                            strPesosMN = lcBloque + " mil" + strPesosMN;
                        }
                        break;
                    case 3:
                        //if (lnDigito == 1 && lnPrimerDigito == 0 && lnSegundoDigito == 0)
                        if (lnPrimerDigito == 1 && lnSegundoDigito == 0 && lnTercerDigito == 0)
                        {
                            strPesosMN = lcBloque + " millón" + strPesosMN;
                            //lcBloque = " " + "cien" + lcBloque;
                        }
                        else
                        {
                            strPesosMN = lcBloque + " millones" + strPesosMN;
                            //lcBloque = " " + laCentenas[lnDigito - 1] + lcBloque;
                        }

                        lnTercerDigito = lnDigito;
                        break;
                    case 4:
                        if (lnBloqueCero == 0)
                        {
                            strPesosMN = lcBloque + null + strPesosMN;
                        }
                        else
                        {
                            strPesosMN = lcBloque + " cero" + strPesosMN;
                        }
                        break;
                }
                lnNumeroBloques = lnNumeroBloques + 1;
            }
            while (lyCantidad != 0);
            if ((int)tyCantidad == 1)
            {
                strPesosMN = "(" + strPesosMN.Trim().Substring(0, 1).ToUpper() + strPesosMN.TrimStart().Substring(1, strPesosMN.TrimStart().Length - 1) + " peso " + String.Format("{0:0}", lyCentavos) + "/100 M.N.)";
            }
            else
            {
                strPesosMN = "(" + strPesosMN.Trim().Substring(0, 1).ToUpper() + strPesosMN.TrimStart().Substring(1, strPesosMN.TrimStart().Length - 1) + " pesos " + String.Format("{0:0}", lyCentavos) + "/100 M.N.)";
            }

            return strPesosMN;

        }
    }
}