using System;
using System.IO;
using System.Text;
using static System.Console;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace Prot_2___Menus
{
    class SubMenuListaUtilizadores
    {
        public static void submenu_ListaUtilizadores(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[7].Length) / 2) + menus[7].Length) - 2) + "}", menus[7]));
            ConsultaHistograma.consultaHistograma(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
        }
    }
}
