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
    class SubMenuContactos
    {
        public static ConsoleKeyInfo submenu_Contactos(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[3].Length) / 2) + menus[3].Length) - 2) + "}", menus[3]));
            string contactos = File.ReadAllText("contactos.txt", Encoding.Default);
            WriteLine("\n\n" + contactos);
            WriteLine("\n\n9 - Voltar");
            WriteLine("0 - Sair");
            ConsoleKeyInfo decisao;
            do
            {
                decisao = ReadKey(true);
                if (decisao.Key == ConsoleKey.D9)
                {
                    if (estado == "gestor")
                    {
                        Console.Clear();
                        MenuGestor.menuGestor(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                    }
                    if (estado == "docente")
                    {
                        Console.Clear();
                        MenuDocente.menuDocente(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                    }
                    if (estado == "aluno")
                    {
                        Console.Clear();
                        MenuAluno.menuAluno(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                    }
                    if (estado == "visitante")
                    {
                        Console.Clear();
                        MenuVisitante.menuVisitante(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                    }
                }
                if (decisao.Key == ConsoleKey.D0)
                {
                    Console.Clear();
                    WriteLine("Tem a certeza que quer sair? Se sim, volte a premir a mesma tecla.");
                    decisao = ReadKey(true);
                    Console.Clear();

                    if (decisao.Key == ConsoleKey.D0)
                    {
                        Console.Clear();
                        WriteLine("Saiu do programa.");
                        Environment.Exit(0);
                    }
                }
            }
            while (decisao.Key != ConsoleKey.D0);
            return decisao;
        }
    }
}
