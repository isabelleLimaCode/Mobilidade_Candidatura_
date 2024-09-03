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
    class MenuAluno
    {
        public static ConsoleKeyInfo menuAluno(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            estado = "aluno";
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[11].Length) / 2) + menus[11].Length) - 2) + "}", menus[11]));
            string menualuno = File.ReadAllText("menuAluno.txt", Encoding.Default);
            WriteLine("\n\n" + menualuno);

            ConsoleKeyInfo decisao;
            do
            {
                decisao = ReadKey(true);
                switch (decisao.Key)
                {
                    case ConsoleKey.D1:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[11].Length) / 2) + menus[11].Length) - 2) + "}", menus[11]));
                            WriteLine("\n\n" + menualuno);
                        }
                        break;
                    case ConsoleKey.D2:
                        {
                            Console.Clear();
                            SubMenuInformacaoGeral.submenu_InfoGeral_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                        }
                        break;
                    case ConsoleKey.D3:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[2].Length) / 2) + menus[2].Length) - 2) + "}", menus[2]));
                        }
                        break;
                    case ConsoleKey.D4:
                        {
                            Console.Clear();
                            SubMenuContactos.submenu_Contactos(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                        }
                        break;
                    case ConsoleKey.D5:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[9].Length) / 2) + menus[9].Length) - 2) + "}", menus[9]));
                            Login.login(usernames, passwords, ref username, ref password);
                            if (username == "gestor" && password == "gestor")
                            {
                                Console.Clear();
                                MenuGestor.menuGestor(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                            }
                            if (username == "docente" && password == "docente")
                            {
                                Console.Clear();
                                MenuDocente.menuDocente(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                            }
                            if (username == "aluno" && password == "aluno")
                            {
                                Console.Clear();
                                MenuAluno.menuAluno(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                            }
                        }
                        break;
                    case ConsoleKey.D9:
                        {
                            Console.Clear();
                            MenuAluno.menuAluno(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                        }
                        break;
                    case ConsoleKey.D0:
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
                        break;
                    default:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[11].Length) / 2) + menus[11].Length) - 2) + "}", menus[11]));
                            WriteLine("\n\n" + menualuno);
                        }
                        break;
                }
            }
            while (decisao.Key != ConsoleKey.D0);

            return decisao;
        }
    }
}
