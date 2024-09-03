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
    class MenuDocente
    {
        public static ConsoleKeyInfo menuDocente(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            estado = "docente";
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[10].Length) / 2) + menus[10].Length) - 2) + "}", menus[10]));
            string menudocente = File.ReadAllText("menuDocente.txt", Encoding.Default);
            WriteLine("\n\n" + menudocente);

            ConsoleKeyInfo decisao;
            do
            {
                decisao = ReadKey(true);
                switch (decisao.Key)
                {
                    case ConsoleKey.D1:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[10].Length) / 2) + menus[10].Length) - 2) + "}", menus[10]));
                            WriteLine("\n\n" + menudocente);
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
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[5].Length) / 2) + menus[5].Length) - 2) + "}", menus[5]));
                        }
                        break;
                    case ConsoleKey.D6:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[6].Length) / 2) + menus[6].Length) - 2) + "}", menus[6]));
                            WriteLine("Novo utilizador registado com sucesso!");
                        }
                        break;
                    case ConsoleKey.D7:
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[7].Length) / 2) + menus[7].Length) - 2) + "}", menus[7]));
                            WriteLine();
                        }
                        break;
                    case ConsoleKey.D8:
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
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[10].Length) / 2) + menus[10].Length) - 2) + "}", menus[10]));
                            WriteLine("\n\n" + menudocente);
                        }
                        break;
                }
            }
            while (decisao.Key != ConsoleKey.D0);

            return decisao;
        }
    }
}
