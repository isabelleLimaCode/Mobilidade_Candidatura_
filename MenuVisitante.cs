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
    class MenuVisitante
    {
        public static ConsoleKeyInfo menuVisitante(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            estado = "visitante";
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[31].Length) / 2) + menus[31].Length) - 2) + "}", menus[31]));
            string bemvindo = File.ReadAllText("bemvindo.txt", Encoding.Default);
            WriteLine("\n\n" + bemvindo);

            ConsoleKeyInfo decisao;
            do
            {
                decisao = ReadKey(true);
                //Console.Clear();
                switch (decisao.Key)
                {
                    case ConsoleKey.D1: // 1 - Página Inicial
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[31].Length) / 2) + menus[31].Length) - 2) + "}", menus[31]));
                            WriteLine("\n\n" + bemvindo);
                        }
                        break;
                    case ConsoleKey.D2: // 2 - Login
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
                    case ConsoleKey.D3: // 3 - Informação Geral
                        {
                            Console.Clear();
                            SubMenuInformacaoGeral.submenu_InfoGeral_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                        }
                        break;
                    case ConsoleKey.D4: // 4 - Candidatura
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[2].Length) / 2) + menus[2].Length) - 2) + "}", menus[2]));
                            WriteLine("Tem que fazer login para poder aceder à candidatura.");
                            WriteLine("Para fazer login prima a tecla '2'.");
                        }
                        break;
                    case ConsoleKey.D5: // 5 - Contactos
                        {
                            Console.Clear();
                            SubMenuContactos.submenu_Contactos(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                        }
                        break;
                    case ConsoleKey.D0: // 0 - Sair
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
                    default: // Qualquer outra tecla faz voltar à Página Inicial = 1
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[31].Length) / 2) + menus[31].Length) - 2) + "}", menus[31]));
                            WriteLine("\n\n" + bemvindo);
                        }
                        break;
                }
            }
            while (decisao.Key != ConsoleKey.D0);

            return decisao;
        }
    }
}
