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
    class SubMenu_IG_RegulamentosDocumentacao
    {
        public static ConsoleKeyInfo submenu_InfoGeral_1_1_1(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
            WriteLine("\n\n1 - Documentos institucionais e distinções");
            WriteLine("2 - Acordos bilaterais");
            WriteLine("3 - Regulamentos");
            WriteLine("4 - Modelos de Curriculum Vitae");
            WriteLine("5 - Formulários de Mobilidade");
            WriteLine("\n9 - Voltar");
            WriteLine("0 - Sair");
            ConsoleKeyInfo decisao;
            do
            {
                decisao = ReadKey(true);
                if (decisao.Key == ConsoleKey.D1)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[17].Length) / 2) + menus[17].Length) - 2) + "}", menus[17]));
                    string documentospdf = File.ReadAllText("documentospdf.txt", Encoding.Default);
                    WriteLine("\n\n" + documentospdf);
                    WriteLine("\n\n1 - Carta Erasmus 2014-2020 do Instituto Politécnico de Bragança");
                    WriteLine("2 - Selo ECTS 2011-2014");
                    WriteLine("3 - Avaliação dos peritos (referente ao documento 2)");
                    WriteLine("4 - Selo Suplemento ao Diploma 2013-2016");
                    WriteLine("5 - Avaliação dos peritos (referente ao documento 4)");
                    WriteLine("\n9 - Voltar");
                    WriteLine("0 - Sair");
                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            pdfile = "doc1.1.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D2)
                        {
                            pdfile = "doc1.2.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D3)
                        {
                            pdfile = "doc1.3.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D4)
                        {
                            pdfile = "doc1.4.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D5)
                        {
                            pdfile = "doc1.5.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenu_IG_RegulamentosDocumentacao.submenu_InfoGeral_1_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
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
                }
                if (decisao.Key == ConsoleKey.D2)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[18].Length) / 2) + menus[18].Length) - 2) + "}", menus[18]));
                    string documentospdf = File.ReadAllText("documentospdf.txt", Encoding.Default);
                    WriteLine("\n\n" + documentospdf);
                    WriteLine("\n\n1 - Lista de IES parceiras para mobilidade Erasmus e áreas de estudo");
                    WriteLine("2 - Lista de IES parceiras para mobilidade extracomunitária (internacional)");
                    WriteLine("\n9 - Voltar");
                    WriteLine("0 - Sair");

                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            pdfile = "doc2.1.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D2)
                        {
                            pdfile = "doc2.2.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenu_IG_RegulamentosDocumentacao.submenu_InfoGeral_1_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
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
                    while (decisao.Key != ConsoleKey.D0 || decisao.Key != ConsoleKey.D9);
                }
                if (decisao.Key == ConsoleKey.D3)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[19].Length) / 2) + menus[19].Length) - 2) + "}", menus[19]));
                    string documentospdf = File.ReadAllText("documentospdf.txt", Encoding.Default);
                    WriteLine("\n\n" + documentospdf);
                    WriteLine("\n\n1 - Regulamento do programa de mobilidade de estudantes Erasmus+");
                    WriteLine("\n9 - Voltar");
                    WriteLine("0 - Sair");

                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            pdfile = "doc3.1.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenu_IG_RegulamentosDocumentacao.submenu_InfoGeral_1_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
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
                    while (decisao.Key != ConsoleKey.D0 || decisao.Key != ConsoleKey.D9);
                }
                if (decisao.Key == ConsoleKey.D4)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[20].Length) / 2) + menus[20].Length) - 2) + "}", menus[20]));
                    string documentospdf = File.ReadAllText("documentospdf.txt", Encoding.Default);
                    WriteLine("\n\n" + documentospdf);
                    WriteLine("\n\n1 - Modelo de CV (língua portuguesa)");
                    WriteLine("2 - instruções para preenchimento (Documento 1)");
                    WriteLine("3 - Modelo de CV (língua inglesa)");
                    WriteLine("4 - instruções para preenchimento (Documento 3)");
                    WriteLine("\n9 - Voltar");
                    WriteLine("0 - Sair");

                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            pdfile = "doc4.1.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D2)
                        {
                            pdfile = "doc4.2.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D3)
                        {
                            pdfile = "doc4.3.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D4)
                        {
                            pdfile = "doc4.4.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenu_IG_RegulamentosDocumentacao.submenu_InfoGeral_1_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
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
                    while (decisao.Key != ConsoleKey.D0 || decisao.Key != ConsoleKey.D9);
                }
                if (decisao.Key == ConsoleKey.D5)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[21].Length) / 2) + menus[21].Length) - 2) + "}", menus[21]));
                    string documentospdf = File.ReadAllText("documentospdf.txt", Encoding.Default);
                    WriteLine("\n\n" + documentospdf);
                    WriteLine("\n\n1 - Formulário para mobilidade Erasmus - Estudos");
                    WriteLine("2 - Formulário para mobilidade Erasmus - Estágios");
                    WriteLine("\n9 - Voltar");
                    WriteLine("0 - Sair");

                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            pdfile = "doc5.1.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D2)
                        {
                            pdfile = "doc5.2.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenu_IG_RegulamentosDocumentacao.submenu_InfoGeral_1_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
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
                    while (decisao.Key != ConsoleKey.D0 || decisao.Key != ConsoleKey.D9);
                }
                if (decisao.Key == ConsoleKey.D9)
                {
                    SubMenuInformacaoGeral.submenu_InfoGeral_1_1(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
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
            while (decisao.Key != ConsoleKey.D9);
            return decisao;
        }
    }
}
