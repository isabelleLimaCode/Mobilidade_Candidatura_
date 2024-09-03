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
    class SubMenu_IG_HistoricoRelatorios
    {
        public static ConsoleKeyInfo submenu_InfoGeral_1_1_2(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[14].Length) / 2) + menus[14].Length) - 2) + "}", menus[14]));
            string historicoerelatorios = File.ReadAllText("historicoerelatorios.txt", Encoding.Default);
            WriteLine("\n\n" + historicoerelatorios);
            WriteLine("\n\nLinks:");
            WriteLine("1 - www.nowportugal.pt");
            WriteLine("2 - www.ipb.pt/admissions");
            WriteLine("\n4 - Histórico da mobilidade internacional do IPB");
            WriteLine("\n6 - Relatórios de Atividades 2006/2016");
            WriteLine("\n\n9 - Voltar");
            WriteLine("0 - Sair");
            ConsoleKeyInfo decisao;

            do
            {
                decisao = ReadKey(true);
                if (decisao.Key == ConsoleKey.D1)
                {
                    Process.Start("http://www.nowportugal.pt");
                }
                if (decisao.Key == ConsoleKey.D2)
                {
                    Process.Start("http://www.ipb.pt/admissions");
                }
                if (decisao.Key == ConsoleKey.D4)
                {
                    pdfile = "histdoc1.1.pdf";
                    LeitorPDF.pdfreader(pdfile);
                }
                if (decisao.Key == ConsoleKey.D6)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[14].Length) / 2) + menus[14].Length) - 2) + "}", menus[14]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[22].Length) / 2) + menus[22].Length) - 2) + "}", menus[22]));
                    string documentospdf = File.ReadAllText("documentospdf.txt", Encoding.Default);
                    WriteLine("\n\n" + documentospdf);
                    WriteLine("\n\n1 - Relatório de atividades 2015/2016");
                    WriteLine("2 - Relatório de atividades 2014/2015");
                    WriteLine("3 - Relatório de atividades 2013/2014");
                    WriteLine("4 - Relatório de atividades 2012/2013");
                    WriteLine("5 - Relatório de atividades 2011/2012");
                    WriteLine("6 - Relatório de atividades 2010/2011");
                    WriteLine("7 - Relatório de atividades 2009/2010");
                    WriteLine("8 - Relatório de atividades 2008/2009");
                    WriteLine("9 - Relatório de atividades 2007/2008");
                    WriteLine("0 - Relatório de atividades 2006/2007");
                    WriteLine("\nV - Voltar");
                    WriteLine("S - Sair");

                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            pdfile = "histdoc2.1.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D2)
                        {
                            pdfile = "histdoc2.2.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D3)
                        {
                            pdfile = "histdoc2.3.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D4)
                        {
                            pdfile = "histdoc2.4.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D5)
                        {
                            pdfile = "histdoc2.5.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D6)
                        {
                            pdfile = "histdoc2.6.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D7)
                        {
                            pdfile = "histdoc2.7.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D8)
                        {
                            pdfile = "histdoc2.8.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D9)
                        {
                            pdfile = "histdoc2.9.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.D0)
                        {
                            pdfile = "histdoc2.10.pdf";
                            LeitorPDF.pdfreader(pdfile);
                        }
                        if (decisao.Key == ConsoleKey.V)
                        {
                            SubMenu_IG_HistoricoRelatorios.submenu_InfoGeral_1_1_2(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
                        }
                        if (decisao.Key == ConsoleKey.S)
                        {
                            Console.Clear();
                            WriteLine("Tem a certeza que quer sair? Se sim, volte a premir a mesma tecla.");
                            decisao = ReadKey(true);
                            Console.Clear();

                            if (decisao.Key == ConsoleKey.S)
                            {
                                Console.Clear();
                                WriteLine("Saiu do programa.");
                                Environment.Exit(0);
                            }
                        }
                    }
                    while (decisao.Key != ConsoleKey.S);
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
            while (decisao.Key != ConsoleKey.S);
            return decisao;
        }
    }
}
