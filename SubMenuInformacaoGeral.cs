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
    class SubMenuInformacaoGeral
    {
        public static ConsoleKeyInfo submenu_InfoGeral_1_1(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            ConsoleKeyInfo decisao;
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
            WriteLine("\n1 - Calendário Académico");
            WriteLine("2 - Regulamentos e Documentação");
            WriteLine("3 - Histórico e Relatórios");
            WriteLine("4 - Avaliação da Mobilidade");
            WriteLine("5 - Estórias da Mobilidade");
            WriteLine("\n9 - Voltar");
            WriteLine("\n0 - Sair");

            do
            {
                decisao = ReadKey(true);
                Console.Clear();

                // Calendário académico
                if (decisao.Key == ConsoleKey.D1)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[12].Length) / 2) + menus[12].Length) - 2) + "}", menus[12]));
                    string calendarioacademico = File.ReadAllText("calendarioacademico.txt", Encoding.Default);
                    WriteLine("\n\n" + calendarioacademico);
                    WriteLine("\n\n9 - Voltar");
                    WriteLine("0 - Sair");
                    decisao = ReadKey(true);
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
                // Regulamentos e Documentação
                if (decisao.Key == ConsoleKey.D2)
                {
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[13].Length) / 2) + menus[13].Length) - 2) + "}", menus[13]));
                    WriteLine("\n\n1 - Documentos institucionais e distinções");
                    WriteLine("2 - Acordos bilaterais");
                    WriteLine("3 - Regulamentos");
                    WriteLine("4 - Modelos de Curriculum Vitae");
                    WriteLine("5 - Formulários de Mobilidade");
                    WriteLine("\n9 - Voltar");
                    WriteLine("0 - Sair");

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
                            while (decisao.Key != ConsoleKey.D0 || decisao.Key != ConsoleKey.D9);
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
                    while (decisao.Key != ConsoleKey.D9 || decisao.Key != ConsoleKey.D0);
                }
                // Historico e Relatórios
                if (decisao.Key == ConsoleKey.D3)
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
                }
                // Avaliação da Mobilidade
                if (decisao.Key == ConsoleKey.D4)
                {
                    Console.Clear();
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                    WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                    WriteLine("\n\n1 - {0}", menus[24]);
                    WriteLine("2 - {0}", menus[25]);
                    WriteLine("3 - {0}", menus[26]);
                    WriteLine("4 - {0}", menus[27]);
                    WriteLine("5 - {0}", menus[28]);
                    WriteLine("6 - {0}", menus[29]);
                    WriteLine("7 - {0}", menus[30]);
                    WriteLine("\n9 - Voltar");
                    WriteLine("\n0 - Sair");

                    do
                    {
                        decisao = ReadKey(true);
                        if (decisao.Key == ConsoleKey.D1)
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[24].Length) / 2) + menus[24].Length) - 2) + "}", menus[24]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2015/2016");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2015/2016");
                            WriteLine("3 - Programa Erasmus: estudantes enviados para estágios 2015/2016");
                            WriteLine("4 - Programa Erasmus: docentes e colaboradores enviados 2015/2016");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes recebidos 2015/2016");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc1.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc1.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc1.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc1.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc1.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[25].Length) / 2) + menus[25].Length) - 2) + "}", menus[25]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2014/2015");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2014/2015");
                            WriteLine("3 - Programa Erasmus: estudantes enviados para estágios 2014/2015");
                            WriteLine("4 - Programa Erasmus: docentes e colaboradores enviados 2014/2015");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes recebidos 2014/2015");
                            WriteLine("6 - Mobilidade internacional (extracomunitária): estudantes enviados 2014/2015");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc2.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc2.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc2.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc2.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc2.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D6)
                                {
                                    pdfile = "avamobdoc2.6.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                        if (decisao.Key == ConsoleKey.D3)
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[26].Length) / 2) + menus[26].Length) - 2) + "}", menus[26]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2013/2014");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2013/2014");
                            WriteLine("3 - Programa Erasmus: estudantes enviados para estágios 2013/2014");
                            WriteLine("4 - Programa Erasmus: docentes e colaboradores enviados 2013/2014");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes recebidos 2013/2014");
                            WriteLine("6 - Mobilidade internacional (extracomunitária): estudantes enviados 2013/2014");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc3.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc3.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc3.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc3.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc3.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D6)
                                {
                                    pdfile = "avamobdoc3.6.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                        if (decisao.Key == ConsoleKey.D4)
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[27].Length) / 2) + menus[27].Length) - 2) + "}", menus[27]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2012/2013");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2012/2013");
                            WriteLine("3 - Programa Erasmus: estudantes enviados para estágios 2012/2013");
                            WriteLine("4 - Programa Erasmus: docentes e colaboradores enviados 2012/2013");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes recebidos 2012/2013");
                            WriteLine("6 - Mobilidade internacional (extracomunitária): estudantes enviados 2012/2013");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc4.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc4.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc4.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc4.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc4.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D6)
                                {
                                    pdfile = "avamobdoc4.6.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                        if (decisao.Key == ConsoleKey.D5)
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[28].Length) / 2) + menus[28].Length) - 2) + "}", menus[28]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2011/2012");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2011/2012");
                            WriteLine("3 - Programa Erasmus: estudantes enviados para estágios 2011/2012");
                            WriteLine("4 - Programa Erasmus: docentes e colaboradores enviados 2011/2012");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes recebidos 2011/2012");
                            WriteLine("6 - Mobilidade internacional (extracomunitária): estudantes enviados 2011/2012");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc5.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc5.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc5.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc5.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc5.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D6)
                                {
                                    pdfile = "avamobdoc5.6.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                        if (decisao.Key == ConsoleKey.D6)
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[29].Length) / 2) + menus[29].Length) - 2) + "}", menus[29]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2010/2011");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2010/2011");
                            WriteLine("3 - Programa Erasmus: docentes e colaboradores enviados 2010/2011");
                            WriteLine("4 - Mobilidade internacional (extracomunitária): estudantes recebidos 2010/2011");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes enviados 2010/2011");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc5.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc5.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc5.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc5.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc5.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                        if (decisao.Key == ConsoleKey.D7)
                        {
                            Console.Clear();
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[1].Length) / 2) + menus[1].Length) - 2) + "}", menus[1]));
                            WriteLine(string.Format("{0," + ((((WindowWidth - menus[15].Length) / 2) + menus[15].Length) - 2) + "}", menus[15]));
                            WriteLine(string.Format("\n{0," + ((((WindowWidth - menus[30].Length) / 2) + menus[30].Length) - 2) + "}", menus[30]));
                            WriteLine("\n\n1 - Programa Erasmus: estudantes recebidos 2009/2010");
                            WriteLine("2 - Programa Erasmus: estudantes enviados 2009/2010");
                            WriteLine("3 - Programa Erasmus: docentes e colaboradores enviados 2009/2010");
                            WriteLine("4 - Mobilidade internacional (extracomunitária): estudantes recebidos 2009/2010");
                            WriteLine("5 - Mobilidade internacional (extracomunitária): estudantes enviados 2009/2010");
                            WriteLine("\n9 - Voltar");
                            WriteLine("\n0 - Sair");

                            do
                            {
                                decisao = ReadKey(true);
                                if (decisao.Key == ConsoleKey.D1)
                                {
                                    pdfile = "avamobdoc6.1.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D2)
                                {
                                    pdfile = "avamobdoc6.2.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D3)
                                {
                                    pdfile = "avamobdoc6.3.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D4)
                                {
                                    pdfile = "avamobdoc6.4.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D5)
                                {
                                    pdfile = "avamobdoc6.5.pdf";
                                    LeitorPDF.pdfreader(pdfile);
                                }
                                if (decisao.Key == ConsoleKey.D9)
                                {
                                    SubMenu_IG_AvaliacaoMobilidade.submenu_InfoGeral_1_1_3(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                    while (decisao.Key != ConsoleKey.D0);
                }
                if (decisao.Key == ConsoleKey.D5)
                {

                }
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

                    if (decisao.Key == ConsoleKey.S)
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
