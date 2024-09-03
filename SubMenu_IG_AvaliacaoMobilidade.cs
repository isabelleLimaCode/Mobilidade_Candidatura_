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
    class SubMenu_IG_AvaliacaoMobilidade
    {
        public static ConsoleKeyInfo submenu_InfoGeral_1_1_3(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
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
            ConsoleKeyInfo decisao;

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
            return decisao;
        }
    }
}
