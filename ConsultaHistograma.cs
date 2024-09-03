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
    class ConsultaHistograma
    {
        public static void consultaHistograma(ref int linhas, ref int colunas, dynamic[,] excel, string[] usernames, string[] passwords, ref string username, ref string password, ref string estado, string[] menus,
                                ref int number, ref string pdfile)
        {
            Clear();
            string aviso = "Dados carregados com sucesso!";
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[7].Length) / 2) + menus[7].Length) - 2) + "}", menus[7]));
            WriteLine();
            WriteLine(string.Format("{0," + ((((WindowWidth - aviso.Length) / 2) + aviso.Length) - 2) + "}", aviso));
            WriteLine("\n\n1 - Consultar dados de candidatura Erasmus");
            WriteLine("\n2 - Consultar Histogramas");
            WriteLine("\n9 - Voltar");
            WriteLine("\n0 - Sair");

            int contador = 0;
            string[] titulos = new string[colunas];

            ConsoleKeyInfo decisao;
            do
            {
                decisao = ReadKey(true);
                if (decisao.Key == ConsoleKey.D1)
                {
                    contador = 0;
                    do
                    {
                        Clear();
                        for (int i = 1; i <= colunas; i++)
                        {
                            titulos[contador] = (i + " - " + excel[1, i].ToString());
                            WriteLine(titulos[contador]);
                        }
                        WriteLine();

                        int numeroCampos;
                        WriteLine("Digite o número de campos que quer consultar:");
                        numeroCampos = Convert.ToInt32(ReadLine());

                        int[] campos = new int[numeroCampos];

                        WriteLine("Escolha os campos que pretende consultar digitando os números correspondentes, um de cada vez:");
                        for (int i = 0; i < campos.Length; i++)
                        {
                            campos[i] = Convert.ToInt32(ReadLine());
                        }
                        Clear();

                        for (int i = 1; i <= linhas; i++)
                        {
                            for (int j = 0; j < campos.Length; j++)
                            {
                                int tamanho = 10;
                                string resultado = excel[i, campos[j]].ToString();
                                SetCursorPosition(tamanho * j, i);

                                if (j != 0)
                                {
                                    Write("  ");
                                }
                                Write(resultado);
                            }
                            WriteLine();
                        }
                        WriteLine("\n\nPretende fazer uma nova consulta?");
                        WriteLine("\n1 - Sim");
                        WriteLine("\n9 - Voltar");
                        WriteLine("\n0 - Sair");

                        decisao = ReadKey(true);

                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenuListaUtilizadores.submenu_ListaUtilizadores(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                if (decisao.Key == ConsoleKey.D2)
                {
                    do
                    {
                        Clear();
                        for (int i = 1; i <= colunas; i++)
                        {
                            titulos[contador] = (i + " - " + excel[1, i].ToString());
                            WriteLine(titulos[contador]);
                        }
                        WriteLine();
                        WriteLine("Digite o número respetivo da coluna pretendida:");
                        number = Convert.ToInt32(ReadLine());

                        //Histograma(ref linhas, excel, ref number);
                        Dictionary<string, int> histograma = new Dictionary<string, int>();
                        {
                            dynamic key = null;
                            int value = 0;

                            for (int i = 2; i <= linhas; i++)
                            {
                                key = excel[i, number];
                                foreach (var item in colunas.ToString())
                                {
                                    if (!histograma.ContainsKey(key.ToString()))
                                    {
                                        histograma.Add(key.ToString(), value = 1);
                                        break;
                                    }
                                    else
                                    {
                                        histograma[key.ToString()]++;
                                        break;
                                    }
                                }
                            }
                        }
                        Clear();
                        foreach (KeyValuePair<string, int> item in histograma)
                        {
                            WriteLine("Key: " + item.Key + "  Value: " + item.Value);
                        }
                        WriteLine("\n\nPretende fazer uma nova consulta?");
                        WriteLine("\n1 - Sim");
                        WriteLine("\n9 - Voltar");
                        WriteLine("\n0 - Sair");
                        decisao = ReadKey(true);

                        if (decisao.Key == ConsoleKey.D9)
                        {
                            SubMenuListaUtilizadores.submenu_ListaUtilizadores(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
                if (decisao.Key == ConsoleKey.D9)
                {
                    MenuGestor.menuGestor(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
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
    }
}
