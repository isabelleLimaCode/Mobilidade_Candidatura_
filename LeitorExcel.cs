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
    class LeitorExcel
    {
        public static void lerExcel(ref int linhas, ref int colunas, dynamic[,] excel, string[] menus)
        {
            string aviso = ("Aviso - Aguarde enquanto os dados são carregados...");
            Console.Clear();
            WriteLine(string.Format("{0," + ((((WindowWidth - menus[7].Length) / 2) + menus[7].Length) - 2) + "}", menus[7]));
            WriteLine();
            WriteLine(string.Format("{0," + ((((WindowWidth - aviso.Length) / 2) + aviso.Length) - 2) + "}", aviso));

            Application Excelapp = new Application();

            if (Excelapp == null)
            {
                WriteLine("Não tem o Excel instalado!!!");
            }

            Workbook excelbook = Excelapp.Workbooks.Open(@"C:\candidatura_erasmus.xlsx");
            _Worksheet excelsheet = excelbook.Sheets[1];
            Range excelrange = excelsheet.UsedRange;

            linhas = excelrange.Rows.Count;
            colunas = excelrange.Columns.Count;
            string[] titulos = new string[colunas];

            for (int i = 1; i <= linhas; i++)
            {
                for (int j = 1; j <= colunas; j++)
                {
                    if (excelrange.Cells[i, j] != null && excelrange.Cells[i, j].Value2 != null)
                    {
                        excel[i, j] = excelrange.Cells[i, j].Value2;
                    }
                }
            }
            Excelapp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Excelapp);
        }
    }
}
