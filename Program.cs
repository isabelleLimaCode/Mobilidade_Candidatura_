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
    class Program
    {
        #region Centrar e alterar a consola para ecrã inteiro em qualquer resolução
        [DllImport("kernel32.dll", ExactSpelling = true)]
        private static extern IntPtr GetConsoleWindow();
        private static IntPtr ThisConsole = GetConsoleWindow();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        private const int HIDE = 0;
        private const int MAXIMIZE = 3;
        private const int MINIMIZE = 6;
        private const int RESTORE = 9;
        #endregion

        static void Main(string[] args)
        {
            #region Centrar e alterar a consola para ecrã inteiro em qualquer resolução
            Console.SetWindowSize(Console.LargestWindowWidth, Console.LargestWindowHeight);
            ShowWindow(ThisConsole, MAXIMIZE);
            #endregion
            #region variáveis int
            int number = 0;
            int linhas = 0;
            int colunas = 0;
            #endregion
            #region variáveis string
            string username = null;
            string password = null;
            string pdfile = null;
            string estado = null;
            #endregion
            #region arrays string
            string[] usernames = { "gestor", "docente", "aluno" };
            string[] passwords = { "gestor", "docente", "aluno" };
            
            // menus[0 a 9 MenuGestor]
            // menus[11 a 30 MenuAluno] [MenuAluno inicial de 11 a 16] [MenuAluno Regulamentos e Documentação de 17 a 21]
            // [MenuAluno Relatórios de Atividades (Histórico e Relatórios 22) de 24 a 30]
            string[] menus = {/*0*/"Página Inicial - Conta: Gestor",/*1*/"Informação Geral",/*2*/"Candidatura",/*3*/"Contactos",/*4*/"Consultas",
                                   /*5*/"Gestão de Candidaturas",/*6*/"Registo de Novo Utilizador",/*7*/"Lista de Utilizadores",/*8*/"Dados de Login",
                                   /*9*/"Login",/*10*/"Página Inicial - Conta: Docente",/*11*/"Página Inicial - Conta: Aluno",/*12*/"Calendário Académico",
                                   /*13*/"Regulamentos e Documentação",/*14*/"Histórico e Relatórios",/*15*/"Avaliação da Mobilidade",/*16*/"Estórias de Mobilidade",
                                   /*17*/"Documentos Institucionais e Distinções",/*18*/"Acordos Bilaterais",/*19*/"Regulamentos",/*20*/"Modelos de Curriculum Vitae",
                                   /*21*/"Formulários de Mobilidade",/*22*/"Relatórios de Atividades 2006/2016",/*23*/"Avaliação da Mobilidade",/*24*/"2015/2016",
                                   /*25*/"2014/2015",/*26*/"2013/2014",/*27*/"2012/2013",/*28*/"2011/2012",/*29*/"2010/2011",/*30*/"2009/2010",
                                   /*31*/"Página Inicial - Conta: Visitante"};

            #endregion
            #region array dinâmico para guardar dados do excel
            dynamic[,] excel = new dynamic[1000, 1000];
            #endregion

            //Iniciar o programa como visitante
            ConsoleKeyInfo op = MenuVisitante.menuVisitante(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
            do
            {
                op = MenuVisitante.menuVisitante(ref linhas, ref colunas, excel, usernames, passwords, ref username, ref password, ref estado, menus,
                                ref number, ref pdfile);
            }
            while (op.Key != ConsoleKey.D0);

            /*Menu m = new Menu();
            m.addOpcao('a', "dfçgmlergmdeojge");
            m.addOpcao('%', "rgfhregijeglote");
            m.titulo = "O Meu Menu";
            Menu.MenuItem opcaoEscolhida = m.exec();
            opcaoEscolhida.print();*/
        }
    }
}