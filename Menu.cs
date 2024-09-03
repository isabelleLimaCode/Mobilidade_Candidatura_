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
    class Menu
    {
        public class MenuItem
        {
            public char opcao { get; set; }   
            public string item { get; set; }

            public MenuItem(char opcao, string item)
            {
                this.opcao = opcao;
                this.item = item;
            }

            public void print()
            {
                Console.WriteLine("| \t" + opcao + ": " + item);
            }
        }

        public Dictionary<char, MenuItem> opcoes = new Dictionary<char, MenuItem>();

        public string titulo = "Sem T+itulo";

        public string descricao = "sem descricao";

        public MenuItem exec()
        {
            Console.WriteLine("              === " + titulo + " === \n");
            Console.WriteLine(descricao + "\n");

            Console.WriteLine("/-----------------------------------------------------\\");
            foreach (MenuItem opcao in opcoes.Values)
            {
                opcao.print();
            }
            Console.WriteLine("\\-----------------------------------------------------/");
            char op = Console.ReadLine()[0];
            return opcoes[op];
        }

        public void addOpcao(char opcao, string item)
        {
            opcoes[opcao] = new MenuItem(opcao, item);
        }
    }
}
