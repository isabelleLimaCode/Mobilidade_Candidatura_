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
    class Login
    {
        public class log
        {
            public string[] usernames { get; set; }
            public string[] passwords { get; set; }

            public log(string[] usernames, string[] passwords)
            {
                this.usernames = usernames;
                this.passwords = passwords;
            }
        }

        public static void login(string[] usernames, string[] passwords, ref string username, ref string password)
        {
            do
            {
                WriteLine("Introduza o nome de utilizador:");
                username = ReadLine();
                if (username == usernames[0] || username == usernames[1] || username == usernames[2])
                {
                    WriteLine();
                    do
                    {
                        WriteLine("Introduza a password:");
                        password = ReadLine();
                        if ((username == usernames[0] && password == passwords[0]) || (username == usernames[1] && password == passwords[1])
                            || (username == usernames[2] && password == passwords[2]))
                        {
                            WriteLine("Login bem sucedido!");
                        }
                        else
                        {
                            WriteLine("\nPassword incorreta!");
                            WriteLine("Tente outra vez.\n");
                        }
                    }
                    while (password != username);
                }
                else
                {
                    WriteLine("\nO nome de utilizador que introduziu não existe.");
                    WriteLine("Tente outra vez.\n");
                }
            }
            while (username != usernames[0] && username != usernames[1] && username != usernames[2]);
        }
    }
}
