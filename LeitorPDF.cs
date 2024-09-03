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
    class LeitorPDF
    {

        public class pdf
        {
            public string pdfile { get; set; }

            public pdf(string pdfile)
            {
                this.pdfile = pdfile;
            }
        }


        public static void pdfreader(string pdfile)
        {

            using (Process pdf = new Process())
            {
                pdf.StartInfo.FileName = pdfile;
                pdf.Start();
            }
        }
    }
}
