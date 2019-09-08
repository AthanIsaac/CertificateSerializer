using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace VisualSerializer
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }

        private static string regex = @"No.[0-9]{4}-M[0-9]*";

        private class SerialNumber
        {
            public int year;
            public int id;
            public SerialNumber(string serialNumber)
            {
                year = int.Parse(serialNumber.Substring(3, 4));
                id = int.Parse(serialNumber.Substring(9));
            }

            public SerialNumber Increment()
            {
                SerialNumber s = new SerialNumber(this.ToString());
                s.id++;
                return s;
            }
            public SerialNumber Increment(int n)
            {
                SerialNumber s = new SerialNumber(this.ToString());
                s.id+=n;
                return s;
            }

            public override string ToString()
            {
                return "No." + year + "-M" + id.ToString("000");
            }
        }

        //Find and Replace Method
        private static void FindAndReplace(Word.Application wordApp, object toFind, object toReplace)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundLike = false;
            object nmatchAllforms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref toFind,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundLike,
                ref nmatchAllforms, ref forward,
                ref wrap, ref format, ref toReplace,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        //Creeate the Doc Method
        public static void CreateWordDocument(object filename, object SaveAs, int copies)
        {
            Word.Application wordApp = new Word.Application();
            Word.Document myWordDoc = null;
            wordApp.Visible = false;

            if (!File.Exists((string)filename))
            {
                throw new IOException("File not Found!");
            }

            for (int i = 1; i <= copies; i++)
            {
                myWordDoc = wordApp.Documents.Open(ref filename);

                myWordDoc.Activate();

                //find and replace
                SerialNumber oldNum = new SerialNumber(WhatToFind(myWordDoc));
                FindAndReplace(wordApp, oldNum.ToString(), oldNum.Increment(i).ToString());

                //Save as
                object name = (string)SaveAs + i;
                myWordDoc.SaveAs2(ref name);

                myWordDoc.Close();
            }
            wordApp.Quit();
            Console.WriteLine("Success!");
        }

        private static string WhatToFind(Word.Document myWordDoc)
        {
            string text = myWordDoc.Content.Text;
            return Regex.Match(text, regex).ToString();
        }



    }
}
