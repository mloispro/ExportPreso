using Novacode;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportPreso
{
    public class WordEx
    {
        static DocX _doc;
        static Formatting _titleFormat;
        static Formatting _headerFormat;
        static Formatting _textFormat;
        static string _filePath;
        static System.Drawing.FontFamily _calibri;
        static WordEx()
        {
            _calibri = new System.Drawing.FontFamily("Calibri");
            _titleFormat = new Formatting();
            _titleFormat.FontFamily = _calibri;
            _titleFormat.Size = 12D;
            _titleFormat.FontColor = System.Drawing.Color.DarkBlue;

            _headerFormat = new Formatting();
            _headerFormat.FontFamily = _calibri;
            _headerFormat.Size = 12D;
            _headerFormat.FontColor = System.Drawing.Color.Blue;
            //_headerFormat.Position = 12;

            // A formatting object for our normal paragraph text:
            _textFormat = new Formatting();
            _textFormat.FontFamily = _calibri;
            _textFormat.Size = 10D;
        }
        public static void CreateDoc(string filePath)
        {
            _filePath = filePath;

            _doc = DocX.Create(filePath);
            _doc.ApplyTemplate(@"C:\Projects\ExportPreso\ExportPreso\WordTemplate\_template.dotx");

        }
        public static void AddText(string text, bool isNote = false)
        {
            if (isNote && IsNumeric(text))
                return;

            var para = _doc.InsertParagraph(CleanInvalidXmlChars(text));//, false, _textFormat);
            //var para = _doc.InsertParagraph(text);
            para.StyleName = "Normal";
            if (isNote)
            {
                para.StyleName = "ProfNote";
            }
        }
        public static void AddHeader(string text)
        {
            //_doc.InsertParagraph(text, false, _headerFormat);
            var para = _doc.InsertParagraph(CleanInvalidXmlChars(text));//, false, _headerFormat);
            para.StyleName = "Heading2";
        }
        public static void AddTitle(string text)
        {
            //_doc.InsertParagraph(text, false, _headerFormat);
            var para = _doc.InsertParagraph(CleanInvalidXmlChars(text));//, false, _headerFormat);
            para.StyleName = "Heading1";
        }
        public static void AddBulletList(List<string> bullets, bool isNote = false)
        {
            var list = _doc.AddList(listType: ListItemType.Bulleted);

            foreach (var bullet in bullets)
            {
                if (isNote && IsNumeric(bullet))
                    continue;

                var li = _doc.AddListItem(list, CleanInvalidXmlChars(bullet));

            }
            foreach (var item in list.Items)
            {
                item.StyleName = "ListParagraph";
                if (isNote)
                {
                    item.StyleName = "ProfNoteBullet";
                }
            }

            //_doc.InsertList(list, _calibri, 10D);
            _doc.InsertList(list);
        }
        public static void AddImage(string text)
        {
            var image = _doc.AddImage(text);
            //para.StyleName = "ListParagraph";
        }
        public static void AddBlankLine()
        {
            var para = _doc.InsertParagraph();
            para.StyleName = "Normal";
        }
        public static void Save()
        {
            try
            {
                _doc.Save();
                Process.Start("WINWORD.EXE", "\"" + _filePath + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static string CleanInvalidXmlChars(string text)
        {
            string re = @"[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000-x10FFFF]";
            return Regex.Replace(text, re, "");
        }
        public static bool IsNumeric(string text)
        {
            int result;
            if (int.TryParse(text, out result))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
