using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using W = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp1
{
    public partial class frm : Form
    {
        public frm()
        {
            InitializeComponent();
            
        }
        private void frm_Load(object sender, EventArgs e)
        {

        }
        private void btn1_Click(object sender, EventArgs e)
        {
            object EndOfDoc = "\\endofdoc";
            W.Application oWord = new W.Application();
            W.Document oDoc = oWord.Documents.Add();
            object ObjMissing = Missing.Value; 

            W.Paragraph oPrg = oDoc.Paragraphs.Add();

            oPrg.Range.Text = "Федеральное государственное бюджетное образовательное учреждение высшего образования";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Times New Roman";
            oPrg.Range.Font.Size = 14f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "«Российский университет транспорта» (РУТ (МИИТ))";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Times New Roman";
            oPrg.Range.Font.Size = 14f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "Институт транспортной техники и систем управления";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Times New Roman";
            oPrg.Range.Font.Size = 16f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "Кафедра \"Управление и защита информации\"";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 1;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 0;
            oPrg.Range.Font.Name = "Times New Roman";
            oPrg.Range.Font.Size = 16f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "Отчёт";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 1;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 1;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Times New Roman";
            oPrg.Range.Font.Size = 26f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "по учебной практике";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Times New Roman";
            oPrg.Range.Font.Size = 14f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            

            oWord.Visible = true;


            table(EndOfDoc, oDoc, ref ObjMissing);

            /*
            //Таблица
            oDoc.PageSetup.TopMargin = 0.75f / 0.03f;
            oDoc.PageSetup.BottomMargin = 0.75f / 0.03f;
            Object start = 250;
            Object end = 250;
            W.Range wordrange = oDoc.Range(ref start, ref end);
            Object defaultTableBehavior = W.WdDefaultTableBehavior.wdWord9TableBehavior;
            Object autoFitBehavior = W.WdAutoFitBehavior.wdAutoFitWindow;
            //Добавляем таблицу и получаем объект wordtable
            W.Table wordtable = oDoc.Tables.Add(wordrange, 2, 2, ref defaultTableBehavior, ref autoFitBehavior);
            W.Range wordcellrange = oDoc.Tables[1].Cell(1, 2).Range;
            wordcellrange.Borders[WdBorderType.wdBorderRight].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderLeft].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderTop].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderBottom].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange = wordtable.Cell(1, 1).Range;
            wordcellrange.Text = " ";
            wordcellrange.Borders[WdBorderType.wdBorderLeft].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderTop].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderBottom].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange = wordtable.Cell(1, 1).Range;
            wordcellrange.Text = "Выполнил:";
            wordcellrange = wordtable.Cell(2, 1).Range;
            wordcellrange.Text = "Проверил:";
            wordcellrange.Borders[WdBorderType.wdBorderRight].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderLeft].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderBottom].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange = wordtable.Cell(1, 2).Range;
            wordcellrange.Text = "Студент";
            wordcellrange = wordtable.Cell(2, 2).Range;
            wordcellrange.Text = "Преподаватель";
            wordcellrange.Borders[WdBorderType.wdBorderRight].ColorIndex = W.WdColorIndex.wdWhite;
            wordcellrange.Borders[WdBorderType.wdBorderBottom].ColorIndex = W.WdColorIndex.wdWhite;
            */

        }

        private void table(object EndOfDoc, W.Document ObjDoc, ref object ObjMissing)
        {
            W.Table ObjTable;
            W.Range ObjWordRange;
            ObjWordRange = ObjDoc.Bookmarks.get_Item(ref EndOfDoc).Range;
            ObjTable = ObjDoc.Tables.Add(ObjWordRange, 6, 2, ref ObjMissing, ref ObjMissing);
            ObjTable.Range.ParagraphFormat.SpaceAfter = 0;
            ObjTable.Cell(1, 1).Range.Text = "Выполнил: ";
            ObjTable.Cell(1, 1).Width = 1;
        }

    }

}
