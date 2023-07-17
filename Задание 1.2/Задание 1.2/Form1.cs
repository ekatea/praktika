using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using W = Microsoft.Office.Interop.Word;

namespace Задание_1._2
{ 
    public partial class Form1 : Form
    {
        public Form1()
        {
        InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Object EndOfDoc = "\\endofdoc";
            W.Application oWord = new W.Application();
            W.Document oDoc = oWord.Documents.Add();

            W.Paragraph oPrg = oDoc.Paragraphs.Add();

            oPrg.Range.Text = "Договор";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //все заглавные
            oPrg.Range.Font.AllCaps = 1;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "коммерческой концессии";
            //отступа слева нет
            oPrg.Format.LeftIndent = 0;
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //все обычные, не заглавные
            oPrg.Range.Font.AllCaps = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание по центру
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphCenter;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "г.\t ";
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints (3.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertAfter("\t");
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(11.5f), WdAlignmentTabAlignment.wdRight);

            oPrg.Range.InsertAfter("«\t»");
            oPrg.TabStops.Add(oWord.CentimetersToPoints(13f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);
            
            oPrg.Range.InsertAfter("\t");
            oPrg.TabStops.Add(oWord.CentimetersToPoints(15.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertAfter("20\tг.");
            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();
            
            oPrg.Range.Text = "«Правообладатель»\t, ";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;
            
            oPrg.TabStops.Add(oWord.CentimetersToPoints(8.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertAfter("в лице\t, действующего");
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "на основании\t, и";
            //жирный шрифт
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(9f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "«Пользователь»\t, ";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(8.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertAfter("в лице\t, действующего на");
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "1. Правообладатель передает во временное и платное распоряжение Пользователю такие";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "объекты интеллектуальной собственности \t (товарный знак,";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "торговую марку, технологию и т.д.).";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "2. Пользователь обязуется использовать полученные права для \t";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "\t (производства определенных товаров или";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "оказания услуг)";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "3. За предоставленные Правообладателем права Пользователь выплачивает вознаграждение в";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "таком порядке: \t.";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "4. Договор коммерческой концессии подлежит государственной регистрации в установленном";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "законодательством порядке.";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "5. Правобладатель имеет следующие права и обязанности: \t";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "\t.";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "6. У пользователя имеются такие права и обязанности:\t";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oPrg.Range.Text = "\t";
            //жирный шрифт\
            oPrg.Range.Font.Bold = 0;
            //шрифт и размер шрифта
            oPrg.Range.Font.Name = "Calibri";
            oPrg.Range.Font.Size = 12f;
            //выравнивание 
            oPrg.Alignment = W.WdParagraphAlignment.wdAlignParagraphLeft;

            oPrg.TabStops.Add(oWord.CentimetersToPoints(16.5f), WdAlignmentTabAlignment.wdRight,
            W.WdTabLeader.wdTabLeaderLines);

            oPrg.Range.InsertParagraphAfter();

            oWord.Visible = true;
        }
    }
}