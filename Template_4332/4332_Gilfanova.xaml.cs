using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4332
{
    /// <summary>
    /// Interaction logic for _4332_Gilfanova.xaml
    /// </summary>
    public partial class _4332_Gilfanova : Window
    {
        public Excel.Range xlSheetRange;

        public _4332_Gilfanova()
        {
            InitializeComponent();
        }
        private void export_word_elina_Click(object sender, RoutedEventArgs e)
        {
            List<Gilfanova_4332_10variant> allusers;
            string role = "";
            using (GilfanovaContext usersEntities = new GilfanovaContext())
            {
                allusers = usersEntities.Gilfanova_4332_10variant.ToList().OrderBy(p => p.Role).ToList();
                var app = new Word.Application();
                Word.Document document = app.Documents.Add();
                for (int a = 1; a < 4; a++)
                {
                    Word.Paragraph paragraph = document.Paragraphs.Add();
                    Word.Range range = paragraph.Range;
                    range.Text = "Роль " + Convert.ToString(a);
                    string worksheet = "Роль " + Convert.ToString(a);
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();
                    Word.Paragraph tableParagraph = document.Paragraphs.Add();
                    Word.Range tableRange = tableParagraph.Range;
                    Word.Table uTabel = document.Tables.Add(tableRange, 5, 4);
                    uTabel.Borders.InsideLineStyle = uTabel.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    uTabel.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    Word.Range cellRange;
                    cellRange = uTabel.Cell(1, 1).Range;
                    cellRange.Text = "Роль";
                    cellRange = uTabel.Cell(1, 2).Range;
                    cellRange.Text = "ФИО";
                    cellRange = uTabel.Cell(1, 3).Range;
                    cellRange.Text = "Логин";
                    cellRange = uTabel.Cell(1, 4).Range;
                    cellRange.Text = "Пароль";
                    uTabel.Rows[1].Range.Bold = 1;
                    uTabel.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    int i = 1;
                    foreach (Gilfanova_4332_10variant user in allusers)
                    {
                        if (user.Role == "Администратор") { role = "Роль 1"; }
                        if (user.Role == "Менеджер") { role = "Роль 2"; }
                        if (user.Role == "Клиент") { role = "Роль 3"; }
                        if (role == worksheet)
                        {
                            cellRange = uTabel.Cell(i + 1, 1).Range;
                            cellRange.Text = user.Role;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            cellRange = uTabel.Cell(i + 1, 2).Range;
                            cellRange.Text = user.FIO;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            cellRange = uTabel.Cell(i + 1, 3).Range;
                            cellRange.Text = user.Login;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            cellRange = uTabel.Cell(i + 1, 4).Range;
                            cellRange.Text = user.Password;
                            cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            i++;
                        }
                    }
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
                }
                MessageBox.Show("Данные экспортированы в Word");
                app.Visible = true;
                document.SaveAs2(@"D:\Desktop\elina.docx");
            }
        }

        private void import_json_elina_Click(object sender, RoutedEventArgs e)
        {

            using (GilfanovaContext usersEntities = new GilfanovaContext())
            {
                string jsonFilePath = @"D:\Desktop\Импорт3лр\5.json";
                string json = File.ReadAllText(jsonFilePath);
                json = json.Substring(0, json.Length - 1);
                string[] obj = json.Split('}');
                string a = "";
                foreach (string s in obj)
                {
                    a = s + "}"; a = a.Substring(1);
                    if (a != "")
                    {
                        Gilfanova_4332_10variant us = JsonConvert.DeserializeObject<Gilfanova_4332_10variant>(a);
                        usersEntities.Gilfanova_4332_10variant.Add(new Gilfanova_4332_10variant()
                        {
                            Role = us.Role,
                            FIO = us.FIO,
                            Login = us.Login,
                            Password = GetHashString(us.Password)
                        });
                        usersEntities.SaveChanges();
                    }
                }
                MessageBox.Show("Данные импортированы");

            }
        }
        private string GetHashString(string s)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(s);

            MD5CryptoServiceProvider CSP = new
            MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            string hash = "";
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            return hash;
        }
        private void import_elina_Click(object sender, RoutedEventArgs e)
        {
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"D:\загрузкиD\Импорт\5.xlsx");
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
            using (GilfanovaContext usersEntities = new GilfanovaContext())
            {
                for (int i = 1; i < _rows; i++)
                {
                    usersEntities.Gilfanova_4332_10variant.Add(new Gilfanova_4332_10variant()
                    {
                        Role = list[i, 0],
                        FIO = list[i, 1],
                        Login = list[i, 2],
                        Password = GetHashString(list[i, 3])
                    });
                }
                usersEntities.SaveChanges();
                MessageBox.Show("Данные импортированы");
            }
        }
        private void export_elina_Click(object sender, RoutedEventArgs e)
        {
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = 3;
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            using (GilfanovaContext usersEntities = new GilfanovaContext())
            {
                var admins = usersEntities.Gilfanova_4332_10variant.Where(p => p.Role == "Администратор");
                for (int i = 0; i < admins.Count(); i++)
                {
                    Excel.Worksheet worksheet = app.Worksheets.Item[1];

                    //выбираем лист на котором будем работать (Лист 1)
                    worksheet = (Excel.Worksheet)app.Sheets[1];
                    //Название листа
                    worksheet.Name = "Администраторы";
                    int startRowIndex = 1;
                    worksheet.Cells[1][startRowIndex] = "Роль";
                    worksheet.Cells[2][startRowIndex] = "ФИО";
                    worksheet.Cells[3][startRowIndex] = "Логин";
                    worksheet.Cells[4][startRowIndex] = "Пароль";
                    startRowIndex++;

                    foreach (Gilfanova_4332_10variant admin in admins)
                    {
                        worksheet.Cells[1][startRowIndex] = admin.Role;
                        worksheet.Cells[2][startRowIndex] = admin.FIO;
                        worksheet.Cells[3][startRowIndex] = admin.Login;
                        worksheet.Cells[4][startRowIndex] = admin.Password;
                        startRowIndex++;
                        Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[4][startRowIndex - 1]];
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight]
                        .LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        worksheet.Columns.AutoFit();
                    }
                    var managers = usersEntities.Gilfanova_4332_10variant.Where(a => a.Role == "Менеджер");
                    for (int j = 0; j < managers.Count(); j++)
                    {
                        Excel.Worksheet worksheet2 = app.Worksheets.Item[2];

                        //выбираем лист на котором будем работать (Лист 2)
                        worksheet2 = (Excel.Worksheet)app.Sheets[2];
                        //Название листа
                        worksheet2.Name = "Менеджеры";
                        int startRowIndex2 = 1;
                        worksheet2.Cells[1][startRowIndex2] = "Роль";
                        worksheet2.Cells[2][startRowIndex2] = "ФИО";
                        worksheet2.Cells[3][startRowIndex2] = "Логин";
                        worksheet2.Cells[4][startRowIndex2] = "Пароль";
                        startRowIndex2++;

                        foreach (Gilfanova_4332_10variant manager in managers)
                        {
                            worksheet2.Cells[1][startRowIndex2] = manager.Role;
                            worksheet2.Cells[2][startRowIndex2] = manager.FIO;
                            worksheet2.Cells[3][startRowIndex2] = manager.Login;
                            worksheet2.Cells[4][startRowIndex2] = manager.Password;
                            startRowIndex2++;

                            Excel.Range rangeBorders2 = worksheet2.Range[worksheet2.Cells[1][1], worksheet2.Cells[4][startRowIndex2 - 1]];
                            rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeRight]
                            .LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                            worksheet2.Columns.AutoFit();
                        }
                        var clients = usersEntities.Gilfanova_4332_10variant.Where(c => c.Role == "Клиент");
                        for (int k = 0; k < clients.Count(); k++)
                        {
                            Excel.Worksheet worksheet3 = app.Worksheets.Item[3];

                            //выбираем лист на котором будем работать (Лист 2)
                            worksheet3 = (Excel.Worksheet)app.Sheets[3];
                            //Название листа
                            worksheet3.Name = "Клиенты";
                            int startRowIndex3 = 1;
                            worksheet3.Cells[1][startRowIndex3] = "Роль";
                            worksheet3.Cells[2][startRowIndex3] = "ФИО";
                            worksheet3.Cells[3][startRowIndex3] = "Логин";
                            worksheet3.Cells[4][startRowIndex3] = "Пароль";
                            startRowIndex3++;

                            foreach (Gilfanova_4332_10variant client in clients)
                            {
                                worksheet3.Cells[1][startRowIndex3] = client.Role;
                                worksheet3.Cells[2][startRowIndex3] = client.FIO;
                                worksheet3.Cells[3][startRowIndex3] = client.Login;
                                worksheet3.Cells[4][startRowIndex3] = client.Password;
                                startRowIndex3++;

                                Excel.Range rangeBorders2 = worksheet3.Range[worksheet3.Cells[1][1], worksheet3.Cells[4][startRowIndex3 - 1]];
                                rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                                rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlEdgeRight]
                                .LineStyle = rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                                rangeBorders2.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                                worksheet3.Columns.AutoFit();
                            }
                        }
                    }
                }
                MessageBox.Show("Файл создан");
                app.Visible = true;
            }
        }
    }
}
