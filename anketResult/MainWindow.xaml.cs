using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace anketResult
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private Db db;
        public MainWindow()
        {
            db = new Db();
            if (db.Connect())
            {
                MessageBox.Show("Нет доступных серверов");
                this.Close();
            }
            InitializeComponent();
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var ComboGroups = db.DbSelect("SELECT spec.id, spec.name, spec.`key` FROM spec").Select();
                if (ComboGroups.Length > 0)
                    foreach (var Items in ComboGroups)
                        SelectSpec.Items.Add(Items.ItemArray[2]+" - "+ Items.ItemArray[1]);

            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (FileName.Text == "") FileName.Text = Convert.ToString(SelectSpec.SelectedValue);
            stud();
            FileName.Text = "";

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)

        {
            if (FileName.Text == "") FileName.Text = Convert.ToString(SelectSpec.SelectedValue);
            prep();
            FileName.Text = "";
        }

        private bool prep()
        {
            // Протокол для базы 11 класса
            {
                //Объект документа пдф
                iTextSharp.text.Document doc = new iTextSharp.text.Document();

                //Создаем объект записи пдф-документа в файл
                NewFile11:
                int newf = 2;
                try
                {
                    PdfWriter.GetInstance(doc, new FileStream("Преподаватели - " + FileName.Text + " - 11.pdf", FileMode.Create));
                }
                catch
                {
                    try
                    {
                        PdfWriter.GetInstance(doc, new FileStream("Преподаватели - " + FileName.Text + "11 (" + newf + ").pdf", FileMode.Create));
                    }
                    catch
                    {
                        newf++;
                        goto NewFile11;
                    }
                }

                doc.SetMargins(80, 40, 40, 30);
                doc.SetPageSize(new iTextSharp.text.Rectangle(PageSize.A4));

                //Открываем документ
                doc.Open();

                //Определение шрифта необходимо для сохранения кириллического текста
                //Иначе мы не увидим кириллический текст
                //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
                BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\times.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                BaseFont baseFontBold = BaseFont.CreateFont("C:\\Windows\\Fonts\\timesbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName = new iTextSharp.text.Font(baseFontBold, 14, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName2 = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.NORMAL);

                //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                PdfPTable table = new PdfPTable(4);

                float[] widths = new float[] { 1.1f, 12f, 6.5f, 4.5f };
                table.SetWidths(widths);
                table.WidthPercentage = 100;
                //Добавим первую строку
                PdfPCell cell = new PdfPCell(new Phrase("Эксперт " + expert.Text, font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку
                cell = new PdfPCell(new Phrase("Наименование основной профессиональной образовательной программы", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим строку с названием программы
                cell = new PdfPCell(new Phrase(Convert.ToString(SelectSpec.SelectedValue), font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    BorderWidthBottom = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим пустую строку
                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим Название "протокол"
                cell = new PdfPCell(new Phrase("Протокол", fontName))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку названия
                cell = new PdfPCell(new Phrase("анкетирования педагогических работников\nреализующих программу", fontName2))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим аннотацию
                string usa = "SELECT statprepcount(" + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + ",11)";
                var UsAll = db.DbSelect(usa).Select();
                string us = "SELECT statprepcountspec(" + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + ",11)";
                var Us = db.DbSelect(us).Select();

                //Добавим аннотацию
                cell = new PdfPCell(new Phrase("В анкетировании приняли участие " + Us[0].ItemArray[0] + " преподавателей, что составило " + Math.Round(Convert.ToDouble(UsAll[0].ItemArray[0]) * 100, 2) + "% от количества научно-педагогических работников, реализующих программу.", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим заголовок таблицы
                cell = new PdfPCell(new Phrase("Результаты анкетирования", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                // Добавляем заголовок таблицы
                cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Вопросы педагогическим работникам аккредитуемой программы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);

                //Добавляем все остальные ячейки
                for (int j = 17; j < 29; j++)
                {
                    if (j == 28)
                    {
                        cell = new PdfPCell(new Phrase("Результаты анкетирования\n ", font))
                        {
                            Colspan = 4,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        // Добавляем заголовок таблицы
                        cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Вопросы педагогическим работникам аккредитуемой программы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                    }
                    //var SpecCount = db.DbSelect("SELECT COUNT(anket_emp.id) AS expr1 FROM anket_emp INNER JOIN specemp ON anket_emp.iduser = specemp.idemp INNER JOIN ans ON anket_emp.ans = ans.id WHERE specemp.base = 11 AND specemp.idspec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + " AND ans.score >= 3 AND anket_emp.idque = "+j).Select();
                    var SpecCountAll = db.DbSelect("Select statprep(" + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + "," + j + ",11)").Select();
                    var Ques = db.DbSelect("SELECT ques.name FROM ques WHERE ques.id = " + j).Select();
                    var Ans = db.DbSelect("SELECT ans.text FROM ans INNER JOIN ques ON ans.idque = ques.id WHERE ques.id = " + j).Select();
                    cell = new PdfPCell(new Phrase(j - 16 + ".", font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(Ques[0].ItemArray[0].ToString(), font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    string an = "";
                    foreach (var Items in Ans)
                        an += "- " + Items[0] + "\n";
                    cell = new PdfPCell(new Phrase(an, font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(Math.Round(Convert.ToDouble(SpecCountAll[0].ItemArray[0]) * 100, 2) + "%", font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 2
                    };
                    table.AddCell(cell);

                }

                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);


                {
                    PdfPTable table2 = new PdfPTable(2);
                    cell = new PdfPCell(new Phrase("Оценочная шкала результатов анкетирования", font))
                    {
                        Colspan = 2,
                        HorizontalAlignment = 1,
                        Border = 0,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Степень удовлетворенности", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Процентный интервал удовлетворенности", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Неудовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("До 50%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная неудовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 50% до 65%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная удовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 65% до 80%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Полная удовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 80% до 100%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);

                    cell = new PdfPCell(table2)
                    {
                        Colspan = 4,
                        HorizontalAlignment = 1,
                        Border = 0
                    };
                    table.AddCell(cell);
                }

                cell = new PdfPCell(new Phrase("Общие выводы эксперта:\n1.Удовлетворенность требованиями к условиям реализации программы(вопросы 1 - 9) _____________________________________________________________________\n\n2.Удовлетворенность материально - техническим обеспечением программы(вопросы 10 - 11) ________________________________________________________________\n\n3.Общая удовлетворенность условиями организации образовательного процесса по программе(вопрос 12) __________________________________________________\n\n\nДата __________________\n\nРуководитель экспертной группы _______________ / Иванова Л.В. /\n\nПодпись представителя ОО,\nответственного за государственную аккредитацию\n программ по образовательной организации,\nдолжность                                                   _____________________ / Говорова О.Ю. /\n                                                                                    подпись               Ф.И.О.полностью\n\n(подпись представителя ОО заверяется)", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);

                //Добавляем таблицу в документ
                doc.Add(table);
                //Закрываем документ
                doc.Close();
            }
            // Протокол для базы 9 класса
            {
                //Объект документа пдф
                iTextSharp.text.Document doc = new iTextSharp.text.Document();

                //Создаем объект записи пдф-документа в файл
                NewFile9:
                int newf = 2;
                try
                {
                    PdfWriter.GetInstance(doc, new FileStream("Преподаватели - " + FileName.Text + " - 9.pdf", FileMode.Create));
                }
                catch
                {
                    try
                    {
                        PdfWriter.GetInstance(doc, new FileStream("Преподаватели - " + FileName.Text + "9 (" + newf + ").pdf", FileMode.Create));
                    }
                    catch
                    {
                        newf++;
                        goto NewFile9;
                    }
                }

                doc.SetMargins(80, 40, 40, 30);
                doc.SetPageSize(new iTextSharp.text.Rectangle(PageSize.A4));

                //Открываем документ
                doc.Open();

                //Определение шрифта необходимо для сохранения кириллического текста
                //Иначе мы не увидим кириллический текст
                //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
                BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\times.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                BaseFont baseFontBold = BaseFont.CreateFont("C:\\Windows\\Fonts\\timesbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName = new iTextSharp.text.Font(baseFontBold, 14, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName2 = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.NORMAL);

                //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                PdfPTable table = new PdfPTable(4);

                float[] widths = new float[] { 1.1f, 12f, 6.5f, 4.5f };
                table.SetWidths(widths);
                table.WidthPercentage = 100;
                //Добавим первую строку
                PdfPCell cell = new PdfPCell(new Phrase("Эксперт " + expert.Text, font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку
                cell = new PdfPCell(new Phrase("Наименование основной профессиональной образовательной программы", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим строку с названием программы
                cell = new PdfPCell(new Phrase(Convert.ToString(SelectSpec.SelectedValue), font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    BorderWidthBottom = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим пустую строку
                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим Название "протокол"
                cell = new PdfPCell(new Phrase("Протокол", fontName))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку названия
                cell = new PdfPCell(new Phrase("анкетирования педагогических работников\nреализующих программу", fontName2))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим аннотацию
                string usa = "SELECT statprepcount(" + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + ",9)";
                var UsAll = db.DbSelect(usa).Select();
                string us = "SELECT statprepcountspec(" + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + ",9)";
                var Us = db.DbSelect(us).Select();

                //Добавим аннотацию
                cell = new PdfPCell(new Phrase("В анкетировании приняли участие " + Us[0].ItemArray[0] + " преподавателей, что составило " + Math.Round(Convert.ToDouble(UsAll[0].ItemArray[0]) * 100, 2) + "% от количества научно-педагогических работников, реализующих программу.", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим заголовок таблицы
                cell = new PdfPCell(new Phrase("Результаты анкетирования", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                // Добавляем заголовок таблицы
                cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Вопросы педагогическим работникам аккредитуемой программы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);

                //Добавляем все остальные ячейки
                for (int j = 17; j < 29; j++)
                {
                    if (j == 28)
                    {
                        cell = new PdfPCell(new Phrase("Результаты анкетирования\n ", font))
                        {
                            Colspan = 4,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        // Добавляем заголовок таблицы
                        cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Вопросы педагогическим работникам аккредитуемой программы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                    }
                
                //var SpecCount = db.DbSelect("SELECT COUNT(ans.score) AS expr1 FROM spec CROSS JOIN anket_emp INNER JOIN ans ON anket_emp.ans = ans.id WHERE specemp.base = 9 AND spec.id = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + " AND ans.score >= 3 AND ans.idque = " + j).Select();
                var SpecCountAll = db.DbSelect("Select statprep(" + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + "," + j + ",9)").Select();
                    var Ques = db.DbSelect("SELECT ques.name FROM ques WHERE ques.id = " + j).Select();
                    var Ans = db.DbSelect("SELECT ans.text FROM ans INNER JOIN ques ON ans.idque = ques.id WHERE ques.id = " + j).Select();
                    cell = new PdfPCell(new Phrase(j - 16 + ".", font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(Ques[0].ItemArray[0].ToString(), font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    string an = "";
                    foreach (var Items in Ans)
                        an += "- " + Items[0] + "\n";
                    cell = new PdfPCell(new Phrase(an, font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(Math.Round(Convert.ToDouble(SpecCountAll[0].ItemArray[0]) * 100, 2) + "%", font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 2
                    };
                    table.AddCell(cell);

                }

                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);


                {
                    PdfPTable table2 = new PdfPTable(2);
                    cell = new PdfPCell(new Phrase("Оценочная шкала результатов анкетирования", font))
                    {
                        Colspan = 2,
                        HorizontalAlignment = 1,
                        Border = 0,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Степень удовлетворенности", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Процентный интервал удовлетворенности", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Неудовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("До 50%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная неудовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 50% до 65%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная удовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 65% до 80%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Полная удовлетворенность", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 80% до 100%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);

                    cell = new PdfPCell(table2)
                    {
                        Colspan = 4,
                        HorizontalAlignment = 1,
                        Border = 0
                    };
                    table.AddCell(cell);
                }

                cell = new PdfPCell(new Phrase("Общие выводы эксперта:\n1.Удовлетворенность требованиями к условиям реализации программы(вопросы 1 - 9) _____________________________________________________________________\n\n2.Удовлетворенность материально - техническим обеспечением программы(вопросы 10 - 11) ________________________________________________________________\n\n3.Общая удовлетворенность условиями организации образовательного процесса по программе(вопрос 12) __________________________________________________\n\n\nДата __________________\n\nРуководитель экспертной группы _______________ / Иванова Л.В. /\n\nПодпись представителя ОО,\nответственного за государственную аккредитацию\n программ по образовательной организации,\nдолжность                                                   _____________________ / Говорова О.Ю. /\n                                                                                    подпись               Ф.И.О.полностью\n\n(подпись представителя ОО заверяется)", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);

                //Добавляем таблицу в документ
                doc.Add(table);
                //Закрываем документ
                doc.Close();
            }
            return true;
        }

        private bool stud()
        {
            // Протокол для базы 9 класса
            string exi = "SELECT COUNT(groups.id) AS expr1 FROM groups WHERE groups.name LIKE '%9%' AND groups.spec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1);
            var exist = db.DbSelect(exi).Select();
            if (Convert.ToInt32(exist[0].ItemArray[0]) > 0)
            {
                //Объект документа пдф
                iTextSharp.text.Document doc = new iTextSharp.text.Document();

                //Создаем объект записи пдф-документа в файл
                NewFile9:
                int newf = 2;
                try
                {
                    PdfWriter.GetInstance(doc, new FileStream("Студенты - " + FileName.Text + " - 9.pdf", FileMode.Create));
                }
                catch
                {
                    try
                    {
                        PdfWriter.GetInstance(doc, new FileStream("Студенты - " + FileName.Text + " - 9 (" + newf + ").pdf", FileMode.Create));
                    }
                    catch
                    {
                        newf++;
                        goto NewFile9;
                    }
                }

                doc.SetMargins(80, 40, 40, 30);
                doc.SetPageSize(new iTextSharp.text.Rectangle(PageSize.A4));
                //Открываем документ
                doc.Open();

                //Определение шрифта необходимо для сохранения кириллического текста
                //Иначе мы не увидим кириллический текст
                //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
                BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\times.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                BaseFont baseFontBold = BaseFont.CreateFont("C:\\Windows\\Fonts\\timesbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName = new iTextSharp.text.Font(baseFontBold, 14, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName2 = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.NORMAL);

                //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                PdfPTable table = new PdfPTable(4);
                float[] widths = new float[] { 1.1f, 12f, 6.5f, 4.5f };
                table.SetWidths(widths);
                table.WidthPercentage = 100;

                //Добавим первую строку
                PdfPCell cell = new PdfPCell(new Phrase("Эксперт " + expert.Text, font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку
                cell = new PdfPCell(new Phrase("Наименование основной профессиональной образовательной программы", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим строку с названием программы
                cell = new PdfPCell(new Phrase(Convert.ToString(SelectSpec.SelectedValue), font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    BorderWidthBottom = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим пустую строку
                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим Название "протокол"
                cell = new PdfPCell(new Phrase("Протокол", fontName))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку названия
                cell = new PdfPCell(new Phrase("анкетирования обучающихся", fontName2))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                Renew:
                string usa = "SELECT stat.count FROM stat WHERE stat.type = 1 AND stat.base = 9 AND stat.idspec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1);
                var UsAll = db.DbSelect(usa).Select();
                string us = "SELECT COUNT(user.id) AS expr1 FROM user INNER JOIN groups ON user.`group` = groups.id WHERE groups.name LIKE '%9%' AND groups.spec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1);
                var Us = db.DbSelect(us).Select();

                //Добавим аннотацию
                try
                {
                    cell = new PdfPCell(new Phrase("В анкетировании приняли " + Us[0].ItemArray[0] + " обучающихся, что составило " + Math.Round(Convert.ToDouble(Us[0].ItemArray[0]) / Convert.ToDouble(UsAll[0].ItemArray[0]) * 100,2) + "% от количества обучающихся по программе.\n", font))
                    {
                        Colspan = 4,
                        HorizontalAlignment = 0,
                        Border = 0,
                        VerticalAlignment = 1
                    };
                }
                catch
                {
                    goto Renew;
                }
                table.AddCell(cell);
                //Добавим заголовок таблицы
                cell = new PdfPCell(new Phrase("Результаты анкетирования\n ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                // Добавляем заголовок таблицы
                cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Вопросы обучающимся аккредитуемой программы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавляем все остальные ячейки
                for (int j = 1; j < 17; j++)
                {
                    var SpecCount = db.DbSelect("SELECT COUNT(ans.score) AS expr1 FROM anket INNER JOIN ans ON anket.ans = ans.id INNER JOIN user ON anket.iduser = user.id INNER JOIN groups ON user.`group` = groups.id WHERE groups.spec = " + (SelectSpec.SelectedIndex + 1) + " AND anket.idque = " + j + " AND ans.score >= 3 AND groups.name LIKE '%9%'").Select();
                    var SpecCountAll = db.DbSelect("SELECT COUNT(ans.score) AS expr1 FROM anket INNER JOIN ans ON anket.ans = ans.id INNER JOIN user ON anket.iduser = user.id INNER JOIN groups ON user.`group` = groups.id WHERE groups.spec = " + (SelectSpec.SelectedIndex + 1) + " AND groups.name LIKE '%9%' AND anket.idque = " + j).Select();
                    var Ques = db.DbSelect("SELECT ques.name FROM ques WHERE ques.id = " + j).Select();
                    var Ans = db.DbSelect("SELECT ans.text FROM ans INNER JOIN ques ON ans.idque = ques.id WHERE ques.id = " + j).Select();
                    cell = new PdfPCell(new Phrase(j + ".", font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(Ques[0].ItemArray[0].ToString(), font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    string an = "";
                    foreach (var Items in Ans)
                        an += "- " + Items[0] + "\n";
                    cell = new PdfPCell(new Phrase(an, font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    string proc = "";
                    if (Convert.ToInt32(SpecCountAll[0].ItemArray[0]) != 0)
                        proc = Convert.ToString(Math.Round(Convert.ToDouble(SpecCount[0].ItemArray[0]) / Convert.ToDouble(SpecCountAll[0].ItemArray[0]) * 100,2)) + "%";
                    else
                        proc = "100%";
                    cell = new PdfPCell(new Phrase(proc, font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 2
                    };
                    table.AddCell(cell);

                    if (j == 9)
                    {
                        cell = new PdfPCell(new Phrase("Результаты анкетирования\n ", font))
                        {
                            Colspan = 4,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        // Добавляем заголовок таблицы
                        cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Вопросы обучающимся аккредитуемой программы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                    }
                }

                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);


                {
                    PdfPTable table2 = new PdfPTable(2);
                    cell = new PdfPCell(new Phrase("Оценочная шкала результатов анкетирования\n ", font))
                    {
                        Colspan = 4,
                        HorizontalAlignment = 1,
                        Border = 0,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Степень удовлетворенности\n ", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Процентный интервал удовлетворенности\n ", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Неудовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("До 50%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная неудовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 50% до 65%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная удовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 65% до 80%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Полная удовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 80% до 100%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);

                    cell = new PdfPCell(table2)
                    {
                        Colspan = 4,
                        HorizontalAlignment = 1,
                        Border = 0
                    };
                    table.AddCell(cell);
                }

                cell = new PdfPCell(new Phrase("Общие выводы эксперта.\n\n1.Удовлетворенность структурой программы(вопросы 1, 2, 3, 4) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("2.Удовлетворенность требованиями к условиям реализации программы(вопросы 5, 6, 8, 9, 10, 11, 12, 13) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("3.Удовлетворенность учебно - методическим обеспечением программы(вопросы 7) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("4.Удовлетворенность материально - техническим обеспечением программы(вопросы 14, 15) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("5.Общая удовлетворенность качеством предоставления образовательных услуг по программе(вопросы 16) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("Дата __________________\n\nПодпись эксперта ____________________ / ________________ /.\n\nПодпись представителя ОО,\nответственного за государственную аккредитацию\n программ по образовательной организации,\nдолжность                                          _____________________ / _______________ /\n                                                                           подпись                Ф.И.О.полностью\n (подпись представителя ОО заверяется)", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);

                //Добавляем таблицу в документ
                doc.Add(table);
                //Закрываем документ
                doc.Close();
            }
            // Протокол для базы 11 класса 

            exi = "SELECT COUNT(groups.id) AS expr1 FROM groups WHERE groups.name LIKE '%11%' AND groups.spec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1);
            exist = db.DbSelect(exi).Select();
            if (Convert.ToInt32(exist[0].ItemArray[0])>0){
                //Объект документа пдф
                iTextSharp.text.Document doc = new iTextSharp.text.Document();

                //Создаем объект записи пдф-документа в файл
                NewFile11:
                int newf = 2;
                try
                {
                    PdfWriter.GetInstance(doc, new FileStream("Студенты - " + FileName.Text + " - 11.pdf", FileMode.Create));
                }
                catch
                {
                    try
                    {
                        PdfWriter.GetInstance(doc, new FileStream("Студенты - " + FileName.Text + " - 11 (" + newf + ").pdf", FileMode.Create));
                    }
                    catch
                    {
                        newf++;
                        goto NewFile11;
                    }
                }

                doc.SetMargins(80, 40, 40, 30);
                doc.SetPageSize(new iTextSharp.text.Rectangle(PageSize.A4));
                //Открываем документ
                doc.Open();

                //Определение шрифта необходимо для сохранения кириллического текста
                //Иначе мы не увидим кириллический текст
                //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
                BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\times.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                BaseFont baseFontBold = BaseFont.CreateFont("C:\\Windows\\Fonts\\timesbd.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName = new iTextSharp.text.Font(baseFontBold, 14, iTextSharp.text.Font.NORMAL);
                iTextSharp.text.Font fontName2 = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.NORMAL);

                //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                PdfPTable table = new PdfPTable(4);
                float[] widths = new float[] { 1.1f, 12f, 6.5f, 4.5f };
                table.SetWidths(widths);
                table.WidthPercentage = 100;

                //Добавим первую строку
                PdfPCell cell = new PdfPCell(new Phrase("Эксперт " + expert.Text, font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку
                cell = new PdfPCell(new Phrase("Наименование основной профессиональной образовательной программы", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим строку с названием программы
                cell = new PdfPCell(new Phrase(Convert.ToString(SelectSpec.SelectedValue), font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    BorderWidthBottom = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим пустую строку
                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим Название "протокол"
                cell = new PdfPCell(new Phrase("Протокол", fontName))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавим вторую строку названия
                cell = new PdfPCell(new Phrase("анкетирования обучающихся", fontName2))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                Renew:
                string usa = "SELECT stat.count FROM stat WHERE stat.type = 1 AND stat.base = 11 AND stat.idspec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1);
                var UsAll = db.DbSelect(usa).Select();
                string us = "SELECT COUNT(user.id) AS expr1 FROM user INNER JOIN groups ON user.`group` = groups.id WHERE groups.name LIKE '%11%' AND groups.spec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1);
                var Us = db.DbSelect(us).Select();

                //Добавим аннотацию
                try
                {
                    cell = new PdfPCell(new Phrase("В анкетировании приняли " + Us[0].ItemArray[0] + " обучающихся, что составило " + Math.Round(Convert.ToDouble(Us[0].ItemArray[0]) / Convert.ToDouble(UsAll[0].ItemArray[0]) * 100,2) + "% от количества обучающихся по программе.\n", font))
                    {
                        Colspan = 4,
                        HorizontalAlignment = 0,
                        Border = 0,
                        VerticalAlignment = 1
                    };
                }
                catch
                {
                    goto Renew;
                }
                table.AddCell(cell);
                //Добавим заголовок таблицы
                cell = new PdfPCell(new Phrase("Результаты анкетирования\n ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                // Добавляем заголовок таблицы
                cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Вопросы обучающимся аккредитуемой программы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                {
                    //Фоновый цвет
                    BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                    HorizontalAlignment = 1,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);
                //Добавляем все остальные ячейки
                for (int j = 1; j < 17; j++)
                {
                    var SpecCount = db.DbSelect("SELECT COUNT(ans.score) AS expr1 FROM anket INNER JOIN ans ON anket.ans = ans.id INNER JOIN user ON anket.iduser = user.id INNER JOIN groups ON user.`group` = groups.id WHERE groups.spec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + " AND anket.idque = " + j + " AND ans.score >= 3 AND groups.name LIKE '%11%'").Select();
                    var SpecCountAll = db.DbSelect("SELECT COUNT(ans.score) AS expr1 FROM anket INNER JOIN ans ON anket.ans = ans.id INNER JOIN user ON anket.iduser = user.id INNER JOIN groups ON user.`group` = groups.id WHERE groups.spec = " + (Convert.ToInt32(SelectSpec.SelectedIndex) + 1) + " AND groups.name LIKE '%11%' AND anket.idque = " + j).Select();
                    var Ques = db.DbSelect("SELECT ques.name FROM ques WHERE ques.id = " + j).Select();
                    var Ans = db.DbSelect("SELECT ans.text FROM ans INNER JOIN ques ON ans.idque = ques.id WHERE ques.id = " + j).Select();
                    cell = new PdfPCell(new Phrase(j + ".", font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    cell = new PdfPCell(new Phrase(Ques[0].ItemArray[0].ToString(), font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    string an = "";
                    foreach (var Items in Ans)
                        an += "- " + Items[0] + "\n";
                    cell = new PdfPCell(new Phrase(an, font))
                    {
                        HorizontalAlignment = 0,
                        VerticalAlignment = 1
                    };
                    table.AddCell(cell);
                    string proc = "";
                    if (Convert.ToInt32(SpecCountAll[0].ItemArray[0]) != 0)
                        proc = Convert.ToString(Math.Round(Convert.ToDouble(SpecCount[0].ItemArray[0]) / Convert.ToDouble(SpecCountAll[0].ItemArray[0]) * 100,2)) + "%";
                    else
                        proc = "100%";
                    cell = new PdfPCell(new Phrase(proc, font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 2
                    };
                    table.AddCell(cell);

                    if (j == 9)
                    {
                        cell = new PdfPCell(new Phrase("Результаты анкетирования\n ", font))
                        {
                            Colspan = 4,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        // Добавляем заголовок таблицы
                        cell = new PdfPCell(new Phrase(new Phrase("№\nп\\п", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Вопросы обучающимся аккредитуемой программы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Ответы", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                        cell = new PdfPCell(new Phrase(new Phrase("Результаты анкетирования, %", font)))
                        {
                            //Фоновый цвет
                            BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                            HorizontalAlignment = 1,
                            VerticalAlignment = 1
                        };
                        table.AddCell(cell);
                    }
                }

                cell = new PdfPCell(new Phrase(" ", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 1,
                    Border = 0,
                    VerticalAlignment = 1
                };
                table.AddCell(cell);


                {
                    PdfPTable table2 = new PdfPTable(2);
                    cell = new PdfPCell(new Phrase("Оценочная шкала результатов анкетирования\n ", font))
                    {
                        Colspan = 4,
                        HorizontalAlignment = 1,
                        Border = 0,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Степень удовлетворенности\n ", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase(new Phrase("Процентный интервал удовлетворенности\n ", font)))
                    {
                        //Фоновый цвет
                        BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY,
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Неудовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("До 50%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная неудовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 50% до 65%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Частичная удовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 65% до 80%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("Полная удовлетворенность\n ", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);
                    cell = new PdfPCell(new Phrase("От 80% до 100%", font))
                    {
                        HorizontalAlignment = 1,
                        VerticalAlignment = 1
                    };
                    table2.AddCell(cell);

                    cell = new PdfPCell(table2)
                    {
                        Colspan = 4,
                        HorizontalAlignment = 1,
                        Border = 0
                    };
                    table.AddCell(cell);
                }

                cell = new PdfPCell(new Phrase("Общие выводы эксперта.\n\n1.Удовлетворенность структурой программы(вопросы 1, 2, 3, 4) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("2.Удовлетворенность требованиями к условиям реализации программы(вопросы 5, 6, 8, 9, 10, 11, 12, 13) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("3.Удовлетворенность учебно - методическим обеспечением программы(вопросы 7) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("4.Удовлетворенность материально - техническим обеспечением программы(вопросы 14, 15) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("5.Общая удовлетворенность качеством предоставления образовательных услуг по программе(вопросы 16) _____________________________________________________________________\n\n", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);
                cell = new PdfPCell(new Phrase("Дата __________________\n\nПодпись эксперта ____________________ / ________________ /.\n\nПодпись представителя ОО,\nответственного за государственную аккредитацию\n программ по образовательной организации,\nдолжность                                          _____________________ / _______________ /\n                                                                           подпись                Ф.И.О.полностью\n (подпись представителя ОО заверяется)", font))
                {
                    Colspan = 4,
                    HorizontalAlignment = 0,
                    Border = 0
                };
                table.AddCell(cell);

                //Добавляем таблицу в документ
                doc.Add(table);
                //Закрываем документ
                doc.Close();
            }
            return true;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            SelectSpec.SelectedIndex = 0;
            cycle:
                FileName.Text = Convert.ToString(SelectSpec.SelectedValue);
                stud();
                prep();
            if (SelectSpec.SelectedIndex++<15) goto cycle;
            MessageBox.Show("Протоколы составлены");
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.Form ank = new anket();
            ank.Show();
        }
    }
}
