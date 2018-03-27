using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;

namespace anketEmp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private Db db;
        private user user;
        private Ans[] ans = new Ans[100];
        private int[] GetV = new int[100];

        public MainWindow()
        {
            InitializeComponent();
            db = new Db();
            if (db.Connect())
            {
                MessageBox.Show("Нет доступных серверов");
                this.Close();
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (userName.Text.Length > 0)
                {
                    bool q = false;
                    for (int i = 0; i < 100; i++)
                        if (GetV[i] > 0) q = true;
                    if (!q) goto Err;
                    user = new user(userName.Text, GetV);
                    auth.Visibility = Visibility.Collapsed;
                    anket.Visibility = Visibility.Visible;
                    info.Text = user.name;
                    Ques();
                    return;
                }
                else goto Err;
                Err:
                MessageBox.Show("Данные не верны");
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
                this.Close();
            }
        }

        private void Ques()
        {
            try
            {
                var ques = db.DbSelect("SELECT ques.* FROM ques WHERE ques.tea = 1").Select();
                int i = 0;
                int heightGrid = 0;
                Grid qurid = new Grid();
                foreach (var item in ques)
                {
                    i++;
                    Grid grid = new Grid();
                    grid.Name = "ques" + i;
                    var ans = db.DbSelect("SELECT ans.* FROM ans WHERE ans.idque = " + item[0]).Select();
                    int k = 0;
                    Grid gridLeft = new Grid();
                    Grid gridRight = new Grid();
                    TextBlock Text = new TextBlock();
                    RadioButton radio = new RadioButton();
                    gridLeft.Width = 500;
                    gridLeft.HorizontalAlignment = HorizontalAlignment.Left;
                    gridRight.HorizontalAlignment = HorizontalAlignment.Left;
                    gridRight.Margin = new Thickness(500, 0, 0, 0);
                    Text.Text = item[1] + "";
                    Text.TextWrapping = TextWrapping.Wrap;
                    StackPanel stack = new StackPanel();
                    if ((int)item[2] == 1)
                        foreach (var answer in ans)
                        {
                            radio = new RadioButton();
                            radio.GroupName = Convert.ToString(item[0]);
                            radio.Content = answer[3];
                            radio.Tag = answer[4];
                            radio.Uid = Convert.ToString(answer[0]);
                            radio.MinHeight = 20;
                            //radio.Margin = new Thickness(0, k * 30, 0, 0);
                            radio.Checked += radioButton_CheckedChanged;
                            stack.Children.Add(radio);
                            k++;

                        }
                    else
                        foreach (var answer in ans)
                        {
                            CheckBox check = new CheckBox();
                            check.Content = answer[3];
                            check.Tag = item[0];
                            check.Uid = Convert.ToString(answer[0]);
                            check.MinHeight = 20;
                            //radio.Margin = new Thickness(0, k * 30, 0, 0);
                            check.Checked += check_CheckedChanged;
                            check.Unchecked += CheckBox_Unchecked_1;
                            stack.Children.Add(check);
                            k++;

                        }
                    stack.Height = k * 20;
                    stack.VerticalAlignment = VerticalAlignment.Top;
                    gridLeft.Children.Add(Text);
                    gridRight.Children.Add(stack);
                    grid.Children.Add(gridLeft);
                    grid.Children.Add(gridRight);
                    grid.Margin = new Thickness(0, heightGrid, 0, 0);
                    heightGrid += k * 20 + 50;
                    qurid.Children.Add(grid);
                }
                qurid.Height = heightGrid;
                Scroll.Content = qurid;
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
                this.Close();
            }
        }

        private void check_CheckedChanged(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckBox check = (CheckBox)sender;
                if ((bool)check.IsChecked)
                {
                    if (!(ans[Convert.ToInt32(check.Tag)] == null))
                    {
                        int i = 0;
                        while (ans[Convert.ToInt32(check.Tag)].ansarr[i] != 0)
                        {
                            if ((ans[Convert.ToInt32(check.Tag)].ansarr[i] != 0) && (ans[Convert.ToInt32(check.Tag)].ansarr[i + 1] == 0))
                            {
                                ans[Convert.ToInt32(check.Tag)].ansarr[i + 1] = Convert.ToInt32(check.Uid);
                                break;
                            }
                            i++;
                        }
                    }
                    else
                    {
                        int[] arr = new int[10];
                        arr[0] = Convert.ToInt32(check.Uid);
                        ans[Convert.ToInt32(check.Tag)] = new Ans(check.Tag, arr);
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void CheckBox_Unchecked_1(object sender, RoutedEventArgs e)
        {

            CheckBox check = (CheckBox)sender;
            int i = 0;
            while (ans[Convert.ToInt32(check.Tag)].ansarr[i] != Convert.ToInt32(check.Uid)) i++;
            ans[Convert.ToInt32(check.Tag)].ansarr[i] = 0;
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                RadioButton radioButton = (RadioButton)sender;
                if ((bool)radioButton.IsChecked)
                {
                    ans[Convert.ToInt32(radioButton.GroupName)] = new Ans(radioButton.GroupName, radioButton.Tag, radioButton.Uid);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                db.DbInsert("INSERT INTO emp (id, name, stat) VALUES(DEFAULT, '" + user.name + "', DEFAULT)");
                var us = db.DbSelect("SELECT LAST_INSERT_ID() FROM `emp` LIMIT 1;").Select();
                int i = 17;
                for (int q = 0; q < 100; q++)
                    if (GetV[q] > 0)
                    {
                        if (0 != q % 2)
                            db.DbInsert("INSERT INTO specemp (id, idspec, idemp, base) VALUES(DEFAULT, " + GetV[q] + ", " + us[0].ItemArray[0] + ", " + 9 + ")");
                        else
                            db.DbInsert("INSERT INTO specemp (id, idspec, idemp, base) VALUES(DEFAULT, " + GetV[q] + ", " + us[0].ItemArray[0] + ", " + 11 + ")");
                    }
                while ((i < ans.Length) && (ans[i] != null))
                {
                    if (ans[i].ansarr == null)
                    {
                        string sql = "INSERT INTO anket_emp (id, idque, iduser, ans) VALUES(DEFAULT, ";
                        sql += ans[i].queid + ", ";
                        sql += us[0].ItemArray[0] + ", ";
                        sql += ans[i].ansid + ")";
                        db.DbInsert(sql);
                    }
                    else
                    {
                        int q = 0;
                        while (ans[i].ansarr[q] != 0)
                        {
                            string sql = "INSERT INTO anket_emp (id, idque, iduser, ans) VALUES(DEFAULT, ";
                            sql += ans[i].queid + ", ";
                            sql += us[0].ItemArray[0] + ", ";
                            sql += ans[i].ansarr[q] + ")";
                            db.DbInsert(sql);
                            q++;
                        }
                    }
                    i++;
                }
                
                MessageBox.Show("Результат записан");
                this.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void spec_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                var ComboGroups = db.DbSelect("SELECT spec.id, spec.`key`, spec.name FROM spec ORDER BY spec.id").Select();
                int i = 0;
                int t = 1;
                if (ComboGroups.Length > 0)
                    foreach (var Items in ComboGroups)
                    {
                        CheckBox checkBox = new CheckBox { Tag = t++, Content = Items.ItemArray[1] + " (9кл) - " + Items.ItemArray[2], MinHeight = 20, IsChecked = false, Name = "c" + i++, Uid = Items.ItemArray[0].ToString() };
                        // установка обработчика
                        checkBox.Checked += CheckBox_Checked;
                        // добавление в StackPanel
                        spec.Children.Add(checkBox);
                        spec.Height += 20;
                        checkBox = new CheckBox { Tag = t++, Content = Items.ItemArray[1] + " (11кл) - " + Items.ItemArray[2], MinHeight = 20, IsChecked = false, Name = "c" + i++, Uid = Items.ItemArray[0].ToString() };
                        // установка обработчика
                        checkBox.Checked += CheckBox_Checked;
                        // добавление в StackPanel
                        spec.Children.Add(checkBox);
                        spec.Height += 20;
                    }
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckBox chBox = (CheckBox)sender;
                GetV[Convert.ToInt32(chBox.Tag)] = 0;
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                CheckBox chBox = (CheckBox)sender;
                GetV[Convert.ToInt32(chBox.Tag)] = Convert.ToInt32(chBox.Uid);
            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }

        }
    }
}
