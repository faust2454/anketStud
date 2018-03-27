using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace anketResult
{
    public partial class anket : Form
    {
        private Db db;
        public anket()
        {
            db = new Db();
            if (db.Connect())
            {
                MessageBox.Show("Нет доступных серверов");
                this.Close();
            }
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Enabled) comboBox2.Items.Clear();
            if (comboBox4.Enabled) comboBox4.Items.Clear();
            try
            {
                comboBox2.Enabled = true;
                var ComboGroups = db.DbSelect("SELECT spec.name FROM spec").Select();
                if (ComboGroups.Length > 0)
                    foreach (var Items in ComboGroups)
                        comboBox2.Items.Add(Items.ItemArray[0]);

            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Enabled = true;
            if (comboBox4.Enabled) comboBox4.Items.Clear();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string sql = "";
                if (comboBox1.SelectedIndex == 0)
                    sql = "SELECT emp.id, emp.name FROM specemp INNER JOIN spec ON specemp.idspec = spec.id INNER JOIN emp ON specemp.idemp = emp.id WHERE spec.name = '" + comboBox2.SelectedItem + "' AND specemp.base LIKE '%" + comboBox3.SelectedItem + "%'";
                else
                    sql = "SELECT user.id, user.fname FROM groups INNER JOIN spec ON groups.spec = spec.id INNER JOIN user ON user.`group` = groups.id WHERE spec.name = '" + comboBox2.SelectedItem + "' AND groups.name LIKE '%" + comboBox3.SelectedItem + "%'";
                comboBox4.Enabled = true;
                var ComboGroups = db.DbSelect(sql).Select();
                if (ComboGroups.Length > 0)
                    foreach (var Items in ComboGroups)
                        comboBox4.Items.Add(Items.ItemArray[1]);

            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.Clear();
                string sql = "";
                if (comboBox1.SelectedIndex == 0)
                    sql = "SELECT anket_emp.idque, anket_emp.ans FROM anket_emp INNER JOIN specemp ON anket_emp.iduser = specemp.idemp INNER JOIN emp ON specemp.idemp = emp.id WHERE emp.name = '" + comboBox4.SelectedItem + "' GROUP BY anket_emp.idque";
                else
                    sql = "SELECT anket.idque, anket.ans FROM groups INNER JOIN spec ON groups.spec = spec.id INNER JOIN user ON user.`group` = groups.id INNER JOIN anket ON anket.iduser = user.id WHERE spec.name = '" + comboBox2.SelectedItem + "' AND groups.name LIKE '%" + comboBox3.SelectedItem + "%' AND user.fname = '" + comboBox4.SelectedItem + "'";
                var ComboGroups = db.DbSelect(sql).Select();
                int i = 1;
                if (ComboGroups.Length > 0)
                    foreach (var Items in ComboGroups)
                    {
                        var ques = db.DbSelect("SELECT ques.name FROM ques WHERE ques.id = " + Items.ItemArray[0]).Select();
                        var ans = db.DbSelect("SELECT ans.text FROM ans WHERE ans.id = " + Items.ItemArray[1]).Select();
                        dataGridView1.Rows.Add(i++, ques[0].ItemArray[0], ans[0].ItemArray[0]);
                    }

            }
            catch (Exception)
            {
                MessageBox.Show("Произошла ошибка");
            }
        }
    }
}
