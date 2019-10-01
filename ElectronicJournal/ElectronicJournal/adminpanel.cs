using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;
using System.IO;
using xNet;
using System.Threading;


namespace ElectronicJournal
{
    public partial class adminpanel : MetroFramework.Forms.MetroForm
    {
        public adminpanel()
        {
            InitializeComponent();
        }

        connect ConnectDB = new connect();
        string data = "";

        private void adminpanel_Load(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT gruppa.gruppa AS [Группа], kurs.kurs AS[Курс], specialnost.specialnost AS[Специальность] FROM gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds, "gruppa");
            metroGrid1.DataSource = ds.Tables["gruppa"];

            DataTable dt2 = new DataTable();
            string q2 = "SELECT DISTINCT kurs.kurs FROM kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs";
            SqlCeDataAdapter dtAdapter4 = new SqlCeDataAdapter(q2, ConnectDB.ConnectOnDb());
            dtAdapter4.Fill(dt2);
            comboBox2.DataSource = dt2;
            comboBox2.DisplayMember = "kurs";
            comboBox2.ValueMember = "kurs";

            DataTable dt3 = new DataTable();
            string q3 = "SELECT DISTINCT specialnost.specialnost FROM gruppa INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost";
            SqlCeDataAdapter dtAdapter5 = new SqlCeDataAdapter(q3, ConnectDB.ConnectOnDb());
            dtAdapter5.Fill(dt3);
            comboBox1.DataSource = dt3;
            comboBox1.DisplayMember = "specialnost";
            comboBox1.ValueMember = "specialnost";

            DataTable dt = new DataTable();
            string q = "SELECT DISTINCT gruppa FROM gruppa";
            SqlCeDataAdapter dataAdapter2 = new SqlCeDataAdapter(q, ConnectDB.ConnectOnDb());
            dataAdapter2.Fill(dt);
            comboBox3.DataSource = dt;
            comboBox3.DisplayMember = "gruppa";
            comboBox3.ValueMember = "gruppa";

            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT specialnost.id_specialnost FROM gruppa INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost WHERE(gruppa.gruppa = '" + comboBox3.GetItemText(comboBox3.SelectedItem) + "')";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int sp = (int)thisCommand.ExecuteScalar();

            DataSet ds4 = new DataSet();
            SqlCeDataAdapter dtAdapter6 = new SqlCeDataAdapter("SELECT student.id_zachetki AS [Номер зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчество, gruppa.gruppa AS Группа, student.login AS Логин, student.pass AS Пароль FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE(gruppa.gruppa = '" + comboBox3.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter6.Fill(ds4, "student");
            metroGrid2.DataSource = ds4.Tables["student"];

            DataSet ds2 = new DataSet();
            SqlCeDataAdapter dtAdapter3 = new SqlCeDataAdapter("SELECT Sname AS Фамилия, Fname AS Имя, Lname AS Отчество, Login AS Логин, Password AS Пароль FROM teacher", ConnectDB.ConnectOnDb());
            dtAdapter3.Fill(ds2, "teacher");
            metroGrid3.DataSource = ds2.Tables["teacher"];

            DataSet ds3 = new DataSet();
            SqlCeDataAdapter dtAdapter33 = new SqlCeDataAdapter("SELECT predmet.predmet AS Предмет, specialnost.specialnost AS Специальность, kurs.kurs AS Курс FROM predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost", ConnectDB.ConnectOnDb());
            dtAdapter33.Fill(ds3, "predmet");
            metroGrid4.DataSource = ds3.Tables["predmet"];

            DataTable dt22 = new DataTable();
            string q22 = "SELECT DISTINCT kurs.kurs FROM kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs";
            SqlCeDataAdapter dtAdapter44 = new SqlCeDataAdapter(q22, ConnectDB.ConnectOnDb());
            dtAdapter44.Fill(dt22);
            comboBox4.DataSource = dt22;
            comboBox4.DisplayMember = "kurs";
            comboBox4.ValueMember = "kurs";

            DataTable dt33 = new DataTable();
            string q33 = "SELECT DISTINCT specialnost.specialnost FROM gruppa INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost";
            SqlCeDataAdapter dtAdapter55 = new SqlCeDataAdapter(q33, ConnectDB.ConnectOnDb());
            dtAdapter55.Fill(dt33);
            comboBox5.DataSource = dt33;
            comboBox5.DisplayMember = "specialnost";
            comboBox5.ValueMember = "specialnost";

           
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            ConnectDB.ConnectOnDb();
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT id_specialnost FROM specialnost WHERE specialnost = '" + comboBox1.Text + "'";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int x = (int)thisCommand.ExecuteScalar();
            SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand2.CommandText = "SELECT id_kurs FROM kurs WHERE kurs = '" + comboBox2.Text + "'";
            thisCommand2.Connection = ConnectDB.ConnectOnDb();
            int x2 = (int)thisCommand2.ExecuteScalar();
            SqlCeCommand cmd = new SqlCeCommand("Insert into gruppa (gruppa, id_specialnost, id_kurs)  Values ('" + metroTextBox1.Text + "', '" + x + "','" + x2 + "')", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            metroTextBox1.Clear();
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT  gruppa.gruppa AS [Группа], kurs.kurs AS[Курс], specialnost.specialnost AS[Специальность] FROM gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds, "gruppa");
            metroGrid1.DataSource = ds.Tables["gruppa"];
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            string s = metroGrid1[0, metroGrid1.CurrentRow.Index].Value.ToString();
            ConnectDB.ConnectOnDb();
            SqlCeCommand cmd = new SqlCeCommand("DELETE FROM gruppa WHERE gruppa = '" + s + "'", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT  gruppa.gruppa AS [Группа], kurs.kurs AS[Курс], specialnost.specialnost AS[Специальность] FROM gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds, "gruppa");
            metroGrid1.DataSource = ds.Tables["gruppa"];
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataSet ds4 = new DataSet();
            SqlCeDataAdapter dtAdapter6 = new SqlCeDataAdapter("SELECT student.id_zachetki AS [Номер зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчество, gruppa.gruppa AS Группа, student.login AS Логин, student.pass AS Пароль FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE(gruppa.gruppa = '" + comboBox3.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter6.Fill(ds4, "student");
            metroGrid2.DataSource = ds4.Tables["student"];
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            ConnectDB.ConnectOnDb();
            SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand2.CommandText = "SELECT id_gruppa FROM gruppa WHERE gruppa = '" + comboBox3.Text + "'";
            thisCommand2.Connection = ConnectDB.ConnectOnDb();
            int x = (int)thisCommand2.ExecuteScalar();
            SqlCeCommand cmd = new SqlCeCommand("Insert into student Values ('" + metroTextBox2.Text + "', '" + metroTextBox3.Text + "','" + metroTextBox4.Text + "','" + metroTextBox5.Text + "','" + x + "','" + metroTextBox12.Text + "','" + metroTextBox13.Text + "')", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            AuthServer();
            metroTextBox2.Clear();
            metroTextBox3.Clear();
            metroTextBox4.Clear();
            metroTextBox5.Clear();
            metroTextBox12.Clear();
            metroTextBox13.Clear();

            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT student.id_zachetki AS [Номер зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчество, gruppa.gruppa AS Группа, student.login AS Логин, student.pass AS Пароль FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE(gruppa.gruppa = '" + comboBox3.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds, "student");
            metroGrid2.DataSource = ds.Tables["student"];
        }

        private void AuthServer()
        {
            HttpRequest p = new HttpRequest();
            p.UserAgent = Http.ChromeUserAgent();
            RequestParams pd = new RequestParams();

            pd["id_zachetki"] = metroTextBox2.Text;
            pd["Sname"] = metroTextBox3.Text;
            pd["Fname"] = metroTextBox4.Text;
            pd["Lname"] = metroTextBox5.Text;
            pd["gruppa"] = comboBox3.Text;
            pd["login"] = metroTextBox12.Text;
            pd["pass"] = metroTextBox13.Text;
            data = p.Post("http://q961075i.beget.tech/insert.php", pd).ToString();
             
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            string s = metroGrid2[0, metroGrid2.CurrentRow.Index].Value.ToString();
            ConnectDB.ConnectOnDb();
            SqlCeCommand cmd = new SqlCeCommand("DELETE FROM student WHERE id_zachetki = '" + s + "'", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT student.id_zachetki AS [Номер зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчество, gruppa.gruppa AS Группа, student.login AS Логин, student.pass AS Пароль FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE(gruppa.gruppa = '" + comboBox3.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds, "student");
            metroGrid2.DataSource = ds.Tables["student"];
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            string idZachetki = metroGrid2[0, metroGrid2.CurrentRow.Index].Value.ToString();
            label1.Text = idZachetki;
            Form UpS = new UpStudents();
            UpS.Owner = this;
            UpS.Show();
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            ConnectDB.ConnectOnDb();
            if (metroTextBox7.Text == "" || metroTextBox6.Text == "" || metroTextBox8.Text == "" || metroTextBox9.Text == "" || metroTextBox10.Text == "")
            {
                MessageBox.Show("Пустые поля!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                SqlCeCommand cmd = new SqlCeCommand("Insert into teacher (Sname, Fname, Lname, Login, Password, Prava) Values ('" + metroTextBox7.Text + "', '" + metroTextBox6.Text + "','" + metroTextBox8.Text + "','" + metroTextBox9.Text + "','" + metroTextBox10.Text + "','" + 0 + "')", ConnectDB.ConnectOnDb());
                cmd.ExecuteNonQuery();
                ConnectDB.ConnectClose();
                metroTextBox6.Clear();
                metroTextBox7.Clear();
                metroTextBox8.Clear();
                metroTextBox9.Clear();
                metroTextBox10.Clear();

                DataSet ds2 = new DataSet();
                SqlCeDataAdapter dtAdapter3 = new SqlCeDataAdapter("SELECT Sname AS Фамилия, Fname AS Имя, Lname AS Отчество, Login AS Логин, Password AS Пароль FROM teacher", ConnectDB.ConnectOnDb());
                dtAdapter3.Fill(ds2, "teacher");
                metroGrid3.DataSource = ds2.Tables["teacher"];
            }
        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            string s = metroGrid3[3, metroGrid3.CurrentRow.Index].Value.ToString();
            ConnectDB.ConnectOnDb();
            SqlCeCommand cmd = new SqlCeCommand("DELETE FROM teacher WHERE Login = '" + s + "'", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();

            DataSet ds2 = new DataSet();
            SqlCeDataAdapter dtAdapter3 = new SqlCeDataAdapter("SELECT Sname AS Фамилия, Fname AS Имя, Lname AS Отчество, Login AS Логин, Password AS Пароль FROM teacher", ConnectDB.ConnectOnDb());
            dtAdapter3.Fill(ds2, "teacher");
            metroGrid3.DataSource = ds2.Tables["teacher"];
        }

        private void metroButton7_Click(object sender, EventArgs e)
        {
            ConnectDB.ConnectOnDb();
            string s = metroGrid3[3, metroGrid3.CurrentRow.Index].Value.ToString();
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT id_teacher FROM teacher WHERE (Login = '" + s + "')";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int idTeacher = (int)thisCommand.ExecuteScalar();
            label2.Text = Convert.ToString(idTeacher);

            Form UpT = new UpTeachers();
            UpT.Owner = this;
            UpT.Show();
        }

        private void metroButton9_Click(object sender, EventArgs e)
        {
            Form p = new predmet();
            p.Owner = this;
            p.Show();
        }

        private void metroButton11_Click(object sender, EventArgs e)
        {
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT COUNT(id_predmet) AS kol_vo FROM predmet";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int t = (int)thisCommand.ExecuteScalar();
            t = t + 1;

            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT id_specialnost FROM specialnost  WHERE specialnost  = '" + comboBox5.Text + "'";
            thisCommand1.Connection = ConnectDB.ConnectOnDb();
            int sp = (int)thisCommand1.ExecuteScalar();

            SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand2.CommandText = "SELECT id_kurs FROM kurs WHERE kurs  = '" + comboBox4.Text + "'";
            thisCommand2.Connection = ConnectDB.ConnectOnDb();
            int krs = (int)thisCommand2.ExecuteScalar();

            SqlCeCommand cmd = new SqlCeCommand("Insert into predmet (id_predmet, predmet, id_specialnost, id_kurs) Values ('" + t + "', '" + metroTextBox11.Text + "','" + sp + "','" + krs + "')", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            metroTextBox11.Clear();

            DataSet ds3 = new DataSet();
            SqlCeDataAdapter dtAdapter33 = new SqlCeDataAdapter("SELECT predmet.predmet AS Предмет, specialnost.specialnost AS Специальность, kurs.kurs AS Курс FROM predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost", ConnectDB.ConnectOnDb());
            dtAdapter33.Fill(ds3, "predmet");
            metroGrid4.DataSource = ds3.Tables["predmet"];
        }

        private void metroButton10_Click(object sender, EventArgs e)
        {
            string s = metroGrid4[0, metroGrid4.CurrentRow.Index].Value.ToString();
            string s1 = metroGrid4[1, metroGrid4.CurrentRow.Index].Value.ToString();
            string s2 = metroGrid4[2, metroGrid4.CurrentRow.Index].Value.ToString();

            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT id_specialnost FROM specialnost  WHERE specialnost  = '" + s1 + "'";
            thisCommand1.Connection = ConnectDB.ConnectOnDb();
            int sp = (int)thisCommand1.ExecuteScalar();

            SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand2.CommandText = "SELECT id_kurs FROM kurs WHERE kurs  = '" + s2 + "'";
            thisCommand2.Connection = ConnectDB.ConnectOnDb();
            int krs = (int)thisCommand2.ExecuteScalar();

            ConnectDB.ConnectOnDb();
            SqlCeCommand cmd = new SqlCeCommand("DELETE FROM predmet WHERE predmet = '" + s + "' and id_specialnost = '" + sp + "' and id_kurs = '" + krs + "'", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();

            DataSet ds3 = new DataSet();
            SqlCeDataAdapter dtAdapter33 = new SqlCeDataAdapter("SELECT predmet.predmet AS Предмет, specialnost.specialnost AS Специальность, kurs.kurs AS Курс FROM predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost", ConnectDB.ConnectOnDb());
            dtAdapter33.Fill(ds3, "predmet");
            metroGrid4.DataSource = ds3.Tables["predmet"];
        }

        private void metroButton12_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Data Base File (*.sdf)|*.sdf";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    File.Copy(@"D:\Desktop\diplom\Student.sdf", saveFileDialog1.FileName);
                    MessageBox.Show("Копия базы данных успешно сохранена!", "", MessageBoxButtons.OK);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void metroButton14_Click(object sender, EventArgs e)
        {
            try
                {
                    ConnectDB.ConnectOnDb();
                    SqlCeCommand cmd = new SqlCeCommand("DELETE FROM gurnal", ConnectDB.ConnectOnDb());
                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Все записи из журнала успешно удалены!", "", MessageBoxButtons.OK);
                    ConnectDB.ConnectClose();
                }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void metroButton13_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog();
            OFD.Filter = "Data Base File (*.sdf)|*.sdf";
            OFD.Title = "Открытие файла базы данных";
            OFD.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            if (OFD.ShowDialog() == DialogResult.OK) //если в окне была нажата кнопка "ОК"
            {
                try
                {
                    string filenameopen = OFD.FileName.ToString();
                    string p = @"Data Source=" + filenameopen;

                    StreamWriter print = new StreamWriter("1.ini", false);
                    print.Write(p); 
                    print.Close();



                   
                }
                catch
                {
                    DialogResult rezult = MessageBox.Show("Невозможно открыть выбранный файл", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
            

        
    }
}

/*
 * Copyright 2018, 2019 Виктория Ионова
 * This file is part of Electronic journal.

    Electronic journal  is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    Electronic journal is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with Foobar.  If not, see <https://www.gnu.org/licenses/>.
 */
