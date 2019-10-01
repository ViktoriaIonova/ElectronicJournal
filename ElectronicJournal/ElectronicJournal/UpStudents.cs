using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;

namespace ElectronicJournal
{
    public partial class UpStudents : MetroFramework.Forms.MetroForm
    {
        public UpStudents()
        {
            InitializeComponent();
        }

        connect ConnectDB = new connect();

        string id;
        string Sname, Fname, Lname, gruppa;
        int idGr;

        private void UpStudents_Load(object sender, EventArgs e)
        {
            adminpanel main = this.Owner as adminpanel;
            if (main != null)
            {
                ConnectDB.ConnectOnDb();
                id = main.label1.Text;
                SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand1.CommandText = "SELECT Sname FROM student WHERE (id_zachetki = '" + id + "')";
                thisCommand1.Connection = ConnectDB.ConnectOnDb();
                Sname = (string)thisCommand1.ExecuteScalar();
                SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand2.CommandText = "SELECT Fname FROM student WHERE(id_zachetki = '" + id + "')";
                thisCommand2.Connection = ConnectDB.ConnectOnDb();
                Fname = (string)thisCommand2.ExecuteScalar();
                SqlCeCommand thisCommand3 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand3.CommandText = "SELECT Lname FROM student WHERE(id_zachetki = '" + id + "')";
                thisCommand3.Connection = ConnectDB.ConnectOnDb();
                Lname = (string)thisCommand3.ExecuteScalar();
                SqlCeCommand thisCommand4 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand4.CommandText = "SELECT gruppa.gruppa FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE(student.id_zachetki = '" + id + "')";
                thisCommand4.Connection = ConnectDB.ConnectOnDb();
                gruppa = (string)thisCommand4.ExecuteScalar();

                metroTextBox2.Text = id;
                metroTextBox3.Text = Sname;
                metroTextBox4.Text = Fname;
                metroTextBox5.Text = Lname;

                DataTable dt = new DataTable();
                string q = "SELECT DISTINCT gruppa FROM gruppa";
                SqlCeDataAdapter dataAdapter2 = new SqlCeDataAdapter(q, ConnectDB.ConnectOnDb());
                dataAdapter2.Fill(dt);
                comboBox3.DataSource = dt;
                comboBox3.DisplayMember = "gruppa";
                comboBox3.ValueMember = "gruppa";

                comboBox3.Text = gruppa;

                SqlCeCommand thisCommand5 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand5.CommandText = "SELECT id_gruppa FROM gruppa WHERE gruppa = '" + comboBox3.Text + "'";
                thisCommand5.Connection = ConnectDB.ConnectOnDb();
                idGr = (int)thisCommand5.ExecuteScalar();
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            SqlCeCommand cmd = new SqlCeCommand("UPDATE student SET id_zachetki =  '" + metroTextBox2.Text + "', Sname =  '" + metroTextBox3.Text + "', Fname = '" + metroTextBox4.Text + "', Lname = '" + metroTextBox5.Text + "', id_gruppa = '" + idGr + "' WHERE (id_zachetki  =  '" + id + "')", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            adminpanel main = this.Owner as adminpanel;
            if (main != null)
            {
                DataSet ds4 = new DataSet();
                SqlCeDataAdapter dtAdapter6 = new SqlCeDataAdapter("SELECT student.id_zachetki AS [Номер зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчество, gruppa.gruppa AS Группа FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE(gruppa.gruppa = '" + comboBox3.Text + "')", ConnectDB.ConnectOnDb());
                dtAdapter6.Fill(ds4, "student");
                main.metroGrid2.DataSource = ds4.Tables["student"];
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
