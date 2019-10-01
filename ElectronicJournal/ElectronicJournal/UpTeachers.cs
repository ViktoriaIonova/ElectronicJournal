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
    public partial class UpTeachers : MetroFramework.Forms.MetroForm
    {
        public UpTeachers()
        {
            InitializeComponent();
        }

        connect ConnectDB = new connect();
        string Sname, Fname, Lname, login, password;
        string id;

        private void UpTeachers_Load(object sender, EventArgs e)
        {
            adminpanel main = this.Owner as adminpanel;
            if (main != null)
            {
                ConnectDB.ConnectOnDb();
                id = main.label2.Text;
                SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand1.CommandText = "SELECT Sname FROM teacher WHERE (id_teacher = '" + id + "')";
                thisCommand1.Connection = ConnectDB.ConnectOnDb();
                Sname = (string)thisCommand1.ExecuteScalar();
                SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand2.CommandText = "SELECT Fname FROM teacher WHERE(id_teacher = '" + id + "')";
                thisCommand2.Connection = ConnectDB.ConnectOnDb();
                Fname = (string)thisCommand2.ExecuteScalar();
                SqlCeCommand thisCommand3 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand3.CommandText = "SELECT Lname FROM teacher WHERE(id_teacher = '" + id + "')";
                thisCommand3.Connection = ConnectDB.ConnectOnDb();
                Lname = (string)thisCommand3.ExecuteScalar();
                SqlCeCommand thisCommand4 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand4.CommandText = "SELECT Login FROM teacher WHERE(id_teacher = '" + id + "')";
                thisCommand4.Connection = ConnectDB.ConnectOnDb();
                login = (string)thisCommand4.ExecuteScalar();
                SqlCeCommand thisCommand5 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand5.CommandText = "SELECT Password FROM teacher WHERE(id_teacher = '" + id + "')";
                thisCommand5.Connection = ConnectDB.ConnectOnDb();
                password = (string)thisCommand5.ExecuteScalar();
                metroTextBox6.Text = Sname;
                metroTextBox7.Text = Fname;
                metroTextBox8.Text = Lname;
                metroTextBox9.Text = login;
                metroTextBox10.Text = password;
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            SqlCeCommand cmd = new SqlCeCommand("UPDATE teacher SET Sname =  '" + metroTextBox6.Text + "', Fname = '" + metroTextBox7.Text + "', Lname = '" + metroTextBox8.Text + "', Login = '" + metroTextBox9.Text + "', Password = '" + metroTextBox10.Text + "', Prava = '" + 0 + "' WHERE (id_teacher  =  '" + id + "')", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();

            adminpanel main = this.Owner as adminpanel;
            if (main != null)
            {

                DataSet ds2 = new DataSet();
                SqlCeDataAdapter dtAdapter3 = new SqlCeDataAdapter("SELECT Sname AS Фамилия, Fname AS Имя, Lname AS Отчество, Login AS Логин, Password AS Пароль FROM teacher", ConnectDB.ConnectOnDb());
                dtAdapter3.Fill(ds2, "teacher");
                main.metroGrid3.DataSource = ds2.Tables["teacher"];
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
