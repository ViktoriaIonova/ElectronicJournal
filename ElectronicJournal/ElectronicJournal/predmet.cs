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
    public partial class predmet : MetroFramework.Forms.MetroForm
    {
        public predmet()
        {
            InitializeComponent();
        }

        connect ConnectDB = new connect();
        
        List<int> id = new List<int>();

        private void predmet_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            string q = "SELECT id_teacher, [Sname] + ' ' + [Fname] + ' ' + [Lname] as fio FROM teacher WHERE (Prava NOT IN (1))";
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter(q, ConnectDB.ConnectOnDb());
            SqlCeCommand sc = new SqlCeCommand(q, ConnectDB.ConnectOnDb());
            SqlCeDataReader sdr;
            sdr = sc.ExecuteReader();
            while (sdr.Read())
            {
                id.Add(sdr.GetInt32(0));
            }
            dtAdapter.Fill(dt);
            metroComboBox1.DataSource = dt;
            metroComboBox1.DisplayMember = "fio";
            metroComboBox1.DisplayMember = "fio";
            metroComboBox1.ValueMember = "id_teacher";

            DataTable dt2 = new DataTable();
            string q2 = "SELECT specialnost FROM specialnost";
            SqlCeDataAdapter dtAdapter3 = new SqlCeDataAdapter(q2, ConnectDB.ConnectOnDb());
            dtAdapter3.Fill(dt2);
            metroComboBox2.DataSource = dt2;
            metroComboBox2.DisplayMember = "specialnost";
            metroComboBox2.ValueMember = "specialnost";
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter2 = new SqlCeDataAdapter("SELECT DISTINCT predmet.predmet AS Предмет, gruppa.gruppa AS Группа FROM pr_pr INNER JOIN predmet ON pr_pr.id_predmet = predmet.id_predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost AND gruppa.id_specialnost = specialnost.id_specialnost WHERE (pr_pr.id_teacher = '" + id[metroComboBox1.SelectedIndex] + "')", ConnectDB.ConnectOnDb());
            dtAdapter2.Fill(ds, "predmet");
            metroGrid1.DataSource = ds.Tables["predmet"];
        }

        private void metroComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataSet ds2 = new DataSet();
            SqlCeDataAdapter dtAdapter3 = new SqlCeDataAdapter("SELECT DISTINCT predmet.predmet AS Предмет, gruppa.gruppa AS Группа FROM predmet INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN gruppa ON specialnost.id_specialnost = gruppa.id_specialnost AND kurs.id_kurs = gruppa.id_kurs WHERE(specialnost.specialnost = '" + metroComboBox2.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter3.Fill(ds2, "predmet");
            metroGrid2.DataSource = ds2.Tables["predmet"];
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            string s = metroGrid2[0, metroGrid2.CurrentRow.Index].Value.ToString();
            string s2 = metroGrid2[1, metroGrid2.CurrentRow.Index].Value.ToString();
            ConnectDB.ConnectOnDb();
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT DISTINCT predmet.id_predmet FROM predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs INNER JOIN  specialnost ON predmet.id_specialnost = specialnost.id_specialnost AND gruppa.id_specialnost = specialnost.id_specialnost WHERE(predmet.predmet = '" + s + "') AND (gruppa.gruppa = '" + s2 + "')";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int x = (int)thisCommand.ExecuteScalar();

            SqlCeCommand cmd = new SqlCeCommand("Insert into pr_pr (id_predmet, id_teacher) Values ('" + x + "', '" + id[metroComboBox1.SelectedIndex] + "')", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter2 = new SqlCeDataAdapter("SELECT DISTINCT predmet.predmet AS Предмет, gruppa.gruppa AS Группа FROM pr_pr INNER JOIN predmet ON pr_pr.id_predmet = predmet.id_predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost AND gruppa.id_specialnost = specialnost.id_specialnost WHERE (pr_pr.id_teacher = '" + id[metroComboBox1.SelectedIndex] + "')", ConnectDB.ConnectOnDb());
            dtAdapter2.Fill(ds, "predmet");
            metroGrid1.DataSource = ds.Tables["predmet"];
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string s = metroGrid1[0, metroGrid1.CurrentRow.Index].Value.ToString();
            string s2 = metroGrid1[1, metroGrid1.CurrentRow.Index].Value.ToString();
            ConnectDB.ConnectOnDb();
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT DISTINCT predmet.id_predmet FROM predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs INNER JOIN  specialnost ON predmet.id_specialnost = specialnost.id_specialnost AND gruppa.id_specialnost = specialnost.id_specialnost WHERE(predmet.predmet = '" + s + "') AND (gruppa.gruppa = '" + s2 + "')";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int x = (int)thisCommand.ExecuteScalar();
           // button1.Text = Convert.ToString(x);
            SqlCeCommand cmd = new SqlCeCommand("DELETE FROM pr_pr WHERE id_predmet = '" + x + "' AND id_teacher = '" + id[metroComboBox1.SelectedIndex] + "' ", ConnectDB.ConnectOnDb());
            cmd.ExecuteNonQuery();
            ConnectDB.ConnectClose();
            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter2 = new SqlCeDataAdapter("SELECT DISTINCT predmet.predmet AS Предмет, gruppa.gruppa AS Группа FROM pr_pr INNER JOIN predmet ON pr_pr.id_predmet = predmet.id_predmet INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost AND gruppa.id_specialnost = specialnost.id_specialnost WHERE (pr_pr.id_teacher = '" + id[metroComboBox1.SelectedIndex] + "')", ConnectDB.ConnectOnDb());
            dtAdapter2.Fill(ds, "predmet");
            metroGrid1.DataSource = ds.Tables["predmet"];           
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
