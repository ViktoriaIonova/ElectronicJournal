using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlServerCe;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace ElectronicJournal
{
    public partial class datastudents : MetroFramework.Forms.MetroForm
    {
        public datastudents()
        {
            InitializeComponent();
        }

        connect ConnectDB = new connect();
        int ct = 0;
        string fio = "";
        private void datastudents_Load(object sender, EventArgs e)
        {
            Journal main = this.Owner as Journal;
            if (main != null)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                ConnectDB.ConnectOnDb();
                SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand.CommandText = "SELECT Fname, Lname, Sname FROM student WHERE id_zachetki = '" + main.label3.Text + "'";
                SqlCeDataReader thisReader = thisCommand.ExecuteReader();
                string res = string.Empty;
                while (thisReader.Read())
                {
                    res += thisReader["Sname"];
                    res += " ";
                    res += thisReader["Fname"];
                    res += " ";
                    res += thisReader["Lname"];
                }
                thisReader.Close();
                fio = res;

                double x = 0;

                SqlCeCommand commandCount = ConnectDB.ConnectOnDb().CreateCommand();
                commandCount.CommandText = "SELECT COUNT(gurnal.ozenka) AS count FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE(gurnal.id_zachetki = '" + main.label3.Text + "') AND(gurnal.id_teacher = '" + main.label1.Text + "') AND(gurnal.id_predmet = '" + main.label2.Text + "') ";
                ct = (int)commandCount.ExecuteScalar();

                SqlCeCommand thisCommandAVG = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommandAVG.CommandText = "SELECT AVG(CONVERT(DECIMAL, gurnal.ozenka)) AS ozenka FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki WHERE(gurnal.id_zachetki = '" + main.label3.Text + "') AND(gurnal.id_vidOzenki = 1000) AND(gurnal.id_teacher = '" + main.label1.Text + "') AND(gurnal.id_predmet = '" + main.label2.Text + "')";
                
                if (!DBNull.Value.Equals(thisCommandAVG.ExecuteScalar()))
                {
                    x = Convert.ToDouble(thisCommandAVG.ExecuteScalar());
                    x = Math.Round(x, 1);
                }
                int x2 = 0;

                SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand2.CommandText = "SELECT gurnal.ozenka FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki WHERE(gurnal.id_zachetki = '" + main.label3.Text + "') AND(gurnal.id_vidOzenki = 1001) AND(gurnal.id_teacher = '" + main.label1.Text + "') AND(gurnal.id_predmet = '" + main.label2.Text + "')";
                if (thisCommand2.ExecuteScalar() != null)
                    x2 = (int)thisCommand2.ExecuteScalar();

                int x3 = 0;
                SqlCeCommand thisCommand3 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand3.CommandText = "SELECT gurnal.ozenka FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki WHERE (gurnal.id_zachetki = '" + main.label3.Text + "') AND (gurnal.id_vidOzenki = 1002) AND (gurnal.id_teacher = '" + main.label1.Text + "') AND (gurnal.id_predmet = '" + main.label2.Text + "')";
                
                if (thisCommand3.ExecuteScalar() != null)
                    x3 = (int)thisCommand3.ExecuteScalar();

                int x4 = 0;
                SqlCeCommand thisCommand4 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand4.CommandText = "SELECT gurnal.ozenka FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki WHERE(gurnal.id_zachetki = '" + main.label3.Text + "') AND(gurnal.id_vidOzenki = 1003) AND(gurnal.id_teacher = '" + main.label1.Text + "') AND(gurnal.id_predmet = '" + main.label2.Text + "')";
                
                if (thisCommand4.ExecuteScalar() != null)
                    x4 = (int)thisCommand4.ExecuteScalar();

                int x5 = 0;
                SqlCeCommand thisCommand5 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand5.CommandText = "SELECT gurnal.ozenka FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki WHERE(gurnal.id_zachetki = '" + main.label3.Text + "') AND(gurnal.id_vidOzenki = 1004) AND(gurnal.id_teacher = '" + main.label1.Text + "') AND(gurnal.id_predmet = '" + main.label2.Text + "')";
                
                if (thisCommand5.ExecuteScalar() != null)
                    x5 = (int)thisCommand5.ExecuteScalar();
                if (x != 0)
                    metroLabel9.Text = Convert.ToString(x);
                else metroLabel9.Text = "Нет оценки";
                if (x2 != 0)
                    metroLabel10.Text = Convert.ToString(x2);
                else metroLabel10.Text = "Нет оценки";
                if (x3 != 0)
                    metroLabel11.Text = Convert.ToString(x3);
                else metroLabel11.Text = "Нет оценки";
                if (x4 != 0)
                    metroLabel12.Text = Convert.ToString(x4);
                else metroLabel12.Text = "Нет оценки";
                if (x5 != 0)
                    metroLabel13.Text = Convert.ToString(x5);
                else metroLabel13.Text = "Нет оценки";

                ConnectDB.ConnectClose();
                metroLabel1.Text = fio;
                metroLabel2.Text = main.comboBox2.Text;
                metroLabel3.Text = main.comboBox1.Text;

                
                DataSet ds1 = new DataSet();
                SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT gurnal.ozenka AS Оценка, gurnal.date AS Дата, vidOzenki.vidOzenki AS [Вид оценки], tema.tema AS Тема FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN student ON gurnal.id_zachetki = student.id_zachetki INNER JOIN tema ON gurnal.id_tema = tema.id_tema INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki WHERE (predmet.predmet = '" + main.comboBox2.Text + "') AND (gurnal.id_zachetki = '" + main.label3.Text + "')", ConnectDB.ConnectOnDb());
                dtAdapter.Fill(ds1, "gurnal");
                metroGrid1.DataSource = ds1.Tables["gurnal"];
            
            } 
        }

        private void datastudents_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            int rows = metroGrid1.RowCount + 1;
            int columns = 4;
            ConnectDB.ConnectOnDb();
            System.Data.DataTable dt = new System.Data.DataTable();
            

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add();

           
            doc.Paragraphs[1].Range.Text = "Предмет: " + metroLabel2.Text;
            doc.Paragraphs[1].Range.Font.Size = 10;
            doc.Paragraphs[1].Range.Font.Name = "Times New Roman";
            doc.Paragraphs.Add();

            doc.Paragraphs[2].Range.Text = fio + ", " + metroLabel3.Text;
            doc.Paragraphs[2].Range.Font.Size = 10;
            doc.Paragraphs[2].Range.Font.Name = "Times New Roman";
            doc.Paragraphs[2].Range.Font.Bold = 1;            
            doc.Paragraphs.Add();

            Table t = doc.Tables.Add(doc.Paragraphs[3].Range, ct + 1, 4);
            t.Borders.Enable = 1;

            t.Rows[1].Cells[1].Range.Text = "Оценка";
            t.Rows[1].Cells[1].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[1].Range.Font.Size = 10;
            t.Rows[1].Cells[1].Range.Font.Bold = 0;
            t.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[2].Range.Text = "Дата";
            t.Rows[1].Cells[2].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[2].Range.Font.Size = 10;
            t.Rows[1].Cells[2].Range.Font.Bold = 0;
            t.Rows[1].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[3].Range.Text = "Вид оценки";
            t.Rows[1].Cells[3].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[3].Range.Font.Size = 10;
            t.Rows[1].Cells[3].Range.Font.Bold = 0;
            t.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[4].Range.Text = "Тема";
            t.Rows[1].Cells[4].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[4].Range.Font.Size = 10;
            t.Rows[1].Cells[4].Range.Font.Bold = 0;
            t.Rows[1].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            string data = "";
            for (int i = 0; i < rows - 1; i++)
            {
                for (int j = 0; j < columns; j++)
                {
                    if (j == 1)
                    {
                        data = Convert.ToString(metroGrid1[j, i].Value);
                        doc.Tables[1].Cell(i + 2, j + 1).Range.Text = data.Substring(0, 11);
                    }
                    else doc.Tables[1].Cell(i + 2, j + 1).Range.Text = metroGrid1[j, i].Value.ToString();
                    doc.Tables[1].Cell(i + 2, j + 1).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, j + 1).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, j + 1).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, j + 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
            
            doc.SaveAs(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\Otchet.docx");
            app.Documents.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\Otchet.docx", ReadOnly: true);
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
