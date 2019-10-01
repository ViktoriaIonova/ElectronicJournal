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
using xNet;
using System.Threading;

namespace ElectronicJournal
{
    public partial class Journal : MetroFramework.Forms.MetroForm
    {
        public Journal()
        {
            InitializeComponent();
        }

        connect ConnectDB = new connect();
        int ks = 0;
        int x = 0;
        int x1 = 0;
        string o = "";
        string n = "";
        double sru = 0, srk = 0, srsr = 0;
        double kp2 = 0;
        string log = "";

        string data = "";
     

        private void Journal_Load(object sender, EventArgs e)
        {
            Form1 main = this.Owner as Form1;
            if (main != null)
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                ConnectDB.ConnectOnDb();
                SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand.CommandText = "SELECT Fname, Lname  FROM teacher WHERE Login = '" + main.metroTextBox1.Text + "'";
                SqlCeDataReader thisReader = thisCommand.ExecuteReader();
                string res = string.Empty;
                while (thisReader.Read())
                {
                    res += thisReader["Fname"];
                    res += " ";
                    res += thisReader["Lname"];
                }
                thisReader.Close();
                metroLabel1.Text = "Здравствуйте, " + res;
                log = main.metroTextBox1.Text;

                string q1 = "SELECT DISTINCT gruppa.gruppa FROM teacher INNER JOIN pr_pr ON teacher.id_teacher = pr_pr.id_teacher INNER JOIN predmet ON pr_pr.id_predmet = predmet.id_predmet INNER JOIN student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost ON predmet.id_kurs = kurs.id_kurs AND predmet.id_specialnost = specialnost.id_specialnost WHERE teacher.Login = '" + main.metroTextBox1.Text + "'";
                //"SELECT DISTINCT gruppa.gruppa FROM gruppa INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost INNER JOIN predmet ON specialnost.id_specialnost = predmet.id_specialnost INNER JOIN pr_pr ON predmet.id_predmet = pr_pr.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher WHERE Login = '" + main.textBox1.Text + "'";

                SqlCeDataAdapter dataAdapter = new SqlCeDataAdapter(q1, ConnectDB.ConnectOnDb());
                dataAdapter.Fill(dt);
                comboBox1.DataSource = dt;
                comboBox1.DisplayMember = "gruppa";
                comboBox1.ValueMember = "gruppa";

                SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand1.CommandText = "SELECT id_teacher FROM teacher WHERE login = '" + main.metroTextBox1.Text + "'";
                thisCommand1.Connection = ConnectDB.ConnectOnDb();
                int x1 = (int)thisCommand1.ExecuteScalar();
                label1.Text = Convert.ToString(x1);
                ConnectDB.ConnectClose();
                /* DateTime dt2 = new DateTime();
                 dt2 = DateTime.Now;
                 label3.Text = Convert.ToString(dt2);*/


            }

        }



        private void metroButton1_Click(object sender, EventArgs e)
        {
            string o1 = ""; int v1 = 0;
            for (int i = 0; i < ks; i++)
            {
                o1 = metroGrid1.Rows[i].Cells[4].Value.ToString();
                if (o1.Length > 0)
                    v1++;
            }
            if (v1 == 0) MessageBox.Show("Нет оценок!");
            else
            {
                int v = 0;
               // SqlCeCommand thisCommand3 = ConnectDB.ConnectOnDb().CreateCommand();

                SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand.CommandText = "SELECT id_tema FROM tema WHERE tema = '" + comboBox3.Text + "'";
                thisCommand.Connection = ConnectDB.ConnectOnDb();
                int x2 = (int)thisCommand.ExecuteScalar();

                DateTime dt = new DateTime();
                dt = DateTime.Now;

                if (metroRadioButton1.Checked == true)
                    v = 1000;
                if (metroRadioButton2.Checked == true)
                    v = 1001;
                if (metroRadioButton3.Checked == true)
                    v = 1002;
                if (metroRadioButton4.Checked == true)
                    v = 1003;
                if (metroRadioButton5.Checked == true)
                    v = 1004;

                SqlCeCommand thisCommand5 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand5.CommandText = "SELECT id_teacher FROM teacher WHERE login = '" + log + "'";
                thisCommand5.Connection = ConnectDB.ConnectOnDb();
                x1 = (int)thisCommand5.ExecuteScalar();
                for (int i = 0; i < ks; i++)
                {
                    o = metroGrid1.Rows[i].Cells[4].Value.ToString();
                    if (o.Length > 0)
                    {
                        n = metroGrid1.Rows[i].Cells[0].Value.ToString();
                        Convert.ToInt32(n);
                        Convert.ToInt32(o);
                        //thisCommand3.Connection = ConnectDB.ConnectOnDb();

                        string comm = "INSERT INTO gurnal (id_zachetki, id_predmet, id_teacher, id_vidOzenki, ozenka, id_tema, date) VALUES (@idZachetki, @idPredmet, @idTeacher, @idVOzenki, @ozenka, @idTema, @SomeDateValue)";
                        SqlCeCommand cmd = new SqlCeCommand(comm, ConnectDB.ConnectOnDb());
                        cmd.Parameters.AddWithValue("@SomeDateValue", dt);
                        cmd.Parameters.AddWithValue("@idZachetki", n);
                        cmd.Parameters.AddWithValue("@idPredmet", x);
                        cmd.Parameters.AddWithValue("@idTeacher", x1);
                        cmd.Parameters.AddWithValue("@idVOzenki", v);
                        cmd.Parameters.AddWithValue("@ozenka", o);
                        cmd.Parameters.AddWithValue("@idTema", x2);
                        cmd.ExecuteNonQuery();
                    }
                }
                ConnectDB.ConnectClose();
                ConnectDB.ConnectDispose();
                
             

                AuthServer();

                MetroFramework.MetroMessageBox.Show(this, "Оценки добавлены!", " ", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                //MessageBox.Show("Оценки добавлены");

                // для mysql
                
            }

            DataSet ds2 = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT student.id_zachetki AS [№ зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчества, NULL AS Оценка FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE (gruppa.gruppa = '" + comboBox1.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds2, "student");
            metroGrid1.DataSource = ds2.Tables["student"];
            metroGrid1.Columns[0].Width = 93;
            metroGrid1.Columns[1].Width = 93;
            metroGrid1.Columns[2].Width = 93;
            metroGrid1.Columns[3].Width = 93;
            metroGrid1.Columns[4].Width = 92;

        }

        private void AuthServer()
        {            
            HttpRequest p = new HttpRequest();
            p.UserAgent = Http.ChromeUserAgent();
            RequestParams pd = new RequestParams();

            string v = "";

            string dt = DateTime.Now.ToString("dd.MM.yyyy");

            if (metroRadioButton1.Checked == true)
                v = "Текущая";
            if (metroRadioButton2.Checked == true)
                v = "Семестровая";
            if (metroRadioButton3.Checked == true)
                v = "Экзаменационная";
            if (metroRadioButton4.Checked == true)
                v = "Годовая";
            if (metroRadioButton5.Checked == true)
                v = "Дипломная";

            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT Sname, Fname, Lname  FROM teacher WHERE id_teacher = '" + label1.Text + "'";
            SqlCeDataReader thisReader1 = thisCommand1.ExecuteReader();
            string res1 = string.Empty;
            while (thisReader1.Read())
            {
                res1 += thisReader1["Sname"];
                res1 += " ";
                res1 += thisReader1["Fname"];
                res1 += " ";
                res1 += thisReader1["Lname"];
            }
            thisReader1.Close();

            for (int i = 0; i < ks; i++)
            {
                o = metroGrid1.Rows[i].Cells[4].Value.ToString();
                if (o.Length > 0)
                {
                    n = metroGrid1.Rows[i].Cells[0].Value.ToString();
                    Convert.ToInt32(n);
                    Convert.ToInt32(o);
                    pd["id_zachetki"] = n;
                    pd["predmet"] = comboBox2.Text;
                    pd["teacher"] = res1;
                    pd["vidozenki"] = v;
                    pd["ozenka"] = o;
                    pd["tema"] = comboBox3.Text;
                    pd["date"] = dt;
                    data = p.Post("http://q961075i.beget.tech/gurnal.php", pd).ToString();
                }
            }
        }
        
        private void Journal_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }





        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            string q2 = "SELECT DISTINCT predmet.predmet FROM predmet INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost INNER JOIN gruppa ON specialnost.id_specialnost = gruppa.id_specialnost INNER JOIN pr_pr ON predmet.id_predmet = pr_pr.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs AND gruppa.id_kurs = kurs.id_kurs WHERE(gruppa.gruppa = '" + comboBox1.Text + "') AND(teacher.Login = '" + log + "')";
            SqlCeDataAdapter dataAdapter2 = new SqlCeDataAdapter(q2, ConnectDB.ConnectOnDb());
            dataAdapter2.Fill(dt);
            comboBox2.DataSource = dt;
            comboBox2.DisplayMember = "predmet";
            comboBox2.ValueMember = "predmet";

            DataSet ds = new DataSet();
            SqlCeDataAdapter dtAdapter = new SqlCeDataAdapter("SELECT student.id_zachetki AS [№ зачетки], student.Sname AS Фамилия, student.Fname AS Имя, student.Lname AS Отчества, NULL AS Оценка FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE (gruppa.gruppa = '" + comboBox1.Text + "')", ConnectDB.ConnectOnDb());
            dtAdapter.Fill(ds, "student");
            metroGrid1.DataSource = ds.Tables["student"];

            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT COUNT(student.id_zachetki) AS Expr1 FROM  student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE (gruppa.gruppa = '" + comboBox1.Text + "')";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            ks = (int)thisCommand.ExecuteScalar();
            ConnectDB.ConnectClose();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT id_predmet FROM predmet WHERE predmet = '" + comboBox2.Text + "'";
            thisCommand1.Connection = ConnectDB.ConnectOnDb();
            x = (int)thisCommand1.ExecuteScalar();
            label2.Text = Convert.ToString(x);

            // тема
            comboBox3.DisplayMember = "";
            /*  if (metroGrid1.SelectedRows.Count != 0)
                  metroGrid1.Rows.Clear();*/

            System.Data.DataTable dt3 = new System.Data.DataTable();
            string q2 = "SELECT DISTINCT tema.tema FROM pr_pr INNER JOIN predmet ON pr_pr.id_predmet = predmet.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher INNER JOIN tema ON predmet.id_predmet = tema.id_predmet WHERE (predmet.predmet = '" + comboBox2.Text + "') AND (teacher.Login = '" + log + "')";
            SqlCeDataAdapter dataAdapter2 = new SqlCeDataAdapter(q2, ConnectDB.ConnectOnDb());
            dataAdapter2.Fill(dt3);
            comboBox3.DataSource = dt3;
            comboBox3.DisplayMember = "tema";
            comboBox3.ValueMember = "tema";
            comboBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDown;
            comboBox3.AutoCompleteSource = AutoCompleteSource.ListItems;

            /* SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
             thisCommand2.CommandText = "SELECT id_predmet FROM predmet WHERE predmet = '" + metroComboBox2.Text + "'";
             thisCommand2.Connection = ConnectDB.ConnectOnDb();
             //int x = (int)thisCommand2.ExecuteScalar();
             //label2.Text = Convert.ToString(x);
               int x1 = (int)thisCommand2.ExecuteScalar();
               label2.Text = Convert.ToString(x1);*/

        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT COUNT(id_tema) AS kol_vo FROM tema";
            thisCommand.Connection = ConnectDB.ConnectOnDb();
            int t = (int)thisCommand.ExecuteScalar();
            t = t + 1;

            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT id_predmet  FROM predmet WHERE predmet = '" + comboBox2.Text + "'";
            thisCommand1.Connection = ConnectDB.ConnectOnDb();
            int p = (int)thisCommand1.ExecuteScalar();

            string comm = "INSERT INTO tema (id_tema, id_predmet, tema) VALUES (@idTema, @idPredmet, @tema)";
            SqlCeCommand cmd = new SqlCeCommand(comm, ConnectDB.ConnectOnDb());
            cmd.Parameters.AddWithValue("@idTema", t);
            cmd.Parameters.AddWithValue("@idPredmet", p);
            cmd.Parameters.AddWithValue("@tema", comboBox3.Text);
            cmd.ExecuteNonQuery();

            System.Data.DataTable dt3 = new System.Data.DataTable();
            string q2 = "SELECT DISTINCT tema.tema FROM pr_pr INNER JOIN predmet ON pr_pr.id_predmet = predmet.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher INNER JOIN tema ON predmet.id_predmet = tema.id_predmet WHERE (predmet.predmet = '" + comboBox2.Text + "') AND (teacher.Login = '" + log + "')";
            SqlCeDataAdapter dataAdapter2 = new SqlCeDataAdapter(q2, ConnectDB.ConnectOnDb());
            dataAdapter2.Fill(dt3);
            comboBox3.DataSource = dt3;
            comboBox3.DisplayMember = "tema";
            comboBox3.ValueMember = "tema";
            comboBox3.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            comboBox3.AutoCompleteSource = AutoCompleteSource.ListItems;
            MessageBox.Show("Тема добавлена!");
        }

        private void metroButton7_Click(object sender, EventArgs e)
        {
            string s = metroGrid1[0, metroGrid1.CurrentRow.Index].Value.ToString();
            label3.Text = Convert.ToString(s);

            // наследование
            Form ds = new datastudents();
            ds.Owner = this;
            ds.Show();
            // ds.Text = s;

            Form1 main = this.Owner as Form1;
            if (main != null)
            {
                SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand1.CommandText = "SELECT id_teacher FROM teacher WHERE login = '" + main.metroTextBox1.Text + "'";
                thisCommand1.Connection = ConnectDB.ConnectOnDb();
                int x1 = (int)thisCommand1.ExecuteScalar();
                label1.Text = Convert.ToString(x1);
            }
        }

        private void metroGrid1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            int res;
            if (e.ColumnIndex == 4)
            {
                if (e.FormattedValue.ToString() == string.Empty)
                    return;
                else
                    if ((!int.TryParse(e.FormattedValue.ToString(), out res) || e.FormattedValue.ToString().Length > 1) || (Convert.ToInt32(e.FormattedValue) < 2 || (Convert.ToInt32(e.FormattedValue) > 5)))
                    {
                        metroGrid1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 1;
                        MessageBox.Show("Введите корректную оценку");
                        e.Cancel = true;
                        return;
                    }

            }
        }

        string FIOT = "";

        private void metroButton4_Click(object sender, EventArgs e)
        {
            int rows = ks + 1;
            int columns = 7;

            System.Data.DataTable dt = new System.Data.DataTable();
            ConnectDB.ConnectOnDb();
            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT student.id_zachetki, student.Sname, student.Fname, student.Lname FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE (gruppa.gruppa = '" + comboBox1.Text + "') ORDER BY Sname";
            SqlCeDataReader thisReader = thisCommand.ExecuteReader();
            string[] fam = new string[ks];
            string[] im = new string[ks];
            string[] ot = new string[ks];
            string[] iz = new string[ks];
            int j = 0;
            while (thisReader.Read())
            {
                fam[j] = Convert.ToString(thisReader["Sname"]);
                im[j] = Convert.ToString(thisReader["Fname"]);
                ot[j] = Convert.ToString(thisReader["Lname"]);
                iz[j] = Convert.ToString(thisReader["id_zachetki"]);
                j++;
            }
            thisReader.Close();

            System.Data.DataTable dt1 = new System.Data.DataTable();
            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT Sname, Fname, Lname  FROM teacher WHERE id_teacher = '" + label1.Text + "'";
            SqlCeDataReader thisReader1 = thisCommand1.ExecuteReader();
            string res1 = string.Empty;
            while (thisReader1.Read())
            {
                res1 += thisReader1["Sname"];
                res1 += " ";
                res1 += thisReader1["Fname"];
                res1 += " ";
                res1 += thisReader1["Lname"];
            }
            thisReader1.Close();

            FIOT = res1;
            string d = DateTime.Now.ToString("dd.MM.yyyy");

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\Vedomost.docx");

            doc.Bookmarks["pr1"].Range.Text = comboBox2.Text;
            doc.Bookmarks["pr2"].Range.Text = "                                                                                    " + d + "                                                                         " + comboBox1.Text;
            doc.Bookmarks["pr3"].Range.Text = FIOT;

            Table t = doc.Tables.Add(doc.Bookmarks["pr4"].Range, ks + 1, 7);
            t.Borders.Enable = 1;

            foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
            {
                table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
            }

            t.Rows[1].Cells[1].Range.Text = "№ п/п";
            t.Rows[1].Cells[1].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[1].Range.Font.Size = 10;
            t.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[2].Range.Text = "№ экзамен листа";
            t.Rows[1].Cells[2].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[2].Range.Font.Size = 10;
            t.Rows[1].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[3].Range.Text = "Фамилия";
            t.Rows[1].Cells[3].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[3].Range.Font.Size = 10;
            t.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[4].Range.Text = "Имя";
            t.Rows[1].Cells[4].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[4].Range.Font.Size = 10;
            t.Rows[1].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[5].Range.Text = "Отчество";
            t.Rows[1].Cells[5].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[5].Range.Font.Size = 10;
            t.Rows[1].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[6].Range.Text = "Отметка";
            t.Rows[1].Cells[6].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[6].Range.Font.Size = 10;
            t.Rows[1].Cells[6].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[7].Range.Text = "Подпись экзаменатора";
            t.Rows[1].Cells[7].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[7].Range.Font.Size = 10;
            t.Rows[1].Cells[7].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            for (int i = 0; i < rows - 1; i++)
            {
                for (int js = 0; js < ks; js++)
                {
                    doc.Tables[1].Rows[i + 2].Range.Text = Convert.ToString(i + 1);
                    doc.Tables[1].Rows[i + 2].Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Rows[i + 2].Range.Font.Size = 10;
                    doc.Tables[1].Rows[i + 2].Range.Font.Bold = 0;
                    doc.Tables[1].Rows[i + 2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    doc.Tables[1].Cell(i + 2, 3).Range.Text = fam[i];
                    doc.Tables[1].Cell(i + 2, 3).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 3).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 3).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    doc.Tables[1].Cell(i + 2, 4).Range.Text = im[i];
                    doc.Tables[1].Cell(i + 2, 4).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 4).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 4).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    doc.Tables[1].Cell(i + 2, 5).Range.Text = ot[i];
                    doc.Tables[1].Cell(i + 2, 5).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 5).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 5).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


                    SqlCeCommand thisCommand8;
                    SqlCeDataReader sqlReader;

                    thisCommand8 = new SqlCeCommand("SELECT COUNT (ozenka) FROM gurnal WHERE id_zachetki = '" + iz[i] + "' AND id_VidOzenki = '1002' AND id_predmet = '" + x + "'", ConnectDB.ConnectOnDb());

                    sqlReader = thisCommand8.ExecuteReader();
                    string oz = "";
                    sqlReader.Read();

                    int c = sqlReader.GetInt32(0);
                    sqlReader.Close();
                    if (c == 0)
                    {
                        oz = "Не явился";
                    }
                    else
                    {
                        thisCommand8.CommandText = "SELECT ozenka FROM gurnal WHERE id_zachetki = '" + iz[i] + "' AND id_VidOzenki = '1002' AND id_predmet = '" + x + "'";
                        thisCommand8.Connection = ConnectDB.ConnectOnDb();


                        int x8 = 0;
                        if (Convert.ToString(thisCommand8.Connection).Length == 0) oz = "Не допущен";
                        else x8 = (int)thisCommand8.ExecuteScalar();

                        switch (x8)
                        {
                            case 2: { oz = "2 (неудов.)"; break; }
                            case 3: { oz = "3 (удов.)"; break; }
                            case 4: { oz = "4 (хорошо)"; break; }
                            case 5: { oz = "5 (отлично)"; break; }
                            default: break;
                        }
                    }
                    doc.Tables[1].Cell(i + 2, 6).Range.Text = oz;
                    doc.Tables[1].Cell(i + 2, 6).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 6).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 6).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 6).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
            doc.SaveAs(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocPrintOut.docx");
            app.Documents.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocPrintOut.docx");

        }

        private void metroButton6_Click(object sender, EventArgs e)
        {
            int kp = 0;

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\Itog.docx");

            ConnectDB.ConnectOnDb();

            SqlCeCommand thisCommand2 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand2.CommandText = "SELECT COUNT(predmet.predmet) AS KolPR FROM predmet INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost INNER JOIN gruppa ON specialnost.id_specialnost = gruppa.id_specialnost INNER JOIN pr_pr ON predmet.id_predmet = pr_pr.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs AND gruppa.id_kurs = kurs.id_kurs WHERE (teacher.id_teacher = '" + label1.Text + "')";
            kp = (int)thisCommand2.ExecuteScalar();
            kp2 = (double)kp;

            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT DISTINCT predmet.predmet FROM predmet INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost INNER JOIN gruppa ON specialnost.id_specialnost = gruppa.id_specialnost INNER JOIN pr_pr ON predmet.id_predmet = pr_pr.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher INNER JOIN kurs ON predmet.id_kurs = kurs.id_kurs AND gruppa.id_kurs = kurs.id_kurs WHERE (teacher.id_teacher = '" + label1.Text + "')";
            SqlCeDataReader thisReader1 = thisCommand1.ExecuteReader();
            string[] predmet = new string[kp];
            int j = 0;
            while (thisReader1.Read())
            {
                predmet[j] = Convert.ToString(thisReader1["predmet"]);
                j++;
            }
            thisReader1.Close();

            int idT = 212;


            SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand.CommandText = "SELECT Sname, Fname, Lname  FROM teacher WHERE id_teacher = '212'";
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


            DateTime thisDay = DateTime.Today;
            string data = Convert.ToString(thisDay);
            string m = data.Substring(3, 2);
            string g = data.Substring(6, 4);
            string g2 = Convert.ToString(Convert.ToInt32(g) + 1);
            int zm = Convert.ToInt32(m);
            string zn;
            string d1 = g + "-01-01";
            string d2 = g + "-06-30";
            string d3 = g + "-09-01";
            string d4 = g + "-12-31";
           // if (((zm >= 10) && (zm <= 12)) || ((zm >= 1) && (zm <= 2)))
            if ((zm >= 1) && (zm <= 6))
            {
                zn = "2";
                g = Convert.ToString(Convert.ToInt32(g) - 1);
                g2 = data.Substring(6, 4);
                doc.Bookmarks["pr"].Range.Text = zn;
                doc.Bookmarks["pr2"].Range.Text = g + '-' + g2;
                doc.Bookmarks["pr3"].Range.Text = res;

                Table t = doc.Tables.Add(doc.Bookmarks["pr4"].Range, kp + 1, 5);
                t.Borders.Enable = 1;

                foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
                {
                    table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
                }

                t.Rows[1].Cells[1].Range.Text = "Дисциплина";
                t.Rows[1].Cells[1].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[1].Range.Font.Size = 14;
                t.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[2].Range.Text = "Группа";
                t.Rows[1].Cells[2].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[2].Range.Font.Size = 14;
                t.Rows[1].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[3].Range.Text = "% успев.";
                t.Rows[1].Cells[3].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[3].Range.Font.Size = 14;
                t.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[4].Range.Text = "% качества";
                t.Rows[1].Cells[4].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[4].Range.Font.Size = 14;
                t.Rows[1].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[5].Range.Text = "Средний балл";
                t.Rows[1].Cells[5].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[5].Range.Font.Size = 14;
                t.Rows[1].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                ///

                int rows = kp + 1;
                for (int i = 0; i < rows - 1; i++)
                {
                    for (int js = 0; js < kp; js++)
                    {
                        doc.Tables[1].Rows[i + 2].Range.Text = predmet[i];
                        doc.Tables[1].Rows[i + 2].Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Rows[i + 2].Range.Font.Size = 14;
                        doc.Tables[1].Rows[i + 2].Range.Font.Bold = 0;
                        doc.Tables[1].Rows[i + 2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        ConnectDB.ConnectOnDb();
                        SqlCeCommand thisCommand21 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand21.CommandText = "SELECT gruppa.gruppa FROM gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN predmet ON kurs.id_kurs = predmet.id_kurs INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost AND predmet.id_specialnost = specialnost.id_specialnost CROSS JOIN teacher WHERE(predmet.predmet = '" + predmet[i] + "') AND(teacher.id_teacher = '" + idT + "')";
                        string gr = (string)thisCommand21.ExecuteScalar();

                        doc.Tables[1].Cell(i + 2, 2).Range.Text = gr;
                        doc.Tables[1].Cell(i + 2, 2).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 2).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 2).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        /// 

                        SqlCeCommand thisCommand31 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand31.CommandText = "SELECT COUNT(gurnal.ozenka) AS countOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1001') AND (gurnal.ozenka = '3' OR gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (predmet.predmet = '" + predmet[i] + "') AND (gurnal.id_teacher = '" + idT + "') AND (gurnal.date BETWEEN CONVERT(DATETIME, '" + d1 + "', 102) AND CONVERT(DATETIME, '" + d2 + "', 102))";
                        int oz = (int)thisCommand31.ExecuteScalar();

                        //SELECT COUNT(gurnal.ozenka) AS countOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1000') AND (gurnal.ozenka = '3' OR gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (predmet.predmet = '" + predmet[i] + "') AND (gurnal.id_teacher = '" + idT + "')
                        SqlCeCommand thisCommand41 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand41.CommandText = "SELECT COUNT(student.id_zachetki) AS CountST FROM kurs INNER JOIN predmet ON kurs.id_kurs = predmet.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs AND specialnost.id_specialnost = gruppa.id_specialnost INNER JOIN student ON gruppa.id_gruppa = student.id_gruppa CROSS JOIN teacher WHERE(predmet.predmet = '" + predmet[i] + "')";
                        int st = (int)thisCommand41.ExecuteScalar();

                        double rez = (oz * 100) / st;

                        doc.Tables[1].Cell(i + 2, 3).Range.Text = Convert.ToString(rez);
                        doc.Tables[1].Cell(i + 2, 3).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 3).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 3).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        SqlCeCommand thisCommand51 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand51.CommandText = "SELECT COUNT(gurnal.ozenka) AS countOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1001') AND (gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (gurnal.id_teacher = '" + idT + "') AND (predmet.predmet = '" + predmet[i] + "') AND (gurnal.date BETWEEN CONVERT(DATETIME, '" + d1 + "', 102) AND CONVERT(DATETIME, '" + d2 + "', 102))";
                        int oz2 = (int)thisCommand51.ExecuteScalar();

                        double rez2 = (oz2 * 100) / st;

                        doc.Tables[1].Cell(i + 2, 4).Range.Text = Convert.ToString(rez2);
                        doc.Tables[1].Cell(i + 2, 4).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 4).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 4).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        SqlCeCommand thisCommand71 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand71.CommandText = "SELECT COUNT(student.id_zachetki) AS Expr1 FROM gruppa INNER JOIN student ON gruppa.id_gruppa = student.id_gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN predmet ON kurs.id_kurs = predmet.id_kurs INNER JOIN pr_pr ON predmet.id_predmet = pr_pr.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher WHERE (predmet.predmet = '" + predmet[i] + "') AND (pr_pr.id_teacher = '" + idT + "')";
                        int st2 = (int)thisCommand71.ExecuteScalar();

                        string znPR = predmet[i];
                        SqlCeCommand thisCommand61 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand61.CommandText = "SELECT SUM(gurnal.ozenka) AS SOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1001') AND (gurnal.id_teacher = '" + idT + "') AND (predmet.predmet = '" + znPR + "') AND (gurnal.date BETWEEN CONVERT(DATETIME, '" + d1 + "', 102) AND CONVERT(DATETIME, '" + d2 + "', 102))";

                        int ozs;

                        if (thisCommand61.ExecuteScalar() == DBNull.Value)
                            ozs = 0;
                        else
                            ozs = Convert.ToInt32(thisCommand61.ExecuteScalar());

                        double srd = (double)ozs / (double)st2;
                        srd = Math.Round(srd, 2);



                        doc.Tables[1].Cell(i + 2, 5).Range.Text = Convert.ToString(srd);
                        doc.Tables[1].Cell(i + 2, 5).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 5).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 5).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        sru += rez;
                        srk += rez2;
                        srsr += srd;
                    }
                }
                sru = Math.Round(sru / (kp2 * kp2), 2);
                srk = Math.Round(srk / (kp2 * kp2), 2);
                srsr = Math.Round(srsr / (kp2 * kp2), 2);
                doc.Bookmarks["pr5"].Range.Text = Convert.ToString(sru);
                doc.Bookmarks["pr6"].Range.Text = Convert.ToString(srk);
                doc.Bookmarks["pr7"].Range.Text = Convert.ToString(srsr);
                doc.SaveAs(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocItog.docx");
                app.Documents.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocItog.docx");
            }
            else
            {
                zn = "1";
                g = data.Substring(6, 4);
                g2 = Convert.ToString(Convert.ToInt32(g) + 1);
                doc.Bookmarks["pr"].Range.Text = zn;
                doc.Bookmarks["pr2"].Range.Text = g + '-' + g2;
                doc.Bookmarks["pr3"].Range.Text = res;                

                Table t = doc.Tables.Add(doc.Bookmarks["pr4"].Range, kp + 1, 5);
                t.Borders.Enable = 1;

                foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
                {
                    table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
                }

                t.Rows[1].Cells[1].Range.Text = "Дисциплина";
                t.Rows[1].Cells[1].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[1].Range.Font.Size = 14;
                t.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[2].Range.Text = "Группа";
                t.Rows[1].Cells[2].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[2].Range.Font.Size = 14;
                t.Rows[1].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[3].Range.Text = "% успев.";
                t.Rows[1].Cells[3].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[3].Range.Font.Size = 14;
                t.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[4].Range.Text = "% качества";
                t.Rows[1].Cells[4].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[4].Range.Font.Size = 14;
                t.Rows[1].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                t.Rows[1].Cells[5].Range.Text = "Средний балл";
                t.Rows[1].Cells[5].Range.Font.Name = "Times New Roman";
                t.Rows[1].Cells[5].Range.Font.Size = 14;
                t.Rows[1].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                ///

                int rows = kp + 1;
                for (int i = 0; i < rows - 1; i++)
                {
                    for (int js = 0; js < kp; js++)
                    {
                        doc.Tables[1].Rows[i + 2].Range.Text = predmet[i];
                        doc.Tables[1].Rows[i + 2].Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Rows[i + 2].Range.Font.Size = 14;
                        doc.Tables[1].Rows[i + 2].Range.Font.Bold = 0;
                        doc.Tables[1].Rows[i + 2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        ConnectDB.ConnectOnDb();
                        SqlCeCommand thisCommand21 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand21.CommandText = "SELECT gruppa.gruppa FROM gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN predmet ON kurs.id_kurs = predmet.id_kurs INNER JOIN specialnost ON gruppa.id_specialnost = specialnost.id_specialnost AND predmet.id_specialnost = specialnost.id_specialnost CROSS JOIN teacher WHERE(predmet.predmet = '" + predmet[i] + "') AND(teacher.id_teacher = '" + idT + "')";
                        string gr = (string)thisCommand21.ExecuteScalar();

                        doc.Tables[1].Cell(i + 2, 2).Range.Text = gr;
                        doc.Tables[1].Cell(i + 2, 2).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 2).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 2).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        /// 

                        SqlCeCommand thisCommand31 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand31.CommandText = "SELECT COUNT(gurnal.ozenka) AS countOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1001') AND (gurnal.ozenka = '3' OR gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (predmet.predmet = '" + predmet[i] + "') AND (gurnal.id_teacher = '" + idT + "') AND (gurnal.date BETWEEN CONVERT(DATETIME, '" + d3 + "', 102) AND CONVERT(DATETIME, '" + d4 + "', 102))";
                        int oz = (int)thisCommand31.ExecuteScalar();

                        //SELECT COUNT(gurnal.ozenka) AS countOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1000') AND (gurnal.ozenka = '3' OR gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (predmet.predmet = '" + predmet[i] + "') AND (gurnal.id_teacher = '" + idT + "')
                        SqlCeCommand thisCommand41 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand41.CommandText = "SELECT COUNT(student.id_zachetki) AS CountST FROM kurs INNER JOIN predmet ON kurs.id_kurs = predmet.id_kurs INNER JOIN specialnost ON predmet.id_specialnost = specialnost.id_specialnost INNER JOIN gruppa ON kurs.id_kurs = gruppa.id_kurs AND specialnost.id_specialnost = gruppa.id_specialnost INNER JOIN student ON gruppa.id_gruppa = student.id_gruppa CROSS JOIN teacher WHERE(predmet.predmet = '" + predmet[i] + "') AND (teacher.id_teacher = '" + idT + "')";
                        int st = (int)thisCommand41.ExecuteScalar();

                        double rez = (oz * 100) / st;

                        doc.Tables[1].Cell(i + 2, 3).Range.Text = Convert.ToString(rez);
                        doc.Tables[1].Cell(i + 2, 3).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 3).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 3).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        SqlCeCommand thisCommand51 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand51.CommandText = "SELECT COUNT(gurnal.ozenka) AS countOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1001') AND (gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (gurnal.id_teacher = '" + idT + "') AND (predmet.predmet = '" + predmet[i] + "')  AND (teacher.id_teacher = '" + idT + "') AND (gurnal.date BETWEEN CONVERT(DATETIME, '" + d3 + "', 102) AND CONVERT(DATETIME, '" + d4 + "', 102))";
                        int oz2 = (int)thisCommand51.ExecuteScalar();

                        double rez2 = (oz2 * 100) / st;

                        doc.Tables[1].Cell(i + 2, 4).Range.Text = Convert.ToString(rez2);
                        doc.Tables[1].Cell(i + 2, 4).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 4).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 4).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        SqlCeCommand thisCommand71 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand71.CommandText = "SELECT COUNT(student.id_zachetki) AS Expr1 FROM gruppa INNER JOIN student ON gruppa.id_gruppa = student.id_gruppa INNER JOIN kurs ON gruppa.id_kurs = kurs.id_kurs INNER JOIN predmet ON kurs.id_kurs = predmet.id_kurs INNER JOIN pr_pr ON predmet.id_predmet = pr_pr.id_predmet INNER JOIN teacher ON pr_pr.id_teacher = teacher.id_teacher WHERE (predmet.predmet = '" + predmet[i] + "') AND (pr_pr.id_teacher = '" + idT + "')";
                        int st2 = (int)thisCommand71.ExecuteScalar();

                        string znPR = predmet[i];
                        SqlCeCommand thisCommand61 = ConnectDB.ConnectOnDb().CreateCommand();
                        thisCommand61.CommandText = "SELECT SUM(gurnal.ozenka) AS SOZ FROM gurnal INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet INNER JOIN vidOzenki ON gurnal.id_vidOzenki = vidOzenki.id_vidOzenki INNER JOIN teacher ON gurnal.id_teacher = teacher.id_teacher WHERE (gurnal.id_vidOzenki = '1001') AND (gurnal.id_teacher = '" + idT + "') AND (predmet.predmet = '" + znPR + "') AND (gurnal.date BETWEEN CONVERT(DATETIME, '" + d3 + "', 102) AND CONVERT(DATETIME, '" + d4 + "', 102))";

                        int ozs;

                        if (thisCommand61.ExecuteScalar() == DBNull.Value)
                            ozs = 0;
                        else
                            ozs = Convert.ToInt32(thisCommand61.ExecuteScalar());

                        double srd = (double)ozs / (double)st2;
                        srd = Math.Round(srd, 2);



                        doc.Tables[1].Cell(i + 2, 5).Range.Text = Convert.ToString(srd);
                        doc.Tables[1].Cell(i + 2, 5).Range.Font.Name = "Times New Roman";
                        doc.Tables[1].Cell(i + 2, 5).Range.Font.Size = 10;
                        doc.Tables[1].Cell(i + 2, 5).Range.Font.Bold = 0;
                        doc.Tables[1].Cell(i + 2, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        sru += rez;
                        srk += rez2;
                        srsr += srd;
                    }
                }
                sru = Math.Round(sru / (kp2 * kp2), 2);
                srk = Math.Round(srk / (kp2 * kp2), 2);
                srsr = Math.Round(srsr / (kp2 * kp2), 2);
                doc.Bookmarks["pr5"].Range.Text = Convert.ToString(sru);
                doc.Bookmarks["pr6"].Range.Text = Convert.ToString(srk);
                doc.Bookmarks["pr7"].Range.Text = Convert.ToString(srsr);
                doc.SaveAs(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocItog.docx");
                app.Documents.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocItog.docx");
            }
            
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            int ip = 0;
            int rows = ks + 1;
            int columns = 7;
            ///  

            System.Data.DataTable dt = new System.Data.DataTable();

            SqlCeCommand thisCommandKS1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommandKS1.CommandText = "SELECT COUNT (student.id_zachetki) AS KS1 FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa INNER JOIN gurnal ON student.id_zachetki = gurnal.id_zachetki INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet WHERE (gruppa.gruppa = '" + comboBox1.Text + "') AND (gurnal.id_vidOzenki = '1002') AND (gurnal.ozenka = '3' OR gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (gurnal.id_predmet = '" + x + "') AND (gurnal.id_teacher = '" + label1.Text + "')";
            thisCommandKS1.Connection = ConnectDB.ConnectOnDb();
            int ksid1 = (int)thisCommandKS1.ExecuteScalar();

            SqlCeCommand thisCommandidM1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommandidM1.CommandText = "SELECT student.id_zachetki FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa INNER JOIN gurnal ON student.id_zachetki = gurnal.id_zachetki INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet WHERE (gruppa.gruppa = '" + comboBox1.Text + "') AND (gurnal.id_vidOzenki = '1002') AND (gurnal.ozenka = '3' OR gurnal.ozenka = '4' OR gurnal.ozenka = '5') AND (gurnal.id_predmet = '" + x + "') AND (gurnal.id_teacher = '" + label1.Text + "')";
            SqlCeDataReader thisReaderidM1 = thisCommandidM1.ExecuteReader();
            int[] idM1 = new int[ksid1];
            int k1 = 0;
            while (thisReaderidM1.Read())
            {
                idM1[k1] = Convert.ToInt32(thisReaderidM1["id_zachetki"]);
                k1++;
            }
            thisReaderidM1.Close();


            SqlCeCommand thisCommandidM2 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommandidM2.CommandText = "SELECT DISTINCT student.id_zachetki FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa INNER JOIN gurnal ON student.id_zachetki = student.id_zachetki INNER JOIN predmet ON gurnal.id_predmet = predmet.id_predmet WHERE (gruppa.gruppa = '" + comboBox1.Text + "') AND (gurnal.id_predmet = '" + x + "') AND (gurnal.id_teacher = '" + label1.Text + "')";
            SqlCeDataReader thisReaderidM2 = thisCommandidM2.ExecuteReader();
            int[] idM2 = new int[ks];
            int k2 = 0;
            while (thisReaderidM2.Read())
            {
                idM2[k2] = Convert.ToInt32(thisReaderidM2["id_zachetki"]);
                k2++;
            }
            thisReaderidM2.Close();
            int[] idM3 = new int[ks];
            idM3 = Except(idM2, idM1);

            int j = 0;
            string[] fam = new string[idM3.Length];
            string[] im = new string[idM3.Length];
            string[] ot = new string[idM3.Length];
            string[] iz = new string[idM3.Length];
            for ( ip = 0; ip < idM3.Length; ip++)
            {

                SqlCeCommand thisCommand = ConnectDB.ConnectOnDb().CreateCommand();
                thisCommand.CommandText = "SELECT student.id_zachetki, student.Sname, student.Fname, student.Lname FROM student INNER JOIN gruppa ON student.id_gruppa = gruppa.id_gruppa WHERE (gruppa.gruppa = '" + comboBox1.Text + "') AND (student.id_zachetki = '" + idM3[ip] + "')";
                SqlCeDataReader thisReader = thisCommand.ExecuteReader();
               

                while (thisReader.Read())
                {
                    fam[j] = Convert.ToString(thisReader["Sname"]);
                    im[j] = Convert.ToString(thisReader["Fname"]);
                    ot[j] = Convert.ToString(thisReader["Lname"]);
                    iz[j] = Convert.ToString(thisReader["id_zachetki"]);
                    j++;
                }               
            }


            System.Data.DataTable dt1 = new System.Data.DataTable();

            SqlCeCommand thisCommand1 = ConnectDB.ConnectOnDb().CreateCommand();
            thisCommand1.CommandText = "SELECT Sname, Fname, Lname  FROM teacher WHERE id_teacher = '" + label1.Text + "'";
            SqlCeDataReader thisReader1 = thisCommand1.ExecuteReader();
            string res1 = string.Empty;
            while (thisReader1.Read())
            {
                res1 += thisReader1["Sname"];
                res1 += " ";
                res1 += thisReader1["Fname"];
                res1 += " ";
                res1 += thisReader1["Lname"];
            }
            thisReader1.Close();

            FIOT = res1;
            string d = DateTime.Now.ToString("dd.MM.yyyy");

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Add(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\Vedomost.docx");

            doc.Bookmarks["pr1"].Range.Text = comboBox2.Text;
            doc.Bookmarks["pr2"].Range.Text = "                                                                                    " + d + "                                                                         " + comboBox1.Text;
            doc.Bookmarks["pr3"].Range.Text = FIOT;

            Table t = doc.Tables.Add(doc.Bookmarks["pr4"].Range, idM3.Length + 1, 7);
            t.Borders.Enable = 1;

            foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
            {
                table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
            }

            t.Rows[1].Cells[1].Range.Text = "№ п/п";
            t.Rows[1].Cells[1].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[1].Range.Font.Size = 10;
            t.Rows[1].Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[2].Range.Text = "№ экзамен листа";
            t.Rows[1].Cells[2].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[2].Range.Font.Size = 10;
            t.Rows[1].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[3].Range.Text = "Фамилия";
            t.Rows[1].Cells[3].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[3].Range.Font.Size = 10;
            t.Rows[1].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[4].Range.Text = "Имя";
            t.Rows[1].Cells[4].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[4].Range.Font.Size = 10;
            t.Rows[1].Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[5].Range.Text = "Отчество";
            t.Rows[1].Cells[5].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[5].Range.Font.Size = 10;
            t.Rows[1].Cells[5].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[6].Range.Text = "Отметка";
            t.Rows[1].Cells[6].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[6].Range.Font.Size = 10;
            t.Rows[1].Cells[6].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            t.Rows[1].Cells[7].Range.Text = "Подпись экзаменатора";
            t.Rows[1].Cells[7].Range.Font.Name = "Times New Roman";
            t.Rows[1].Cells[7].Range.Font.Size = 10;
            t.Rows[1].Cells[7].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
            {
                table.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow);
            }            

            for (int i = 0; i < idM3.Length; i++)
            {
                for (int js = 0; js < ks; js++)
                {
                    doc.Tables[1].Rows[i + 2].Range.Text = Convert.ToString(i + 1);
                    doc.Tables[1].Rows[i + 2].Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Rows[i + 2].Range.Font.Size = 10;
                    doc.Tables[1].Rows[i + 2].Range.Font.Bold = 0;
                    doc.Tables[1].Rows[i + 2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    doc.Tables[1].Cell(i + 2, 3).Range.Text = fam[i];
                    doc.Tables[1].Cell(i + 2, 3).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 3).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 3).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 3).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    doc.Tables[1].Cell(i + 2, 4).Range.Text = im[i];
                    doc.Tables[1].Cell(i + 2, 4).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 4).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 4).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 4).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                    doc.Tables[1].Cell(i + 2, 5).Range.Text = ot[i];
                    doc.Tables[1].Cell(i + 2, 5).Range.Font.Name = "Times New Roman";
                    doc.Tables[1].Cell(i + 2, 5).Range.Font.Size = 10;
                    doc.Tables[1].Cell(i + 2, 5).Range.Font.Bold = 0;
                    doc.Tables[1].Cell(i + 2, 5).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }

            doc.SaveAs(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocPrintOut2.docx");
            app.Documents.Open(Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath) + @"\\DocPrintOut2.docx");
        }
                    
       

        public static int[] Except(int[] first, int[] second)
        {
            var intermediateArray = new int[first.Length];
            int resultCount = 0;
            for (int index = 0; index < first.Length; index++)
            {
                bool unique = true;
                for (int secondIndex = 0; secondIndex < second.Length; secondIndex++)
                {
                    if (first[index] == second[secondIndex])
                    {
                        unique = false;
                        break;
                    }
                }
                if (unique)
                {
                    resultCount++;
                    intermediateArray[resultCount - 1] = first[index];
                }
            }
            var outArray = new int[resultCount];
            Array.Copy(intermediateArray, 0, outArray, 0, resultCount);
            return outArray;
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
