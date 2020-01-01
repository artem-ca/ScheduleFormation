using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
// Кастомные библиотеки
using System.Threading;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace WindowsFormsApplication
{ 

    public partial class main : Form
    {

        // Объявление потока
        Thread th;
        // Объявление нашего класса
        ExcelTimeTable t = null;
        // Список преподавателей
        Dictionary<string, ExcelTimeTable.Subject?[,]> timetable = new Dictionary<string, ExcelTimeTable.Subject?[,]>();
        // Список путей к расписаниям курсов
        List<string> fileNames = new List<string>();
        // Определяет, идет-ли загрузка в данный момент (по нажатию кнопки - истина, по окончанию загрузки - ложь)
        bool works = false;
        // Определяет, корректен ли путь к шаблону в конфиге
        bool IsValid = true;
        // Путь к файлу, в который будет загружаться расписание
        string oldLink = "";
        // Новый путь к файлу, если выбрана заргузка в новый файл
        string newLink = "";



        public main()
        {
            InitializeComponent();
        }

        private void main_Load(object sender, EventArgs e)
        {

            if (File.Exists("config.txt"))
            {
                oldLink = File.ReadAllText("config.txt").Trim();     
            }

            if (oldLink != "")
            {
                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;
                открытьToolStripMenuItem.Enabled = true;
            }

        }

        public Excel.Application ScheduleFormation()
        {

            try
            {

                //Перебор преподавателей (Заполнение Excel) 

                Excel.Application app = new Excel.Application();
                Excel.Workbook template = app.Workbooks.Open(oldLink);
                Excel.Worksheet tmp_sheet = template.Sheets[1];

                int _tindex = 1;
                int _i = 1;

                string p1 = "";
                string p2 = "";

                MatchCollection m;

                foreach (var el in timetable)
                {

                    tmp_sheet.Cells[_tindex, "A"] = _i;
                    tmp_sheet.Cells[_tindex, "B"] = el.Key;

                    for (int i = 0; i < 6; i++)
                    {

                        for (int j = 0; j < 5; j++)
                        {

                            if (el.Value[i, j] == null) continue;

                            string groups = el.Value[i, j].Value.group;
                            string[] spl = groups.Split('+');
                            string[] spl2 = el.Value[i, j].Value.pair_name.Split('+');

                            switch (spl.Count())
                            {
                                case 2:

                                    // Если одна группа делится на подгруппы
                                    if (spl[0].Equals(spl[1]))
                                    {
                                        // Вытаскиваем подгруппы
                                        // Пример: Англ 1 п/г+Математика 2 п/г
                                        m = Regex.Matches(el.Value[i, j].Value.pair_name, @".*(\d).*\+.*(\d).*");

                                        if (m.Count > 0)
                                        {

                                            p1 = m[0].Groups[1].Value;
                                            p2 = m[0].Groups[2].Value;

                                            tmp_sheet.Cells[_tindex + 1 + j, i + 4] = $"{spl[0]}({p1}) / ({p2})";
                                        } // Если 1 подгруппа, то ищем для остальных случаев
                                        else
                                        {
                                            // Пример: Англ 1 п/г+Математика
                                            m = Regex.Matches(el.Value[i, j].Value.pair_name, @".*(\d).*\+.*");

                                            if (m.Count > 0)
                                            {
                                                p1 = m[0].Groups[1].Value;
                                                tmp_sheet.Cells[_tindex + 1 + j, i + 4] = $"{spl[0]}({p1}) / {spl[0]}";
                                            }
                                            else
                                            {
                                                // Пример: Англ+Математика 2 п/г
                                                m = Regex.Matches(el.Value[i, j].Value.pair_name, @".*\+.*(\d).*");

                                                if (m.Count > 0)
                                                {
                                                    p1 = m[0].Groups[1].Value;
                                                    tmp_sheet.Cells[_tindex + 1 + j, i + 4] = $"{spl[0]} / {spl[0]}({p1})";
                                                }

                                            }

                                        }

                                    }
                                    else // Блок с двумя разными группами 
                                    {

                                        if (spl2[0].StartsWith("/"))
                                            tmp_sheet.Cells[_tindex + 1 + j, i + 4] = $"{spl[0]} / {spl[1]}";
                                        else
                                            tmp_sheet.Cells[_tindex + 1 + j, i + 4] = $"{spl[1]} / {spl[0]}";
                                    }

                                    break;

                                case 1:

                                    m = Regex.Matches(el.Value[i, j].Value.pair_name, @"(\d)");
                                    string num_subgroup = m.Count > 0 ? $"({m[0].Groups[1].Value})" : "";

                                    // Пара по числителю 
                                    if (el.Value[i, j].Value.pair_name.StartsWith("/"))
                                    {
                                        tmp_sheet.Cells[_tindex + 1 + j, i + 4] = groups + num_subgroup + "/";
                                    }
                                    // Пара по знаменателю 
                                    else if (el.Value[i, j].Value.pair_name.StartsWith("\\"))
                                    {
                                        tmp_sheet.Cells[_tindex + 1 + j, i + 4] = "/" + groups + num_subgroup;
                                    }
                                    // Постоянная пара
                                    else
                                    {
                                        tmp_sheet.Cells[_tindex + 1 + j, i + 4] = groups + num_subgroup;
                                    }

                                    break;

                                default:

                                    tmp_sheet.Cells[_tindex + 1 + j, i + 4] = groups;

                                    break;
                            }

                            tmp_sheet.Cells[_tindex + 7, i + 4] = el.Value[i, j].Value.room;

                        }

                    }

                    // Переход на 8 строчек вниз для записи следующего расписания преподавателя                
                    _tindex += 8;
                    _i++;

                }

                return app;
            }
            catch
            {
                MessageBox.Show("Ошибка конфигурации: Некорректный путь к шаблону.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

                IsValid = false; 

                return null;
            }

        }

        public Excel.Application ClearSchedule()
        {

            Excel.Application app = new Excel.Application();
            Excel.Workbook template = app.Workbooks.Open(oldLink);
            Excel.Worksheet tmp_sheet = template.Sheets[1];

            int _tindex = 1;

            foreach (var el in timetable)
            {

                for (int i = 0; i < 6; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {

                        tmp_sheet.Cells[_tindex + 1 + j, i + 4] = "";
                        tmp_sheet.Cells[_tindex + 7, i + 4] = "";

                    }
                }

                _tindex += 8;

            }

            return app;

        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                if (works)
                {
                    toolStripStatusLabel.Text = "Пожалуйста подождите, пока загрузка будет завершена.";
                    return;
                }

                listBox1.Enabled = true;
                lblDelete.Enabled = true;
                lblClearSelect.Enabled = true;

                listBox1.Items.Clear();

                this.fileNames.AddRange(openFileDialog.FileNames);

                this.fileNames = this.fileNames.Distinct().ToList();

                listBox1.Items.AddRange(this.fileNames.ToArray());

                if (t != null)
                    t.Close();

                t = new ExcelTimeTable(this.fileNames);

                button1.Enabled = true;
                toolStripStatusLabel.Text = "Нажмите <Сформировать>";

            }
            
        }

        private void выбратьШаблонToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                if (works)
                {
                    toolStripStatusLabel.Text = "Пожалуйста подождите, пока загрузка будет завершена.";
                    return;
                }

                radioButton1.Enabled = true;
                radioButton2.Enabled = true;
                radioButton3.Enabled = true;

                File.WriteAllText("config.txt", openFileDialog.FileName);

                oldLink = File.ReadLines("config.txt").Skip(0).First();

                открытьToolStripMenuItem.Enabled = true;

                newLink = oldLink;
                int pos = newLink.LastIndexOf('\\');
                newLink = newLink.Substring(0, pos) + "\\Новое расписание.xls";

                listBox1.Items.Clear();

                button1.Enabled = false;
                toolStripStatusLabel.Text = "Выбран новый шаблон. Можете загрузить необходимые файлы.";

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (works)
            {
                toolStripStatusLabel.Text = "Пожалуйста подождите, пока загрузка будет завершена.";
                return;
            }

            works = true; 
            toolStripStatusLabel.Text = "Загрузка...";
            this.Cursor = Cursors.AppStarting;
            listBox1.Enabled = false;

            // Инициализация потока
            th = new Thread(this.potok);
            th.Start();

        }

        private void label3_Click(object sender, EventArgs e)
        {

            if (listBox1.SelectedIndices.Count != 0)
            {

                foreach (var item in listBox1.SelectedItems.Cast<string>().ToList())
                {

                    if (fileNames.Contains(item))
                        fileNames.Remove(item);

                    listBox1.Items.Remove(item);

                    if (listBox1.Items.Count == 0)
                        button1.Enabled = false;

                }
                
            }
            else if (listBox1.Items.Count == 0)
            {
                return;
            }
            else
            {
                listBox1.Items.Clear();
                this.fileNames.Clear();

                t.Clear();

                button1.Enabled = false;
            }

        }

        private void label4_Click(object sender, EventArgs e)
        {
            listBox1.ClearSelected();
        }

        private void potok()
        {

            // this.Invoke используюется для того, чтобы обращаться к полям и методам основного потока.
            
            // Получаем структуру данных

            //Если выбран (числитель и знаматель)
            if (radioButton3.Checked)
            {
                timetable = t.ListTeachers();
            }
            else
            { 
                // Если выбран числитель, то истина.
                bool IsNumerator = radioButton2.Checked && !radioButton1.Checked;
                timetable = t.ListTeachers(IsNumerator);
            }


            Excel.Application f = ScheduleFormation();  
            // Excel.Application f = ClearSchedule();

            // Остановка потока
            this.Invoke((MethodInvoker)delegate () 
            {

                th.Abort();
                toolStripStatusLabel.Text = IsValid ? "Завершено." : "Ошибка конфигурации. Выберите новый шаблон.";
                works = false;
                listBox1.Enabled = true;
                Cursor = Cursors.Default;

                if ( f != null )
                    f.Visible = true;

                IsValid = true;

            });

        }

        private void main_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (th != null && th.ThreadState != ThreadState.Aborted)
                th.Abort();

            if ( t != null )
                t.Close();
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();  
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox();
            about.Show();
        }


        // Перетаскивание файлов в listBox
        // Обеспечено перетаскивание только файлов формата xls или xlsx

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {

            listBox1.Items.Clear();

            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);

            for (int i = 0; i < s.Length; i++)
            {
                this.fileNames.Add(s[i]);
            }

            this.fileNames = this.fileNames.Distinct().ToList();

            listBox1.Items.AddRange(this.fileNames.ToArray());

            if (this.fileNames.Count > 0)
            {

                if (works)
                {
                    toolStripStatusLabel.Text = "Пожалуйста подождите, пока загрузка будет завершена.";
                    return;
                }

                if (t != null)
                    t.Close();

                t = new ExcelTimeTable(this.fileNames);

                button1.Enabled = true;
                toolStripStatusLabel.Text = "Нажмите <Сформировать>";
            }

        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && (e.Data.GetData(DataFormats.FileDrop, false) as string[]).All(x => (x.EndsWith(".xls") || x.EndsWith(".xlsx")) ) )
                e.Effect = DragDropEffects.Link;
            else
                e.Effect = DragDropEffects.None;
        }

        private void label3_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.ToolTipTitle = "Подсказка";
            toolTip1.Show("Удалить из листа", lblDelete);
        }

        private void lblClearSelect_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.ToolTipTitle = "Подсказка";
            toolTip1.Show("Снять выделение", lblClearSelect);
        }

        private void lblClearSelect_MouseLeave(object sender, EventArgs e)
        {
            toolTip1.Hide(lblClearSelect);
        }

        /*
        private void comboBox1_MouseEnter(object sender, EventArgs e)
        {
            toolTip1.ToolTipTitle = "Путь";

            if (comboBox1.SelectedIndex == 0)
                toolTip1.Show($"{oldLink}", comboBox1);
            else
                toolTip1.Show($"{newLink}", comboBox1);
        }
        */

    }
}