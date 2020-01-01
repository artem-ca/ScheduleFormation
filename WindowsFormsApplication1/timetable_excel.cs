using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Регулярки
using System.Text.RegularExpressions;

// Excel 
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication
{

    class ExcelTimeTable
    {

        private List<string> FileNames; 

        // Конструктор для получения путей Excel файлов

        public ExcelTimeTable(string[] f)
        {
            this.FileNames = f.ToList();
        }

        public ExcelTimeTable(List<string> f)
        {
            this.FileNames = f;
        }

        public void Remove( string fileName )
        {
            this.FileNames.Remove(fileName);
        }

        public void RemoveAt(int id)
        {
            this.FileNames.RemoveAt(id);
        }

        public void Clear()
        {
            this.FileNames.Clear();
        }

        public void SetFileNames(List<string> fileNames)
        {
            this.FileNames = fileNames;
        }

        public void SetFileNames(string[] fileNames)
        {
            this.FileNames = fileNames.ToList();
        }

        public void Close()
        {
            obj.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);

            obj         = null;
            objbook     = null;
            workSheet   = null;

            // Освобождаем память (Сборка мусора)
            GC.Collect();
            // GC.WaitForPendingFinalizers();
        }

        // Инициализация Excel фрагментов

        private Excel.Application obj = new Excel.Application();
        private Excel.Workbook objbook;
        private Excel.Sheets workSheet;

        // Константные значения строк для работы с расписанием всех групп 

        const int start_row = 9;
        const int end_row   = 68;

        // Дни недели 
        public Dictionary<int, string> weeks = new Dictionary<int, string>()
        {
            { 1, "Понедельник" },
            { 2, "Вторник" },
            { 3, "Среда" },
            { 4, "Четверг" },
            { 5, "Пятница" },
            { 6, "Суббота" }
        };


        // Структура предмета

        public struct Subject
        {
            public string pair_name;
            public string room;
            public string group;
        }

        // Расписание преподавателей 

        private Dictionary<string, Subject?[,]> timetable = new Dictionary<string, Subject?[,]>();

        // Список данных 

        private List<string>[] GetData()
        {

            List<string>[] data = new List<string>[4];

            // Объявление экземпляров для массива листов
            for ( int i = 0; i < 4; i ++ )
            {
                data[i] = new List<string>();
            }

            try
            {

                // Перебор всех файлов
                for (int k = 0; k < FileNames.Count(); k++)
                {

                    objbook = obj.Workbooks.Open(this.FileNames[k], 0, true);
                 
                    workSheet = (Excel.Sheets)objbook.Sheets;

                    // Перебор листов в Excel
                    for (int j = 1; j <= (int)objbook.Sheets.Count; j++)
                    {
                        // Перебор строк
                        for (int i = start_row; i <= end_row; i++)
                        {

                            // Вытаскиваем преподавателей
                            string teachers1 = workSheet[j].Cells[i, "D"].Text.Trim();
                            string teachers2 = workSheet[j].Cells[i, "H"].Text.Trim();
                            // Вытаскиваем предметы
                            string subjects1 = workSheet[j].Cells[i, "C"].Text.Trim();
                            string subjects2 = workSheet[j].Cells[i, "G"].Text.Trim();
                            // Вытаскиваем аудитории
                            string rooms1 = workSheet[j].Cells[i, "E"].Text.Trim();
                            string rooms2 = workSheet[j].Cells[i, "I"].Text.Trim();
                            // Вытаскиваем группы
                            string groups1 = workSheet[j].Cells[6, "C"].Text.Trim();
                            string groups2 = workSheet[j].Cells[6, "G"].Text.Trim();

                            // Заполнение преподавателей
                            data[0].Add(teachers1);
                            data[0].Add(teachers2);

                            // Заполнение предметов
                            data[1].Add(subjects1);
                            data[1].Add(subjects2);

                            // Заполнение аудиторий
                            data[2].Add(rooms1);
                            data[2].Add(rooms2);

                            // Заполнение групп
                            data[3].Add(groups1);
                            data[3].Add(groups2);
                        }
                    }

                    objbook.Close();

                }

            }
            catch
            {
                // Если возникает ошибка, то метод GetData прекращает считывать данные, а также возвращает значение 
            }

            return data; 

        }

        public Dictionary<string, Subject?[,]> ListTeachers(bool IsNumerator)
        {

            // IsNumerator - Проверка на числитель или знаменатель ( числитель = true | знаменатель = false )

            // Формирование структуры

            timetable.Clear();

            List<string>[] data = GetData();

            List<string> teachers   = data[0];        
            List<string> pairs      = data[1];
            List<string> rooms      = data[2];
            List<string> groups     = data[3];

            int num_pair = 0;
            int pairs_per_week = 30;

            for (int i = 0; i < pairs.Count; i++)
            {

                // Получаем номер пары
                num_pair = i % 120 / 4;

                // Вытаскиваем предметы, в зависимости от выбора (числитель/знаменатель)
                if (
                    (
                        pairs[i].StartsWith(IsNumerator ? "/" : "\\")
                        ||
                        !pairs[i].StartsWith(!IsNumerator ? "/" : "\\")
                    )
                    &&
                        teachers[i].Trim() != ""
                   )
                {

                    // Пробегаем недели, соответствующую паре
                    int week = num_pair / 5;
                    // Пробегаем номера пар 
                    int cur_numpair = num_pair % 5;
                    // Пробегаем преподавателей 
                    string cur_teacher = teachers[i];
                    // Пробегаем пары
                    string cur_pair = pairs[i];
                    // Пробегаем  аудитории
                    string cur_room = rooms[i] != "" ? rooms[i] : "***";
                    // Пробегаем группы
                    string cur_group = groups[i];

                    // Если у предмета нет преподавателя, то пропускаем итерацию 
                    if (cur_teacher == "")
                        continue;

                    // Проверка делятся ли студенты на подгруппы
                    if (!cur_teacher.Contains("/"))
                    {

                        // Заполняем словарь 
                        if (!timetable.ContainsKey(cur_teacher))
                            timetable[cur_teacher] = new Subject?[6, 5];

                        Subject _t = new Subject { pair_name = cur_pair, room = cur_room, group = cur_group };
                        timetable[cur_teacher][week, cur_numpair] = _t;

                    }
                    else
                    {

                        // Если студенты делятся на подгруппы, то 

                        // Вытаскиваем предметы через регулярки
                        MatchCollection matches = Regex.Matches(cur_pair, @"^(.*)\s\/(.*п\/г)\s\/(.*п\/г)$");

                        // Если не 3 занятия в одной паре, значит ищем для двух занятий
                        if (matches.Count == 0)
                            matches = Regex.Matches(cur_pair, @"^(.*)\s\/(.*п\/г)$");

                        // Вытаскиваем преподов 
                        string[] t_teachers = teachers[i].Split('/');

                        // Вытаскиваем аудитории
                        string[] t_rooms = new string[t_teachers.Count()];

                        if (cur_room.Contains('/'))
                            t_rooms = rooms[i].Split('/');
                        else
                            t_rooms[0] = cur_room;

                        for (int j = 0; j < t_teachers.Count(); j++)
                        {

                            cur_teacher = t_teachers[j];
                            cur_room = t_rooms[j] != null ? t_rooms[j] : "***";

                            // Если регулярное выражение удовлетворило условию, то вытаскиваем предметы
                            if (matches.Count > 0)
                            {
                                // к.п - каждый предмет
                                // Пример 1: /Ин.яз 1 п/г + Информ. 2 п/г [ => (Преобразует к.п к числителю) ]    /Ин.яз 1 п/г + /Информ. 2 п/г
                                // Пример 2: \Ин.яз 2 п/г + Информ. 1 п/г [ => (Преобразует к.п к знаменателю) ]  \Ин.яз 2 п/г + /Информ. 1 п/г
                                // Пример 3: Информ. 2 п/г + Инж.гр. 1 п/г [ => ]  Информ. 2 п/г + Инж.гр. 1 п/г

                                cur_pair = ( pairs[i].StartsWith("/") ? "/" : pairs[i].StartsWith("\\")  ? "\\" : "" ) 
                                    + matches[0].Groups[j + 1].Value;
                            }

                            if (!timetable.ContainsKey(cur_teacher))
                                timetable[cur_teacher] = new Subject?[6, 5];

                            Subject _t = new Subject { pair_name = cur_pair, room = cur_room, group = cur_group };
                            timetable[cur_teacher][week, cur_numpair] = _t;

                        }

                    }

                }

                num_pair = num_pair % pairs_per_week;

            }

            // Сортировка словаря (по фамилии преподавателя или же ключу словаря)
            timetable = timetable.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);

            return timetable; 

        }

        // Перегрузка для извлечения всех параметров
        public Dictionary<string, Subject?[,]> ListTeachers()
        {

            // Формирование структуры

            timetable.Clear();

            List<string>[] data = GetData();

            List<string> teachers   = data[0];
            List<string> pairs      = data[1];
            List<string> rooms      = data[2];
            List<string> groups     = data[3];

            int num_pair = 0;
            int pairs_per_week = 30;

            for (int i = 0; i < pairs.Count; i++)
            {

                // Получаем номер пары
                num_pair = i % 120 / 4;

                if (
                        teachers[i].Trim() != ""
                   )
                {

                    // Пробегаем недели, соответствующую паре
                    int week = num_pair / 5;
                    // Пробегаем номера пар 
                    int cur_numpair = num_pair % 5;
                    // Пробегаем преподавателей 
                    string cur_teacher = teachers[i];
                    // Пробегаем пары
                    string cur_pair = pairs[i];
                    // Пробегаем  аудитории
                    string cur_room = rooms[i] != "" ? rooms[i] : "***";
                    // Пробегаем группы
                    string cur_group = groups[i];

                    // Если у предмета нет преподавателя, то пропускаем итерацию 
                    if (cur_teacher == "")
                        continue;

                    // Проверка не делятся ли студенты на подгруппы
                    if (!cur_teacher.Contains("/"))
                    {

                        // Заполняем словарь 
                        if (!timetable.ContainsKey(cur_teacher))
                            timetable[cur_teacher] = new Subject?[6, 5];

                        Subject _t;

                        if (timetable[cur_teacher][week, cur_numpair] != null)
                        {
                            _t = new Subject
                            {
                                pair_name = timetable[cur_teacher][week, cur_numpair].Value.pair_name + "+" + cur_pair,
                                room = cur_room,
                                group = timetable[cur_teacher][week, cur_numpair].Value.group + "+" + cur_group
                            };
                        }
                        else
                        {
                            _t = new Subject { pair_name = cur_pair, room = cur_room, group = cur_group };
                        }

                        timetable[cur_teacher][week, cur_numpair] = _t;

                    }
                    else
                    {

                        // Если студенты делятся на подгруппы, то 

                        // Вытаскиваем предметы через регулярки
                        MatchCollection matches = Regex.Matches(cur_pair, @"^(.*)\s\/(.*п\/г)\s\/(.*п\/г)$");

                        // Если не 3 занятия в одной паре, значит ищем для двух занятий
                        if (matches.Count == 0)
                            matches = Regex.Matches(cur_pair, @"^(.*)\s\/(.*п\/г)$");

                        // Вытаскиваем преподов 
                        string[] t_teachers = teachers[i].Split('/');

                        // Вытаскиваем аудитории

                        string[] t_rooms = new string[t_teachers.Count()];

                        if (cur_room.Contains('/'))
                            t_rooms = rooms[i].Split('/');
                        else
                            t_rooms[0] = cur_room;

                        for (int j = 0; j < t_teachers.Count(); j++)
                        {

                            cur_teacher = t_teachers[j];
                            cur_room = t_rooms[j] != null ? t_rooms[j] : "***";

                            // Если регулярное выражение удовлетворило условию, то вытаскиваем предметы
                            if (matches.Count > 0)
                                cur_pair = matches[0].Groups[j + 1].Value;

                            if (!timetable.ContainsKey(cur_teacher))
                                timetable[cur_teacher] = new Subject?[6, 5];

                            // Subject _t = new Subject { pair_name = cur_pair, room = cur_room, group = cur_group };

                            Subject _t;

                            if (timetable[cur_teacher][week, cur_numpair] != null)
                            {
                                _t = new Subject
                                {
                                    pair_name = timetable[cur_teacher][week, cur_numpair].Value.pair_name + "+" + cur_pair,
                                    room = cur_room,
                                    group = timetable[cur_teacher][week, cur_numpair].Value.group + "+" + cur_group
                                };
                            }
                            else
                            {
                                _t = new Subject { pair_name = cur_pair, room = cur_room, group = cur_group };
                            }

                            timetable[cur_teacher][week, cur_numpair] = _t;

                        }

                    }

                }

                num_pair = num_pair % pairs_per_week;

            }

            // Сортировка словаря (по фамилии преподавателя или же ключу словаря)
            timetable = timetable.OrderBy(x => x.Key).ToDictionary(x => x.Key, x => x.Value);

            return timetable;

        }

    }
}
