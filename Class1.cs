using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// Autocad
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;


namespace ClassLibrary2
{
    public class Class1
    {
        // значения вводимые по умолчанию. Сначала эти. Далее будут переназначены пользователем во время работы в GUI
        public static string results_dir = "Скопируте сюда путь";  // путь, где будет сохранён файл с результатами
        public static string ResFileName = "results";              // файл с результатами расчётов
        public static string name_object = "ДОМ";                  // имя замеряемого объекта
        public static double S = 0.0;                              // текущее значение площади
        public static string name_object_minus = "окно";           // имя вычитаемого объекта
        public static int object_minus_cnt = 1;                    // количество одинаковых объектов
        
        // Словарь уникальных имён вычитаемых объектов
        public static Dictionary<string, double> object_minus_S_sum = new Dictionary<string, double>(); // площади
        // целевые детали
        public static Dictionary<string, string> object_minus_target = new Dictionary<string, string>(); 

        public static int point_cnt = 4;                           // число точек в замеряемой площади
        public static int index = 1;             // уникальное значение, присваиваемое каждому выходному результату 

        public static double scale;       // масштабный коэффициент
        public static int sign = 1;       // знак (Положительнйы или отрицательный для величины площади)
        public static double LdrPosXstep;      // X позиция выноски относительно нулевой точки фигуры
        public static double LdrPosYstep;      // Y позиция выноски относительно нулевой точки фигуры
        public static double LdrShelfWidth;    // ширина полки выноски

        public static string name_detail = "газоблок";   // имя детали
        // Словарь уникальных имён детали
        public static Dictionary<string, double> name_detail_S_sum = new Dictionary<string, double>();

        //Получение значения текущей Базы данных
        public static Document doc = Application.DocumentManager.MdiActiveDocument;

        [CommandMethod("m2start")]
        public static void mstart() // команда для первого запуска
        {
            mdir();      // Запрос пути, где будет сохранён файл с результатами
            mobject();   // Запрос имени  замеряемого объекта
            mdetail();   // Запрос имени детали 
            mindex();    // Запрос начального значения индекса
            mscale();    // Запрос масштабного коэффициента
            msignarea(); // Запрос знака для значения площади 
            mleader();   // Запрос параметров выноски
        }

        [CommandMethod("m2dir")]
        public static void mdir() // Запрос пути, где будет сохранён файл с результатами
        { 
            // путь к папке для файла с результатами
            PromptStringOptions pStrOpts1 = new PromptStringOptions("\nВведите путь, где будет сохранён файл с результатами ");
            pStrOpts1.AllowSpaces = true; // допускаюстся пробелы в имени
            pStrOpts1.UseDefaultValue = true;            // флаг допустимости использования значения по умолчанию
            pStrOpts1.DefaultValue = results_dir;                            // значение по умолчанию при старте
            PromptResult pStrRes1 = doc.Editor.GetString(pStrOpts1);          // запрос имени у пользователя
            results_dir = pStrRes1.StringResult;                             // запись имени в глобальную переменную

            // название файла для результатов
            PromptStringOptions pStrOpts = new PromptStringOptions("\nВведите имя файла для результататов ");
            pStrOpts.AllowSpaces = true;                             // допускаюстся пробелы в имени
            pStrOpts.UseDefaultValue = true;         // флаг допустимости использования значения по умолчанию
            pStrOpts.DefaultValue = ResFileName;                     // значение по умолчанию при старте
            PromptResult pStrRes = doc.Editor.GetString(pStrOpts);   // запрос имени у пользователя
            ResFileName = pStrRes.StringResult;                      // запись имени в глобальную переменную
        }

        [CommandMethod("m2index")]
        public static void mindex() // Запрос начального значения индекса
        {
            PromptIntegerOptions pIntOpt = new PromptIntegerOptions("");
            pIntOpt.Message = "\nУкажите начальное значение индекса";

            // ограничения на значение
            pIntOpt.AllowNegative = false;    // ввод только положительных значений
            pIntOpt.AllowZero = true;         // ввод значений в том числе равных нулю
            pIntOpt.AllowNone = false;        // нельзя поле поставлять пустым
            pIntOpt.UseDefaultValue = true;   // допускается использование значения по умолчанию
            pIntOpt.DefaultValue = index;     // значение по умолчанию

            PromptIntegerResult pIntRes;      // класс, который сохраняет результат, введённый пользователем
            pIntRes = doc.Editor.GetInteger(pIntOpt); // Запрос индекса
            index = pIntRes.Value; // переписывание значения глобальной переменной индекса
        }

        [CommandMethod("m2object")]
        public static void mobject() // Запрос имени  замеряемого объекта
        {
            PromptStringOptions pStrOpts = new PromptStringOptions("\nВведите имя замеряемого объекта: ");
            pStrOpts.AllowSpaces = true; // допускаюстся пробелы в имени
            pStrOpts.UseDefaultValue = true;            // флаг допустимости использования значения по умолчанию
            pStrOpts.DefaultValue = name_object;                   // значение по умолчанию при старте
            PromptResult pStrRes = doc.Editor.GetString(pStrOpts); // запрос имени у пользователя
            name_object = pStrRes.StringResult;                    // запись имени в глобальную переменную
        }


        [CommandMethod("m2scale")]
        public static void mscale() // Запрос масштабного коэффициента
        {           
            PromptDoubleResult pDblRes;
            pDblRes = doc.Editor.GetDistance("\nУкажите две точки, между которыми указан линейный размер на чертеже ");
            double dist_model = pDblRes.Value;  // расстояние между двумя точками в модели
            pDblRes = doc.Editor.GetDouble("\nУкажите значение линейного размера");
            double dist_layout = pDblRes.Value; // значение линейного размера, указанного на чертеже
            scale = dist_layout / dist_model;   // //запись масштабного коэффициента в глобальную переменную
        }

        [CommandMethod("m2detail")]
        public static void mdetail() // Запрос имени детали 
        {
            PromptStringOptions pStrOpts = new PromptStringOptions("\nВведите название элемента фасада: ");
            pStrOpts.AllowSpaces = false;                           // не допускаюстся пробелы в имени
            pStrOpts.UseDefaultValue = true;            // флаг допустимости использования значения по умолчанию
            pStrOpts.DefaultValue = name_detail;                     // значение по умолчанию при старте
            PromptResult pStrRes = doc.Editor.GetString(pStrOpts); // запрос имени у пользователя
            name_detail = pStrRes.StringResult;                      // запись имени в глобальную переменную
            if (!name_detail_S_sum.ContainsKey(name_detail))
            {
                name_detail_S_sum.Add( name_detail, 0.0 );
            }
        }

        [CommandMethod("m2leader")]
        public static void mleader() // Запрос параметров выноски
        {
            PromptDoubleResult pDblRes;
            pDblRes = doc.Editor.GetDouble("\nУкажите смещение выноски вдоль оси X относительно первой точки");
            LdrPosXstep = pDblRes.Value;  // расстояние между двумя точками в модели
            pDblRes = doc.Editor.GetDouble("\nУкажите смещение выноски вдоль оси Y относительно первой точки");
            LdrPosYstep = pDblRes.Value; // значение линейного размера, указанного на чертеже
            pDblRes = doc.Editor.GetDouble("\nУкажите ширину полки выноски");
            LdrShelfWidth = pDblRes.Value; //значение  ширины полки выноски
        }

        [CommandMethod("m2signarea")]
        public static void msignarea() // Запрос знака для значения площади 
        {
            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
            pKeyOpts.Message = "\nВыберите знак величины площади ";
            pKeyOpts.Keywords.Add("Положительный");
            pKeyOpts.Keywords.Add("Отрицательный");
            pKeyOpts.AllowNone = false;
            PromptResult pKeyRes = doc.Editor.GetKeywords(pKeyOpts);
            if (pKeyRes.StringResult == "Положительный" )
            {
                sign = 1;
            }
            else if (pKeyRes.StringResult == "Отрицательный" )
            {
                sign = -1;
            }
        }
        [CommandMethod("m2")]
        public static void m()
        {
            // текст выноски
            string LeaderText = "\nОбъект: " + name_object + " Деталь: " + name_object;

            marea(LeaderText); // Построение примитива, замер площади, создание аннотации

            // Запись резуьтата
            double ResultS = sign * S; // величина замеренной площади с учётом знака
            name_detail_S_sum[name_detail] += ResultS; // Запись в словарь
            // далее в файл
            string ResFullFileName = results_dir + @"\" + ResFileName + ".txt";
            string TextToWrite = index.ToString() + " " + name_object + " " + name_detail + " " + ResultS.ToString();
            index++;
            using (StreamWriter sw = new StreamWriter(ResFullFileName, true))
            {
                sw.WriteLine(TextToWrite);
            }
        }

        public static void marea(string LeaderTextPlus) 
        {
            Database db = doc.Database;

            // Запрос количества точек

            PromptIntegerOptions pIntOpt = new PromptIntegerOptions("");
            pIntOpt.Message = "\nУкажите количество точек фигуры,у которой необходимо измерить площадь";

            // ограничения на значение
            pIntOpt.AllowNegative = false;    // ввод только положительных значений
            pIntOpt.AllowZero = false;        // ввод значений не равных нулю
            pIntOpt.AllowNone = false;        // нельзя поле поставлять пустым
            pIntOpt.UseDefaultValue = true;   // допускается использование значения по умолчанию
            pIntOpt.DefaultValue = point_cnt; // значение по умолчанию
            
            PromptIntegerResult pIntRes;      // класс, который сохраняет результат, введённый пользователем
            pIntRes = doc.Editor.GetInteger(pIntOpt); // Запрос количества точек 
            point_cnt = pIntRes.Value; // переписывание значения глобальной переменной числа точек


            // Class holds the result of a prompt that returns a point as its primary result    
            PromptPointResult pPtRes;
            Point2dCollection colPt = new Point2dCollection(); // Input collection of fit points

            // Class represents optional parameters for a prompt for point.
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            // Запрос первой точки
            pPtOpts.Message = "Укажите положение первой точки в поле модели:\n";
            pPtRes = doc.Editor.GetPoint(pPtOpts); // Пользователь указывает точку в модели; Программа получает точку
            colPt.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y)); // создание точки по координатам указанной

              // Exit if the user presses ESC or cancels the command
             if (pPtRes.Status == PromptStatus.Cancel) return;

            for (int i = 0; i <= point_cnt - 2; i++)
            {
                pPtOpts.Message = "Укажите положение " + (i+2).ToString()+"-й" + " точки в поле модели:\n";
                // Использование предыдущей точки как базовой точки
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = pPtRes.Value;
 
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                colPt.Add(new Point2d(pPtRes.Value.X, pPtRes.Value.Y));
 
                if (pPtRes.Status == PromptStatus.Cancel) return;
            }
            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(db.BlockTableId,
                                            OpenMode.ForRead) as BlockTable;
 
                // Open the Block table record Model space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;           
                // Create a polyline with 5 points
                Polyline acPoly = new Polyline();
                for (int i = 0; i <= point_cnt - 1; i++)
                {
                    acPoly.AddVertexAt(i, colPt[i], 0, 0, 0);                
                } 
                // Close the polyline
                acPoly.Closed = true;
                acPoly.SetDatabaseDefaults();

                // Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly);
                acTrans.AddNewlyCreatedDBObject(acPoly, true);

                // Offset the polyline a given distance
                DBObjectCollection acDbObjColl = acPoly.GetOffsetCurves(0.25);

                // Step through the new objects created
                foreach (Entity acEnt in acDbObjColl)
                {
                    // Add each offset object
                    acBlkTblRec.AppendEntity(acEnt);
                    acTrans.AddNewlyCreatedDBObject(acEnt, true);
                }

                S = Math.Pow(scale, 2) * acPoly.Area; // вычисление значения площади с учётом масштаба
                
                using (Transaction acTrans1 = db.TransactionManager.StartTransaction()) // Добавление выноски
                {
                    // Create a multiline text object
                    MText acMText = new MText();
                    acMText.SetDatabaseDefaults();
                    acMText.Location = new Point3d(colPt[0].X + LdrPosXstep, colPt[0].Y + LdrPosYstep, 0);
                    acMText.Width = LdrShelfWidth;
                    acMText.TextStyleId = db.Textstyle;
                    // формат "{{\\H1.5x;Big text}\\A2; over text\\A1;/\\A0;under text}"
                    acMText.Contents = "{{\\H1.5x;" + " Индекс: " + index.ToString() + " Площадь: " + S.ToString() +
                                           LeaderTextPlus + "}}";
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans1.AddNewlyCreatedDBObject(acMText, true);

                    // Create the leader
                    Leader acLdr = new Leader();
                    acLdr.SetDatabaseDefaults();
                    acLdr.AppendVertex(new Point3d(colPt[0].X, colPt[0].Y, 0));
                    acLdr.AppendVertex(new Point3d(acMText.Location.X, acMText.Location.Y, 0));
                    
                    // положение третьей точки выноски
                    double LdrNextXCoor = acMText.Location.X + Math.Sign(acMText.Location.X) * LdrShelfWidth;

                    acLdr.AppendVertex(new Point3d(LdrNextXCoor, acMText.Location.Y, 0));
                    acLdr.HasArrowHead = true;

                    // Add the new object to Model space and the transaction
                    acBlkTblRec.AppendEntity(acLdr);
                    acTrans1.AddNewlyCreatedDBObject(acLdr, true);

                    // Attach the annotation after the leader object is added
                    acLdr.Annotation = acMText.ObjectId;
                    acLdr.EvaluateLeader();

                    // Save the new object to the database
                    acTrans1.Commit();
                }
                // Save the new object to the database
                acTrans.Commit();
            }
        }

        [CommandMethod("m2calc")]
        public static void mcalc() // Запрос параметров выноски
        {
            string ResFullFileName = results_dir + @"\" + ResFileName + "_all" + ".txt";
            using (StreamWriter sw = new StreamWriter(ResFullFileName, true))
            {
                foreach (string name in name_detail_S_sum.Keys)
                {
                    string TextToWrite = name + " " + name_detail_S_sum[name].ToString();
                    sw.WriteLine(TextToWrite);
                }
                
            }                        
        }

        [CommandMethod("m2delelasts")]
        public static void mdelelasts() // Удаление последнего значения площади
        {
            name_detail_S_sum[name_detail] -= S;
        }

        [CommandMethod("m2minus")]
        public static void mminus() // Вычитание из указанной "большой" площади суммы "малых" площадей
        {
            // Запрос имени, из площади которого происходит вычитание

            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("\nВыбирите имя объекта," +
                                                                      "из площади которого необходимо вычитание");
            foreach (string name in name_detail_S_sum.Keys)
            {
                pKeyOpts.Keywords.Add(name);
            }            
            pKeyOpts.AllowNone = false;
            PromptResult pKeyRes = doc.Editor.GetKeywords(pKeyOpts);
            string target_name = pKeyRes.StringResult; // выбранное имя объекта, из площади которого происходит вычитание


            // Запрос имени вычитаемого объекта

            PromptStringOptions pStrOpts1 = new PromptStringOptions("\nВведите имя вычитаемого объекта");
            pStrOpts1.AllowSpaces = false; // не допускаюстся пробелы в имени
            pStrOpts1.UseDefaultValue = true;            // флаг допустимости использования значения по умолчанию
            pStrOpts1.DefaultValue = name_object_minus;  // значение по умолчанию при старте
            PromptResult pStrRes1 = doc.Editor.GetString(pStrOpts1); // запрос имени у пользователя
            name_object_minus = pStrRes1.StringResult;               // запись имени в глобальную переменную

            if (!object_minus_S_sum.ContainsKey(name_object_minus))
            {
                object_minus_S_sum.Add(name_object_minus, 0.0);
            }
            if (!object_minus_target.ContainsKey(name_object_minus))
            {
                object_minus_target.Add(name_object_minus, target_name);
            }


            // Запрос количества площадей в вычитаемом объекте

            PromptIntegerOptions pIntOpts = new PromptIntegerOptions("\nВведите количество площадей в вычитаемом объекте");
            pIntOpts.AllowNegative = false;
            pIntOpts.AllowNone = false;
            pIntOpts.UseDefaultValue = true;
            pIntOpts.DefaultValue = 1;
            PromptIntegerResult pIntRes = doc.Editor.GetInteger(pIntOpts);
            int iend = pIntRes.Value; // конечное значение цикла, равное количеству площадей

            // Запрос количества одинаковых вычитаемых объектов

            PromptIntegerOptions pIntOpts1 = new PromptIntegerOptions("\nВведите количество одинаковых вычитаемых объектов");
            pIntOpts1.AllowNegative = false;
            pIntOpts1.AllowNone = false;
            pIntOpts1.UseDefaultValue = true;
            pIntOpts1.DefaultValue = object_minus_cnt;
            PromptIntegerResult pIntRes1 = doc.Editor.GetInteger(pIntOpts1);
            object_minus_cnt = pIntRes1.Value;

            double ResultS = 0.0; // суммарная площадь вычитаемых объектов

            for (int i = 0; i <= iend - 1; i++)
            {
                string LeaderText = "\nОбъект: " + name_object + " Целевая Деталь: " + target_name +
                                    "\nДеталь: " + name_object_minus + " №" + (i+1).ToString()+ "(из "+ iend.ToString()
                                    + ")" + "\nКоличество одинаковых деталей " + object_minus_cnt.ToString();
                marea(LeaderText); // Построение примитива, замер площади, создание аннотации
                ResultS += S;      // Суммарная величина замеренной площади
            }

            // Суммарная площадь одинаковых объектов
            S = - object_minus_cnt * ResultS;
            object_minus_S_sum[name_object_minus] += S; 

            // площадь вычитаемого объета уменьшает площадь целевой детали
            name_detail_S_sum[target_name] += S;

            // далее запись результатов в файл
            string ResFullFileName = results_dir + @"\" + ResFileName+ "_minus" + ".txt";
            string TextToWrite = index.ToString() + " " + name_object + " " + target_name + " " + name_object_minus +
               " " + object_minus_cnt.ToString() + " " +  ResultS.ToString();
            index++;
            using (StreamWriter sw = new StreamWriter(ResFullFileName, true))
            {
                sw.WriteLine(TextToWrite);
            }
        }
    }

}
