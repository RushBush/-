using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            
            if (args.Length!=3)
            {
                Console.Out.WriteLine("Неверно введены данные\nИспользование: programName <excelFile1> <wordFile> <excelFile2>");
                Console.Out.WriteLine("<excelFile1> - файл, содержащий список студентов");
                Console.Out.WriteLine("<wordFile> -шаблон для формирования писем");
                Console.Out.WriteLine("<excelFile2> - файл, содержащий описание и время практик\n\n\n");

                return;
            }

            Dictionary<String, HashSet<String>> contracts = new Dictionary<string, HashSet<string>>(); //словарь ключ-пара <город,договоры>
            Dictionary<String, String> address = new Dictionary<string, string>();//адреса назначения <город,адрес>
            Dictionary<String, Dictionary<String, String>> direct = new Dictionary<string, Dictionary<string, string>>(); //словарь ключ-пара <город,<направление,студенты >>
            Dictionary<Tuple<String, int>, HashSet<Tuple<string, int>>> course = new Dictionary<Tuple<string, int>, HashSet<Tuple<string, int>>>(); //<<направление,курс>,<описание,время>>
            Dictionary<String, Dictionary<String, String>> practice = new Dictionary<string, Dictionary<string, string>>(); // <город,<описание и время,список студентов>>    

            excel.Application excelAppDir = new excel.Application();
            excel.Workbook workbookDir = excelAppDir.Workbooks.Open(args[2], ReadOnly: true);
            excel.Worksheet worksheetDir = workbookDir.Sheets[1];
            excel.Range excelRangeDir = worksheetDir.UsedRange;

            int index = 2;

            while (excelRangeDir.Cells[index,1].Value2 != null) //парсинг excel-файла с практикой
            {
                
                string courseName = excelRangeDir.Cells[index, 1].Value2.ToString();
                string courseType = excelRangeDir.Cells[index, 3].Value2.ToString();
                int courseNum;
                int courseTime;

                if (!int.TryParse(excelRangeDir.Cells[index, 2].Value2.ToString(),out courseNum))
                {
                    Console.WriteLine("Файл " + args[2] + " имеет пустые ячейки в строке " + index);
                }

                if (!int.TryParse(excelRangeDir.Cells[index, 4].Value2.ToString(), out courseTime))
                {
                    Console.WriteLine("Файл " + args[2] + " имеет пустые ячейки в строке " + index);
                }


                if (course.ContainsKey(Tuple.Create(courseName, courseNum)))
                {
                    course[Tuple.Create(courseName, courseNum)].Add(Tuple.Create(courseType, courseTime));
                }
                else
                {
                    course.Add(Tuple.Create(courseName, courseNum), new HashSet<Tuple<string, int>>{ Tuple.Create(courseType, courseTime) });
                }


                index++;
            }

            workbookDir.Close(); // закрытие "книги" excel
            excelAppDir.Quit(); // завершение процесса excel
            

            index = 1;

            excel.Application excelApp = new excel.Application();
            excel.Workbook workbook = excelApp.Workbooks.Open(args[0], ReadOnly: true);
            excel.Worksheet worksheet = workbook.Sheets[1];
            excel.Range excelRange = worksheet.UsedRange;

            int CourseNum = 1;

            while (excelRange.Cells[index,1].Value2!=null || excelRange.Cells[index+1, 1].Value2 != null)
            {
                if (excelRange.Cells[index, 1].Value2 != null && excelRange.Cells[index, 1].Value2[0]!='2')
                {
                    string name = excelRange.Cells[index, 1].Value2.ToString();
                    string faculty = excelRange.Cells[index, 3].Value2.ToString();
                    string city = excelRange.Cells[index, 5].Value2.ToString();
                    string contract = excelRange.Cells[index, 6].Value2.ToString();
                    string addr = excelRange.Cells[index, 7].Value2.ToString();

                    if (course.ContainsKey(Tuple.Create(faculty, CourseNum))) // есть ли у студента практика,проверка по направлению и курсу
                    {
                        if (address.ContainsKey(city)) //есть ли город в списке
                        {
                            contracts[city].Add(contract); //Город есть => добавляем контракт студента

                            if (direct[city].ContainsKey(faculty)) //содержит список 
                            {
                                direct[city][faculty] = String.Join(", ", direct[city][faculty].ToString(), ShortName(name));
                            }
                            else
                            {
                                direct[city].Add(faculty, ShortName(name));
                            }


                            foreach (Tuple<String, int> pair in course[Tuple.Create(faculty, CourseNum)])
                            {
                                string[] practiceInfoArr = pair.Item1.Split(' ');

                                string practiceInfo = practiceInfoArr[0] + " " + practiceInfoArr[1] + " в объеме " + pair.Item2.ToString() + " часов";

                                if (practice[city].ContainsKey(practiceInfo))
                                {
                                    practice[city][practiceInfo] = String.Join(", ", practice[city][practiceInfo], ShortName(name));
                                }
                                else
                                {
                                    practice[city].Add(practiceInfo, ShortName(name));
                                }
                            }
                        }
                        else
                        {
                            contracts.Add(city, new HashSet<string>() { contract });
                            address.Add(city, addr);
                            direct.Add(city, new Dictionary<string, string> { { faculty, ShortName(name) } });
                            practice.Add(city, new Dictionary<string, string>());

                            foreach (Tuple<String, int> pair in course[Tuple.Create(faculty, CourseNum)])
                            {
                                string[] practiceInfoArr = pair.Item1.Split(' ');

                                string practiceInfo = practiceInfoArr[0] + " " + practiceInfoArr[1] + " в объеме " + pair.Item2.ToString() + " часов";

                                if (practice[city].ContainsKey(practiceInfo))
                                {
                                    practice[city][practiceInfo] = String.Join(", ", practice[city][practiceInfo], ShortName(name));
                                }
                                else
                                {
                                    practice[city].Add(practiceInfo, ShortName(name));
                                }
                            }
                        }
                    }
                }
                else if (excelRange.Cells[index, 1].Value2 != null)
                {
                    string year = excelRange.Cells[index, 1].Value2.ToString();
                    int startYear = int.Parse(year.Substring(0, 4));
                    CourseNum = GetCourse(startYear);
                }

                index++;
            }


            workbook.Close();
            excelApp.Quit();

           
            int k = 1;
            foreach (KeyValuePair<string, string> pair in address)
            {
                
                try
                {
                    File.Delete(Path.Combine(args[1], "..", k.ToString() + ".doc"));
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                

                try
                {
                    File.Copy(args[1], Path.Combine(args[1], "..", k.ToString() + ".doc"), true);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                word.Application wordApp = new word.Application();
                word.Document doc = wordApp.Documents.Open(Path.Combine(args[1], "..", k.ToString() + ".doc"));

                ReplaceWord("{Where}", ToWhom(pair.Key, GetHead(address[pair.Key])), doc);
                ReplaceWord("{Address}", GetAddress(address[pair.Key]), doc);

                string head = GetHead(address[pair.Key]);
                string[] names = head.Split(' ');
                ArraySegment<string> segment = new ArraySegment<string>(names, names.Length - 2, 2);
                

                string name = string.Join(" ", segment);
                ReplaceWord("{ToWhom}", name, doc);

                Console.WriteLine("Введите имя и отчество: " + name);
                string secondName = Console.ReadLine();
                Console.WriteLine("\n");
                string genre = secondName[secondName.Length - 1].Equals('ч') ? "Уважаемый" : "Уважаемая";

                ReplaceWord("{ToWhom2}", genre + " " + secondName, doc);
                ReplaceWord("{City}", pair.Key, doc);

                string contractsStr = string.Join<string>(", ", contracts[pair.Key].Select(x => "№ " + x));
                ReplaceWord("{Contracts}", contractsStr, doc);

                string students = string.Empty;

                foreach (KeyValuePair<String, String> stud in direct[pair.Key])
                {
                    string editedKey = stud.Key.Insert(stud.Key.IndexOf(" ") + 1, "«");
                    editedKey += "»";
                    students += stud.Value + " (" + editedKey + "), ";
                }

                students = students.Substring(0, students.Length - 2);


                ReplaceWord("{Students}", students, doc);
                ReplaceWord("{Date}", DateTime.Now.Year.ToString(), doc);

                word.Range range = doc.Content;
                range.Find.Execute(FindText: "{StudentsTable}");

                word.Table table = doc.Tables.Add(range, practice[pair.Key].Count, 2);
                table.Borders.Enable = 1;

                int rowIndex = 1;

                foreach (KeyValuePair<String, String> practicePair in practice[pair.Key])
                {
                    table.Cell(rowIndex, 1).Range.Text = practicePair.Value;
                    table.Cell(rowIndex, 2).Range.Text = practicePair.Key;
                    table.Cell(rowIndex, 1).Range.Font.Size = 11;
                    table.Cell(rowIndex, 2).Range.Font.Size = 11;

                    rowIndex++;
                }
                doc.Save();
                doc.Close();
                wordApp.Quit();


                k++;
            }

            Console.In.Read();        
        }


        // заменяет textToReplace на text в файле wordDoc
        public static void ReplaceWord(string textToReplace,string text,word.Document wordDoc) 
        {
            var range = wordDoc.Content;
            range.Find.ClearFormatting();
            range.Font.Name = "Times New Roman";

            if (text.Length > 255)
            {
                string newStr;
                string replaceStr = textToReplace;
                for (int i = 0; i < text.Length; i += 252)
                {
                    if (i + 252 > text.Length - 1)
                        newStr = text.Substring(i, text.Length - i);
                    else
                        newStr = text.Substring(i, 252) + "{@}";


                    ReplaceWord(replaceStr, newStr, wordDoc);


                    replaceStr = "{@}";
                }
            }
            else
                range.Find.Execute(FindText: textToReplace, ReplaceWith: text);
        }

        public static string GetHead(string address)
        {
            string result = address;
            int lastDigit = 0;
            for (int i=0;i<result.Length;i++)
            {
                if (char.IsDigit(result[i]) || result[i].Equals(','))
                {
                    lastDigit = i;
                }
            }

            if (result[lastDigit + 1].Equals(','))
                lastDigit++;

            result = result.Substring(lastDigit + 1).Trim();

            return result;
        }

        public static string ToWhom(string city,string address)
        {
            string[] vs = address.Split(' ');
            ArraySegment<string> segment = new ArraySegment<string>(vs,0,vs.Length-2);
            string result = string.Join(" ",segment);

            result += " " + city;

            return result;
        }

        public static string ShortName(string name)
        {
            string[] nameList = name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string result;

            result = nameList[0] + " " + nameList[1][0] + "." + nameList[2][0] + ".";
            return result;
        }

        public static int GetCourse(int startYear)
        {
            return (int.Parse(DateTime.Now.Year.ToString())-startYear);
        }

        public static string GetAddress(string address)
        {
            string[] addr = address.Split(',');
            Array.Reverse(addr);
            addr = addr.Where(x => x != addr[0]).ToArray();

            string swap = addr[0];
            addr[0] = addr[1];
            addr[1] = swap;

            return String.Join(",", addr);
        }
    }
}
