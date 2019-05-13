using Microsoft.Office.Interop.Access;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using word = Microsoft.Office.Interop.Word;
using excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
            if (args.Length!=3)
            {
                Console.Out.WriteLine("Usage: ProgramName  ExcelFileWithStudents ExcelFileWithTime DocFile");

                return;
            }*/

           
            

            excel.Application excelApp = new excel.Application(); //Студенты
            word.Application wordApp = new word.Application();

           

            excel.Workbook workbook = excelApp.Workbooks.Open(args[0],ReadOnly:true);
            word.Document doc = wordApp.Documents.Open(Path.Combine(args[1],"..","Karasuk.doc"));

            excel.Worksheet worksheet = workbook.Sheets[1];
            excel.Range excelRange = worksheet.UsedRange;


            Dictionary<String, HashSet<String>> contracts = new Dictionary<string, HashSet<string>>(); //словарь ключ-пара <город,договоры>
            Dictionary<String, List<Student>> dict = new Dictionary<string, List<Student>>(); //словарь ключ-пара <город,список студентов>
            Dictionary<String, String> address = new Dictionary<string, string>();//адреса назначения <город,адрес>

            int index = 1;
            
            while(excelRange.Cells[index,1].Value2!=null || excelRange.Cells[index+1, 1].Value2 != null)
            {
                if (excelRange.Cells[index, 1].Value2 != null && excelRange.Cells[index, 1].Value2[0]!='2')
                {
                    Student student = new Student();
                    student.Name = excelRange.Cells[index, 1].Value2.ToString();
                    student.Group = excelRange.Cells[index, 2].Value2.ToString();
                    student.Faculty = excelRange.Cells[index, 3].Value2.ToString();
                    student.Course = excelRange.Cells[index, 4].Value2.ToString();    
                    student.Contract = excelRange.Cells[index, 6].Value2.ToString();

                    String city = excelRange.Cells[index, 5].Value2.ToString();
                    String addr = excelRange.Cells[index, 7].Value2.ToString();

                    if (dict.ContainsKey(city))
                    {
                        dict[city].Add(student);
                        contracts[city].Add(student.Contract);
                    }
                    else
                    {
                        dict.Add(city, new List<Student>() { student });
                        contracts.Add(city,new HashSet<string>() { student.Contract});
                        address.Add(city,addr);
                    }
                }

                index++;
            }


            /*
            foreach(KeyValuePair<String,String> pair in address)
            {
                Console.WriteLine(ToWhom(pair.Key,GetHead(pair.Value)));
            }
            */
            int k = 1;
            foreach(KeyValuePair<string,List<Student>>pair in dict)
            {
                if (File.Exists(Path.Combine(args[1], "..", k.ToString() + ".doc")))
                {
                    try
                    {
                        File.Delete(Path.Combine(args[1], "..", k.ToString() + ".doc"));
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }

                try
                {
                    File.Copy(args[1], Path.Combine(args[1], "..", k.ToString() + ".doc"), true);
                }
                catch(Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                
                
                k++;
            }

            
            

            doc.Save();


            workbook.Close();
            excelApp.Application.Quit();

            doc.Close();
            wordApp.Application.Quit();


            Console.In.Read();
            


        }

        public static void ReplaceWord(string textToReplace,string text,word.Document wordDoc,word.WdParagraphAlignment alignment) // заменяет textToReplace на text в файле wordDoc
        {
            var range = wordDoc.Content;
            range.Find.ClearFormatting();
            range.Font.Name = "Times New Roman";
            range.ParagraphFormat.Alignment = alignment;
            range.Find.Execute(FindText: textToReplace,ReplaceWith:text);
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
    }
}
