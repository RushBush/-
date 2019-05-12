using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2
{
    class Student
    {
        public string Name { get; set; } 
        public string Group { get; set; }
        public string Course { get; set; }
        public string Faculty { get; set; }
        public string City { get; set; }
        public string Address { get; set; }
        public string Contract { get; set; }

        public string ShortName()
        {
            String[] name = this.Name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            String result;

            result = name[0]+" "+name[1][0]+"."+name[2][0]+".";
            return result;
        }
    }
}
