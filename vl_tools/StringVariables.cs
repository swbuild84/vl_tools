using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace vl_tools
{
    public class strigVariables
    {
        Dictionary <string, string> variables;
        public strigVariables()
        {
            variables = new Dictionary <string, string>();
        }
        // получаем значение переменной по ее имени
        public int GetVariable(string name)
        {
            return variables[name];
        }

        public void SetVariable(string name, string value)
        {
            if (variables.ContainsKey(name))
                variables[name] = value;
            else
                variables.Add(name, value);
        }
    }
}
