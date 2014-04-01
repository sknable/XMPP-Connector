using Mono.CSharp;
using System;
using System.Reflection;

namespace ChatScript
{
    public class CSScriptEngine
    {
        private Evaluator REPLSharpCompiler;
        private Evaluator SharpCompiler;

        public CSScriptEngine()
        {
            CompilerContext context = new CompilerContext(new CompilerSettings(), new ConsoleReportPrinter());
            SharpCompiler = new Evaluator(context);
            SharpCompiler.LoadAssembly("XMPP-Web.dll");

            REPLSharpCompiler = new Evaluator(context);
            REPLSharpCompiler.LoadAssembly("XMPP-Web.dll");
            REPLSharpCompiler.Run(@"        using System;
                                            using XMPP_Web;
                                            using System.Collections.Generic;
                                            using System.Linq;
                                            using System.Text;");
        }

        public string LoadedCSharpCode { get; private set; }

        public dynamic Script { get; private set; }

        public object Eval(String line)
        {
            object result;
            bool resultSet;

            String x = REPLSharpCompiler.Evaluate(line, out result, out resultSet);

            if (resultSet)
            {
                return result;
            }
            else
            {
                return null;
            }
        }

        public bool LoadScriptFile(string fileNameAndPath)
        {
            try
            {
                LoadedCSharpCode = System.IO.File.ReadAllText(fileNameAndPath);
            }
            catch
            {
                return false;
            }

            SharpCompiler.Compile(LoadedCSharpCode);

            Assembly asm = ((Type)SharpCompiler.Evaluate("typeof(XMPP_Script);")).Assembly;

            Script = asm.CreateInstance("XMPP_Script");

            return true;
        }

        public bool Run(String line)
        {
            return REPLSharpCompiler.Run(line);
        }
    }
}