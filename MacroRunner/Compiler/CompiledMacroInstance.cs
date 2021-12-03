using System.IO;
using System.Linq;
using System.Reflection;

namespace MacroRunner.Compiler
{
    public abstract class CompiledMacroInstance
    {
        public CompiledCode Macro { get; private set; }

        public virtual void Init(CompiledCode macro)
        {
            Macro = macro;
        }

        public abstract void LoadFile(Stream file);

        public MethodInfo GetEntryPoint(string entryPointName)
        {
            var nameParts = entryPointName.Split('.');
            var className = string.Join(".", nameParts.Take(nameParts.Length - 1));
            var methodName = nameParts.Last();
            var entryPoint = Macro.Assembly.GetTypes()
                                  .Where(x => string.IsNullOrEmpty(className) || className == x.FullName)
                                  .Select(x => x.GetMethod(methodName))
                                  .Single(x => x != null);

            return entryPoint;
        }

        public void Run(string entryPointName)
        {
            GetEntryPoint(entryPointName).Invoke(null, null);
        }
    }


    public abstract class CompiledMacroInstance<T> : CompiledMacroInstance
    {
        public T Globals { get; protected set; }

        public override void Init(CompiledCode macro)
        {
            base.Init(macro);
            InitGlobals();
        }

        public abstract void InitGlobals();
    }
}