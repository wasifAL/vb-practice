Imports System.Reflection

Module Module1
    Sub Main()
        Dim assem As Assembly = GetType(Module1).Assembly
        Dim x = 1
        Const className As String = "FieldDetails"
        Dim pNames = {"name", "date"}
        Dim pTypes = {GetType(String), GetType(String)}
        While True
            Console.WriteLine("Iteration : " + x.ToString())

            Dim DC As New DynamicClass(className)
            Dim obj = DC.CreateObject(pNames, pTypes)
'            FieldDetails, FieldDetails, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null
            Dim type As Type = obj.GetType()
            Dim ass  = Assembly.GetAssembly(type)
            For Each t As Type In ass.GetTypes()
                Console.WriteLine(t.AssemblyQualifiedName)
            Next
            Dim type5= ass.GetTypes().First(Function(y As Type) y.name = className )
            
            
            DC.setValue(obj, "name", "file " + x.ToString())
            DC.setValue(obj, "date", DateTime.Now.ToString())
            Console.WriteLine("Assembly Qualified Name : "+obj.GetType().AssemblyQualifiedName)
            For Each pi As PropertyInfo In type.GetProperties()
                Console.Write(pi.Name)
                Console.Write("  :  ")
                Console.Write(DC.getValue(obj, pi.Name))
                Console.WriteLine()
            Next
            x += 1

        End While
    End Sub
End Module
