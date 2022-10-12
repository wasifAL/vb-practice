Imports System.Reflection
Imports System.Reflection.Emit

Public Class DynamicClass
    Public Shared types As New Dictionary(Of String,Type) 
    Public assemblyName As String
    Public moduleName As String
    Public className As String

    Public Sub New(assemblyName As String, moduleName As String, className As String)
        Me.assemblyName = assemblyName
        Me.moduleName = moduleName
        Me.className = className
    End Sub
    Public Function CreateClass() As TypeBuilder

        Dim an = new AssemblyName(assemblyName)
        Dim ab As AssemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(an, AssemblyBuilderAccess.RunAndCollect)
        Dim assembly As Assembly = Assembly.GetExecutingAssembly()
        Dim mb As ModuleBuilder = ab.GetDynamicModule(moduleName)
        If mb Is Nothing
            mb = ab.DefineDynamicModule(moduleName)
        End If
        Dim tb As TypeBuilder = mb.DefineType(an.FullName, TypeAttributes.Public Or TypeAttributes.Class Or TypeAttributes.AutoClass Or TypeAttributes.AnsiClass Or TypeAttributes.BeforeFieldInit Or
                                                  TypeAttributes.AutoLayout, Nothing)
        return tb
    End Function
    Public Sub CreateConstructor(tb As TypeBuilder)
        tb.DefineDefaultConstructor(Reflection.MethodAttributes.Public Or Reflection.MethodAttributes.SpecialName Or Reflection.MethodAttributes.RTSpecialName)
    End Sub
    Public Sub CreateProperty(tb As TypeBuilder, pName As String, pType As Type)

        Dim fb As FieldBuilder = tb.DefineField(pName, pType, FieldAttributes.Private)
        Dim pb As PropertyBuilder = tb.DefineProperty(pName, PropertyAttributes.HasDefault, pType, Nothing)
'        ===== Getter Generator =====        
        Dim getMB As MethodBuilder = tb.DefineMethod("get" + pName, Reflection.MethodAttributes.Public Or Reflection.MethodAttributes.SpecialName Or Reflection.MethodAttributes.HideBySig, pType,
                                                     Type.EmptyTypes)
        Dim getIl As ILGenerator = getMB.GetILGenerator()
        getIl.Emit(OpCodes.Ldarg_0)
        getIl.Emit(OpCodes.Ldfld, fb)
        getIl.Emit(OpCodes.Ret)
        pb.SetGetMethod(getMB)

'           ===== Setter Generator =====
        Dim setMB As MethodBuilder = tb.DefineMethod("set" + pName, Reflection.MethodAttributes.Public Or Reflection.MethodAttributes.SpecialName Or Reflection.MethodAttributes.HideBySig, Nothing,
                                                     {pType})
        Dim setIl As ILGenerator = setMB.GetILGenerator()
        Dim modifyProperty As Label = setIl.DefineLabel()
        Dim exitSet As Label = setIl.DefineLabel()
        setIl.MarkLabel(modifyProperty)
        setIl.Emit(OpCodes.Ldarg_0)
        setIl.Emit(OpCodes.Ldarg_1)
        setIl.Emit(OpCodes.Stfld, fb)
        setIl.Emit(OpCodes.Nop)
        setIl.MarkLabel(exitSet)
        setIl.Emit(OpCodes.Ret)
        pb.SetSetMethod(setMB)
    End Sub
    Public Sub setValue(obj As Object, pName As String, value As Object)
        obj.GetType().GetProperty(pName).SetValue(obj, value)
    End Sub
    Public Function getValue(obj As Object, pName As String) As Object
        Return obj.GetType().GetProperty(pName).GetValue(obj)
    End Function
    Public Function CreateObject(propertyNames As String(), types As Type()) As Object
        If Not (PropertyNames.Count = Types.Count) Then
            WriteLine("The number of property names should match their corresopnding types number")
        End If
        Dim DynamicClass As TypeBuilder = CreateClass()
        CreateConstructor(DynamicClass)
        For index = 0 To (PropertyNames.Count - 1)
            CreateProperty(DynamicClass, propertyNames(index), types(index))
        Next
        Dim type = DynamicClass.CreateType()
        return Activator.CreateInstance(type)
    End Function
End Class