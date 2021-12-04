using System;
using System.Reflection;

namespace MacroRunner.Helpers;

public static class ObjectExtensions
{
    private static MethodInfo GetMethod(this object obj, string name)
    {
        var method = obj.GetType().GetMethod(name);
        if(method is not null)
        {
            return method;
        }

        throw new Exception($"Invalid method {name}");
    }
        
    private static PropertyInfo GetProperty(this object obj, string name)
    {
        var property = obj.GetType().GetProperty(name);
        if(property is not null)
        {
            return property;
        }

        throw new Exception($"Invalid property {name}");
    }
        
    private static FieldInfo GetField(this object obj, string name)
    {
        var field = obj.GetType().GetField(name);
        if(field is not null)
        {
            return field;
        }

        throw new Exception($"Invalid field {name}");
    }
        
    public static T Call<T>(this object obj, string name, params object[] parameters)
    {
        obj.ThrowIfNull();
        return (T)obj.GetMethod(name).Invoke(obj, parameters)!;
    }

    public static void CallVoid(this object obj, string name, params object[] parameters)
    {
        obj.ThrowIfNull();
        obj.GetMethod(name).Invoke(obj, parameters);
    }

    public static T GetField<T>(this object obj, string name)
    {
        obj.ThrowIfNull();

        return (T)obj.GetField(name).GetValue(obj)!;
    }

    public static T GetProperty<T>(this object obj, string name)
    {
        obj.ThrowIfNull();

        return (T)obj.GetProperty(name).GetValue(obj)!;
    }

    public static void SetField<T>(this object obj, string name, T value)
    {
        obj.ThrowIfNull();
        obj.GetField(name).SetValue(obj, value);
    }

    public static void SetProperty<T>(this object obj, string name, T value)
    {
        obj.ThrowIfNull();
        obj.GetProperty(name).SetValue(obj, value);
    }

    public static void ThrowIfNull(this object obj)
    {
        if (obj == null)
        {
            throw new Exception("Object can't be null");
        }
    }
}