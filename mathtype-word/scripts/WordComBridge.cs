using System;

public static class WordComBridge
{
    public static bool TryGetCustomProperty(object document, string name, out string value)
    {
        dynamic properties = ((dynamic)document).CustomDocumentProperties;
        try
        {
            dynamic property = properties.Item(name);
            value = Convert.ToString(property.Value);
            return true;
        }
        catch (Exception)
        {
            value = null;
            return false;
        }
    }

    public static bool HasCustomProperty(object document, string name)
    {
        string value;
        return TryGetCustomProperty(document, name, out value);
    }

    public static string GetCustomProperty(object document, string name)
    {
        string value;
        TryGetCustomProperty(document, name, out value);
        return value;
    }

    public static void SetCustomProperty(object document, string name, string value)
    {
        dynamic properties = ((dynamic)document).CustomDocumentProperties;
        try
        {
            dynamic property = properties.Item(name);
            property.Value = value;
        }
        catch (Exception)
        {
            // 4 is msoPropertyTypeString. LinkSource is omitted because this is not linked content.
            properties.Add(name, false, 4, value, Type.Missing);
        }
    }

    public static void DeleteCustomProperty(object document, string name)
    {
        dynamic properties = ((dynamic)document).CustomDocumentProperties;
        try
        {
            dynamic property = properties.Item(name);
            property.Delete();
        }
        catch (Exception)
        {
            // Missing already has the requested result.
        }
    }
}
