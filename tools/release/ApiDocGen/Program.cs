// Генератор справочника API (Markdown) по сборке и её XML-документации.
// Использование: dotnet run --project tools/release/ApiDocGen -- <DocxToPdf.dll> <DocxToPdf.xml> <api.md>
using System.Reflection;
using System.Text;
using System.Xml.Linq;
var asm = Assembly.LoadFrom(args[0]);
var nctx = new NullabilityInfoContext();
string PName(ParameterInfo p) => TName(p.ParameterType) + (!p.ParameterType.IsValueType && nctx.Create(p).WriteState == NullabilityState.Nullable ? "?" : "");
var docs = XDocument.Load(args[1]).Descendants("member").ToDictionary(m => (string)m.Attribute("name")!, m => m);
string Clean(XElement? e)
{
    if (e == null) return "";
    var sb = new StringBuilder();
    foreach (var n in e.Nodes())
    {
        if (n is XText t) sb.Append(t.Value);
        else if (n is XElement x)
        {
            string? cref = (string?)x.Attribute("cref") ?? (string?)x.Attribute("name");
            if (x.Name == "see" || x.Name == "paramref" || x.Name == "typeparamref") sb.Append('`').Append(cref?.Split(':').Last().Split('.').Last()).Append('`');
            else if (x.Name == "c") sb.Append('`').Append(x.Value).Append('`');
            else sb.Append(Clean(x));
        }
    }
    return System.Text.RegularExpressions.Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
}
string Summary(string id) => docs.TryGetValue(id, out var m) ? Clean(m.Element("summary")) : "";
string TName(Type t)
{
    if (t.IsGenericType)
    {
        string n = t.Name[..t.Name.IndexOf('`')];
        if (n == "Nullable") return TName(t.GetGenericArguments()[0]) + "?";
        return n + "<" + string.Join(", ", t.GetGenericArguments().Select(TName)) + ">";
    }
    if (t.IsArray) return TName(t.GetElementType()!) + "[]";
    return t.Name switch { "String" => "string", "Int32" => "int", "Double" => "double", "Boolean" => "bool", "Void" => "void", "Byte" => "byte", "Object" => "object", "UInt16" => "ushort", "Int64" => "long", "Char" => "char", "Single" => "float", _ => t.Name };
}
string XName(Type t) => t.IsGenericType ? t.GetGenericTypeDefinition().FullName!.Replace('+', '.') : (t.FullName ?? t.Name).Replace('+', '.');
string ParamId(Type t)
{
    if (t.IsByRef) return ParamId(t.GetElementType()!) + "@";
    if (t.IsArray) return ParamId(t.GetElementType()!) + "[]";
    if (t.IsGenericType)
    {
        string n = t.GetGenericTypeDefinition().FullName!; n = n[..n.IndexOf('`')];
        return n.Replace('+', '.') + "{" + string.Join(",", t.GetGenericArguments().Select(ParamId)) + "}";
    }
    return (t.FullName ?? t.Name).Replace('+', '.');
}
var types = asm.GetExportedTypes().Where(t => !t.IsNested || t.DeclaringType!.IsPublic).OrderBy(t => t.Namespace).ThenBy(t => t.FullName).ToList();
var key = new[] { "DocxToPdf.Converter", "DocxToPdf.Fonts.FontManager", "DocxToPdf.Html.HtmlOptions", "DocxToPdf.Docx.DocxOptions", "DocxToPdf.Program" };
var sb = new StringBuilder();
sb.AppendLine($"# Справочник API DocxToPdf {asm.GetName().Version?.ToString(3)}").AppendLine();
sb.AppendLine("Сформирован по сборке `DocxToPdf.dll` и её XML-документации. Основные типы — в первом разделе, остальные");
sb.AppendLine("публичные типы (модель документа, вёрстка, PDF) доступны для расширенного использования, но могут меняться между версиями.").AppendLine();
void Type1(Type t, bool full)
{
    string kind = t.IsEnum ? "enum" : t.IsValueType ? "struct" : t.IsInterface ? "interface" : t.IsAbstract && t.IsSealed ? "static class" : t.IsAbstract ? "abstract class" : "class";
    if (t.GetMethod("<Clone>$") != null) kind = "record";
    sb.AppendLine($"### `{TName(t)}` — {kind}").AppendLine();
    sb.AppendLine($"Пространство имён: `{t.Namespace}`").AppendLine();
    if (Summary("T:" + XName(t)) is { Length: > 0 } s) sb.AppendLine(s).AppendLine();
    if (!full) return;
    const BindingFlags F = BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
    var rows = new List<string>();
    foreach (var c in t.GetConstructors(F))
    {
        var ps = c.GetParameters();
        string id = $"M:{XName(t)}.#ctor" + (ps.Length > 0 ? "(" + string.Join(",", ps.Select(p => ParamId(p.ParameterType))) + ")" : "");
        rows.Add($"| `new {TName(t)}({string.Join(", ", ps.Select(p => PName(p) + " " + p.Name + (p.HasDefaultValue ? " = " + (p.DefaultValue ?? "null") : "")))})` | {Summary(id)} |");
    }
    foreach (var m in t.GetMethods(F).Where(m => !m.IsSpecialName && m.DeclaringType == t && !m.Name.StartsWith('<') && m.Name is not ("Equals" or "GetHashCode" or "ToString" or "Deconstruct" or "PrintMembers")))
    {
        var ps = m.GetParameters();
        string id = $"M:{XName(t)}.{m.Name}" + (ps.Length > 0 ? "(" + string.Join(",", ps.Select(p => ParamId(p.ParameterType))) + ")" : "");
        string sig = $"{(m.IsStatic ? "static " : "")}{TName(m.ReturnType)} {m.Name}({string.Join(", ", ps.Select(p => PName(p) + " " + p.Name + (p.HasDefaultValue ? " = " + (p.DefaultValue is string ds ? "\"" + ds + "\"" : p.DefaultValue ?? "null") : "")))})";
        string sum = Summary(id);
        if (docs.TryGetValue(id, out var md))
        {
            foreach (var p in md.Elements("param")) sum += $"<br>`{(string?)p.Attribute("name")}` — {Clean(p)}";
            if (md.Element("returns") is { } r) sum += $"<br>Возвращает: {Clean(r)}";
            foreach (var ex in md.Elements("exception")) sum += $"<br>Исключение `{((string?)ex.Attribute("cref"))?.Split('.').Last()}`: {Clean(ex)}";
            if (md.Element("remarks") is { } rm) sum += $"<br>{Clean(rm)}";
        }
        rows.Add($"| `{sig}` | {sum} |");
    }
    foreach (var f in t.GetFields(F).Where(f => !f.IsSpecialName))
        rows.Add($"| `{(f.IsStatic && !f.IsLiteral ? "static " : "")}{TName(f.FieldType)} {f.Name}`{(t.IsEnum ? "" : f.IsLiteral ? " (const)" : "")} | {Summary($"F:{XName(t)}.{f.Name}")} |");
    foreach (var p in t.GetProperties(F))
        rows.Add($"| `{TName(p.PropertyType)} {p.Name} {{ {(p.GetMethod?.IsPublic == true ? "get; " : "")}{(p.SetMethod?.IsPublic == true ? "set; " : "")}}}` | {Summary($"P:{XName(t)}.{p.Name}")} |");
    if (rows.Count > 0)
    {
        sb.AppendLine("| Член | Описание |").AppendLine("|---|---|");
        foreach (var r in rows) sb.AppendLine(r.Replace("\n", " "));
        sb.AppendLine();
    }
}
sb.AppendLine("## Основные типы").AppendLine();
foreach (var k in key) if (types.FirstOrDefault(t => t.FullName == k) is { } t) Type1(t, true);
sb.AppendLine("## Прочие публичные типы").AppendLine();
foreach (var g in types.Where(t => !key.Contains(t.FullName)).GroupBy(t => t.Namespace))
{
    sb.AppendLine($"### Пространство имён `{g.Key}`").AppendLine();
    sb.AppendLine("| Тип | Описание |").AppendLine("|---|---|");
    foreach (var t in g) sb.AppendLine($"| `{TName(t)}` | {Summary("T:" + XName(t))} |");
    sb.AppendLine();
}
File.WriteAllText(args[2], sb.ToString());
Console.WriteLine($"{types.Count} типов");
