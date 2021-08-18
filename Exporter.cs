/* 
 * ==============================================================================
 * Filename: 
 * Created:  2021 / 8 / 12 15:51
 * Author: HuaHua
 * Purpose: Excel to Lua
 * Excel 格式
 * 第一行：描述
 * 第二行：字段名字（:c:s）表示客户端或者服务端字段
 * 第三行：字段类型
 * 第一列：id。目前只支持One Key
 * 每个Page的字段类型和数量必须保持一致。第一个Page Name将作为Lua Table Name
 * ==============================================================================
**/

using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using LuaField = System.Collections.Generic.KeyValuePair<int, string>;

public static class Exporter
{

    //config types
    enum EType
    {
        EInt,
        EString,
        EFloat,
        EArray,
    }

    class Value
    {
        public string str;
        public bool isTable = false;
        public List<Value> values;
        public int count = 0;

        private string _expand;
        public string StrValue
        {
            get
            {
                if (string.IsNullOrEmpty(_expand))
                {
                    if (values == null)
                    {
                        _expand = str;
                    }
                    else
                    {
                        var sb = new StringBuilder();
                        sb.Append("{");
                        foreach (var v in values)
                        {
                            sb.Append(v.StrValue);
                        }
                        sb.Append("}");
                        _expand = sb.ToString();
                    }
                }

                return _expand;
            }
        }
    }

    class KeyValue
    {
        public string key;
        public Value value;
    }

    #region private
    static readonly string PlaceHodler = "false";
    static bool IsNumberic(string value)
    {
        return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
    }

    static string GetArrayString(string array)
    {
        var array2 = array.Split(',');
        var a2count = array2.Length;
        var sb = new StringBuilder();
        if (string.IsNullOrEmpty(array2[a2count - 1]))
        {
            a2count--;
        }
        for (var a2i = 0; a2i < a2count; ++a2i)
        {
            if (IsNumberic(array2[a2i]))
            {
                sb.Append(array2[a2i]);
            }
            else
            {
                sb.Append($"\"{array2[a2i]}\"");
            }
            if (a2i < a2count - 1)
            {
                sb.Append(",");
            }
        }

        return sb.ToString();
    }

    static string ReadInt(IExcelDataReader reader, int index)
    {
        if (reader.IsDBNull(index))
        {
            return "0";
        }
        else if (reader.GetFieldType(index) == typeof(double))
        {
            return ((int)reader.GetDouble(index)).ToString();
        }
        else if (reader.GetFieldType(index) == typeof(int))
        {
            return reader.GetInt32(index).ToString();
        }
        else if (reader.GetFieldType(index) == typeof(string))
        {
            var str = reader.GetString(index).Trim();
            if (string.IsNullOrEmpty(str))
            {
                return PlaceHodler;
            }
            else
            {
                return int.Parse(str).ToString();
            }
        }

        return PlaceHodler;
    }

    static string ReadDouble(IExcelDataReader reader, int index)
    {
        if (reader.IsDBNull(index))
        {
            return "0";
        }
        else if (reader.GetFieldType(index) == typeof(double))
        {
            return reader.GetDouble(index).ToString();
        }
        else if (reader.GetFieldType(index) == typeof(int))
        {
            return reader.GetInt32(index).ToString();
        }
        else if (reader.GetFieldType(index) == typeof(string))
        {
            var str = reader.GetString(index).Trim();
            if (string.IsNullOrEmpty(str))
            {
                return PlaceHodler;
            }
            else
            {
                return double.Parse(str).ToString();
            }
        }
        return PlaceHodler;
    }

    static string ReadString(IExcelDataReader reader, int index)
    {
        if (reader.IsDBNull(index))
        {
            return PlaceHodler;
        }
        else if (reader.GetFieldType(index) == typeof(double))
        {
            return $"\"{reader.GetDouble(index)}\"";
        }
        else if (reader.GetFieldType(index) == typeof(int))
        {
            return $"\"{reader.GetInt32(index)}\"";
        }
        else if (reader.GetFieldType(index) == typeof(string))
        {
            return $"\"{reader.GetString(index)}\"";
        }
        else if (reader.GetFieldType(index) == typeof(DateTime))
        {
            return $"\"{reader.GetDateTime(index).ToString()}\"";
        }
        return PlaceHodler;
    }

    static string ReadArray(IExcelDataReader reader, int index)
    {
        if (reader.IsDBNull(index))
        {
            return "";
        }
        else if (reader.GetFieldType(index) == typeof(double))
        {
            return reader.GetDouble(index).ToString();
        }
        else if (reader.GetFieldType(index) == typeof(int))
        {
            return reader.GetInt32(index).ToString();
        }
        else if (reader.GetFieldType(index) == typeof(string))
        {
            return reader.GetString(index);
        }
        return "";
    }
    #endregion

    //export all
    public static bool Export(string excelPath, string outDir)
    {
        if (File.Exists(excelPath))
        {
            return ExportSingleLua(excelPath, outDir);
        }
        else if (Directory.Exists(excelPath))
        {
            bool ret = true;
            var files = Directory.GetFiles(excelPath, "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".xls") || s.EndsWith(".xlsx"));
            foreach(var file in files)
            {
                if (!ExportSingleLua(file, outDir))
                {
                    ret = false;
                }
            }

            return ret;
        }

        return false;
    }

    
    public static bool ExportSingleLua(string excelFile, string outDir)
    {
        //read file
        var extension = Path.GetExtension(excelFile);
        IExcelDataReader reader = null;
        if (extension.Equals(".xls"))
        {
            FileStream stream = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.Read);
            reader = ExcelReaderFactory.CreateBinaryReader(stream);
        }
        else if (extension.Equals(".xlsx"))
        {
            FileStream stream = File.Open(excelFile, FileMode.Open, FileAccess.Read, FileShare.Read);
            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        }
        else
        {
            Console.WriteLine($"{excelFile} is not Excel file");
            return false ;
        }

        var configName = reader.Name;
        var writer = File.CreateText(Path.Combine(outDir, configName + ".lua"));

        //parse file
        var fieldList = new List<LuaField>();
        var typeList = new List<EType>(); 
        int line = 0;
        bool moreField = false;
        var allTables = new Dictionary<string, Value>();    //
        var contentSB = new StringBuilder();
        var content = new List<KeyValue>();

        void addToTable(Value v)
        {
            if (allTables.TryGetValue(v.StrValue, out Value value))
            {
                value.count++;
            }
            else
            {
                allTables[v.StrValue] = v;
            }
        }
        
        while(reader.Read())
        {
            ++line;
            if (line == 1)
            {
                continue;
            }
            else if (line == 2)
            {//field name
                for(var i = 0; i< reader.FieldCount; ++i)
                {
                    if(reader.IsDBNull(i))
                    {
                        continue;
                    }

                    var field = reader.GetString(i);
                    if (string.IsNullOrEmpty(field))
                    {
                        continue;
                    }

                    if (field.Contains(":c"))
                    {
                        var fieldName = field.Substring(0, field.IndexOf(':')).Trim();
                        fieldList.Add(new LuaField(i, fieldName));
                    }
                }
            }
            else if (line == 3)
            {//field type
                for (var i = 0; i < fieldList.Count; )
                {
                    var field = fieldList[i];
                    if (reader.IsDBNull(field.Key))
                    {
                        fieldList.RemoveAt(i);
                        continue;
                    }
                    else
                    {
                        ++i;
                    }

                    var type = reader.GetString(field.Key).Trim().ToLower();
                    if (type == "int" || type == "int32")
                    {
                        typeList.Add(EType.EInt);
                    }
                    else if (type == "float")
                    {
                        typeList.Add(EType.EFloat);
                    }
                    else if (type == "array")
                    {
                        typeList.Add(EType.EArray);
                    }
                    else if (type == "string")
                    {
                        typeList.Add(EType.EString);
                    }
                    else if (type == "byte")
                    {
                        typeList.Add(EType.EInt);
                    }
                    else
                    {
                        Console.WriteLine("Error: Type is not supported");
                    }
                }

                moreField = fieldList.Count > 2;
            }
            else
            {//content
                if (fieldList.Count < 2)
                {
                    return false;
                }
                
                string key = string.Empty;
                Value value = new Value();

                bool emptyContent = false;
                for(var i = 0; i < fieldList.Count; ++i)
                {
                    bool firstField = i == 0;
                    if (firstField && reader.IsDBNull(i))
                    {
                        emptyContent = true;
                        break;
                    }

                    var index = fieldList[i].Key;
                    var type = typeList[i];
                    Value element = new Value();

                    if (type == EType.EInt)
                    {
                        element.isTable = false;
                       	element.str = ReadInt(reader, index);
                    }
                    else if (type == EType.EFloat)
                    {
                        element.isTable = false;
                        element.str = ReadDouble(reader, index);
                    }
                    else if (type == EType.EString)
                    {
                        element.isTable = false;
                        element.str = ReadString(reader, index);
                    }
                    else if (type == EType.EArray)
                    {
                        element.isTable = true;
                        string array = ReadArray(reader, index);

                        if( array.Contains(";") )
                        {//;分割二维数组
                            element.isTable = true;
                            element.values = new List<Value>();

                            var array2 = array.Split(';');
                            var a2count = array2.Length;
                            if (string.IsNullOrEmpty(array2[a2count - 1]))
                            {
                                a2count--;
                            }

                            for(var a2i = 0; a2i < a2count; ++a2i)
                            {
                                var v2 = new Value()
                                {
                                    isTable = true,
                                    str = GetArrayString(array2[a2i]),
                                };
                                addToTable(v2);
                                element.values.Add(v2);
                            }

                            addToTable(element);
                        }
                        else
                        {//,分割一维数组
                            element.isTable = true;
                            element.str = GetArrayString(array);
                            addToTable(element);
                        }
                    }

                    if (moreField)
                    {//cache table
                        if (firstField)
                        {
                            key = element.str;
                            value.isTable = true;
                            value.values = new List<Value>();
                        }
                        else
                        {
                            value.values.Add(element);  
                        }
                    }
                    else
                    {
                        if (firstField)
                        {
                            key = element.str;
                        }
                        else
                        {
                            value.isTable = element.isTable;
                            value.values = element.values;
                            value.str = element.str;
                        }
                    }
                }

                if (emptyContent)
                {
                    continue;
                }

                var kv = new KeyValue()
                {
                    key = key,
                    value = value,
                };

                content.Add(kv);
            }
        }

        //去重
        Dictionary<string, string> tableNames = new Dictionary<string, string>();
        var t1sb = new StringBuilder(); var t1count = 0;
        var t2sb = new StringBuilder(); var t2count = 0;
        var t3sb = new StringBuilder(); var t3count = 0;
        void GeneralLocalTable(Value value, StringBuilder sb, int depth)
        {
            if (value.values != null)
            {//t2 or t3
                var ssb = new StringBuilder();
                for (int i = 0; i < value.values.Count; ++i)
                {
                    GeneralLocalTable(value.values[i], ssb, depth - 1);
                    if (i < value.values.Count - 1)
                    {
                        ssb.Append(",");
                    }
                }

                if (allTables.TryGetValue(value.StrValue, out Value rv) && rv.count > 0)
                {//same table
                    if (!tableNames.TryGetValue(value.StrValue, out string tableName))
                    {
                        var tsb = (depth == 3 ? t3sb : t2sb);
                        tableName = $"t{depth}[{(depth == 3 ? ++t3count : ++t2count)}]";
                        tableNames[value.StrValue] = tableName;
                        tsb.AppendLine($"{{{ssb.ToString()}}}, ");
                    }
                    sb?.Append($"{tableName}");
                }
                else
                {
                    sb?.Append($"{{{ssb.ToString()}}}");
                }
            }
            else
            {//t1
                if (value.isTable)
                {
                    if (allTables.TryGetValue(value.StrValue, out Value rv) && rv.count > 0)
                    {//same table
                        if (!tableNames.TryGetValue(value.StrValue, out string tableName))
                        {
                            tableName = $"t1[{++t1count}]";
                            tableNames[value.StrValue] = tableName;
                            t1sb.AppendLine($"{{{value.str}}}, ");
                        }
                        sb?.Append($"{tableName}");
                    }
                    else
                    {
                        sb?.Append($"{{{value.str}}}");
                    }
                }
                else
                {
                    sb?.Append($"{value.str}");
                }
            }
        }
        foreach (var each in content)
        {
            //t3
            GeneralLocalTable(each.value, null, 3);
        }

        //key map
        if (moreField)
        {
            var keyMapSB = new StringBuilder();
            for (var i = 0; i < fieldList.Count; ++i)
            {
                keyMapSB.AppendLine($"  {fieldList[i].Value} = {(i == 0 ? fieldList.Count : i)},");
            }
            writer.WriteLine($"local KeyMap = {{\n{keyMapSB.ToString()}}}\n");
        }

        //repeat table
        if (t1sb.Length > 0)
        {
            writer.WriteLine($@"local t1 = {{
{t1sb.ToString()}}}");
        }
        if (t2sb.Length > 0)
        {
            writer.WriteLine($@"local t2 = {{
{t2sb.ToString()}}}");
        }
        if (t3sb.Length > 0)
        {
            writer.WriteLine($@"local t3 = {{
{t3sb.ToString()}}}");
        }

        //begin content
        contentSB.AppendLine();
        contentSB.AppendLine($"{configName} = {{");

        void GenearlContent(Value v, StringBuilder sb)
        {
            if (v.isTable)
            {
                if (tableNames.TryGetValue(v.StrValue, out string tableName))
                {
                    sb.Append($"{tableName}");
                }
                else
                {
                    if (v.values != null)
                    {
                        sb.Append("{");
                        for (int i = 0; i < v.values.Count; ++i)
                        {
                            GenearlContent(v.values[i], sb);
                            if (i < v.values.Count - 1)
                            {
                                sb.Append(",");
                            }
                        }
                        sb.Append("}");
                    }
                    else
                    {
                        sb.Append($"{{{v.str}}}");
                    }
                }
            }
            else
            {
                sb.Append($"{v.str}");
            }
        }
        foreach(var c in content)
        {
            contentSB.Append($" [{c.key}] = ");
            var sb = new StringBuilder();
            GenearlContent(c.value, sb);
            sb.Append(",");

            contentSB.AppendLine(sb.ToString());
        }

        //end content 
        contentSB.AppendLine("}");
        contentSB.AppendLine();

        writer.WriteLine(contentSB.ToString());

        if (moreField)
        {
#if false   //set meta table
            writer.WriteLine($@"do
    local base = {{
        __index = function(table, key)
            local keyIndex = KeyMap[key]
            if not keyIndex then
                print('key not found: ', key)
                return nil
            end
            return table[keyIndex]
        end,
        __newindex = function()
            ---禁止修改只读表
            error('Forbid to modify read - only table')
        end
    }}
    for k, v in pairs({configName}) do
        v[{fieldList.Count}] = k
        setmetatable(v, base)
    end
    base.__metatable = false ---不让外面获取到元表，防止被无意修改
end");

#else       // set config table

            writer.WriteLine($@"
do
    for k, v in pairs({configName}) do
        v[{fieldList.Count}] = k
        setconfigtable(v, KeyMap)
    end
end");
#endif
        }

        writer.WriteLine();
        writer.WriteLine($"function Get{configName}(id)");
        writer.WriteLine($"  return {configName}[id]");
        writer.WriteLine($"end");

        writer.WriteLine();
        writer.WriteLine($"return {configName}");

        reader.Close();
        writer.Close();

        Console.WriteLine($"General {configName} is OK");

        return true;
    }
}