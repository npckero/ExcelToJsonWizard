using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using System.Text.Json;
using System.Text.Encodings.Web;
using System.Text.Unicode;

namespace ExcelToJsonWizard
{
    class Program
    {
        static Dictionary<string, Dictionary<string, int>> enumMappings;
        static bool emptyValueDetectedInFile = false;  // 빈 값이 감지되었는지 여부를 추적하는 변수

        static void Main()
        {
            // 실행 파일의 현재 디렉토리 가져오기
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            // 설정 파일을 실행 파일 위치에서 찾도록 수정
            string configFilePath = Path.Combine(basePath, "config.txt");
            var config = LoadConfiguration(configFilePath);


            string defaultExcelDirectoryPath = "excel_files";
            string defaultLoaderOutputDirectory = "loader_output";
            string defaultJsonOutputDirectory = "json_output";

            bool useAbsolutePath = config.ContainsKey("useAbsolutePath") && config["useAbsolutePath"].ToLower() == "true";
            bool allowMultipleSheets = config.ContainsKey("allowMultipleSheets") && config["allowMultipleSheets"].ToLower() == "true";
            bool useResources = config.ContainsKey("useResources") && config["useResources"].ToLower() == "true";
            string resourcesInternalPath = config.ContainsKey("resourcesInternalPath") ? config["resourcesInternalPath"] : "default/path";

            string excelDirectoryPath = ValidateOrCreateDirectory(config, "excelDirectoryPath", defaultExcelDirectoryPath, useAbsolutePath, basePath);
            string loaderOutputDirectory = ValidateOrCreateDirectory(config, "loaderOutputDirectory", defaultLoaderOutputDirectory,useAbsolutePath, basePath);
            string jsonOutputDirectory = ValidateOrCreateDirectory(config, "jsonOutputDirectory", defaultJsonOutputDirectory, useAbsolutePath, basePath);
            string logFilePath = Path.Combine("log", $"{DateTime.Now:yyyy-MM-dd}_error_log.txt");

            if (!Directory.Exists("log"))
            {
                Directory.CreateDirectory("log");
            }

            // Load Enum definitions and generate Enum C# file
            enumMappings = LoadEnumDefinitionsAndGenerateCs(excelDirectoryPath, loaderOutputDirectory, logFilePath);

            ProcessExcelFiles(excelDirectoryPath, loaderOutputDirectory, jsonOutputDirectory, logFilePath, allowMultipleSheets, useResources, resourcesInternalPath);

            // 프로그램 완료 후 콘솔 창을 유지합니다.
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
        }

        static Dictionary<string, string> LoadConfiguration(string configFilePath)
        {
            var config = new Dictionary<string, string>();

            if (File.Exists(configFilePath))
            {
                var lines = File.ReadAllLines(configFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#"))
                    {
                        continue; // 주석이거나 빈 줄은 무시
                    }

                    var cleanLine = line.Split('#')[0].Trim(); // 주석 부분을 제거하고 정리
                    var parts = cleanLine.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        config[parts[0].Trim()] = parts[1].Trim();
                    }
                }
            }
            else
            {
                using (var sw = File.CreateText(configFilePath))
                {
                    sw.WriteLine("# 사용자 지정 디렉토리 설정");
                    sw.WriteLine("excelDirectoryPath=excel_files # 엑셀 파일 디렉토리 경로");
                    sw.WriteLine("loaderOutputDirectory=loader_output # 로더 클래스 출력 디렉토리 경로");
                    sw.WriteLine("jsonOutputDirectory=json_output # JSON 파일 출력 디렉토리 경로");
                    sw.WriteLine();

                    sw.WriteLine("# 경로 타입 선택");
                    sw.WriteLine("useAbsolutePath = false # 절대 경로를 사용할지 여부 (true: 절대 경로, false: 실행 파일 기준 상대 경로)");
                    sw.WriteLine();

                    sw.WriteLine("# 다중 시트 설정");
                    sw.WriteLine("allowMultipleSheets=false # 다중 시트를 허용할지 여부 (true/false)");
                    sw.WriteLine();

                    sw.WriteLine("# Resources 사용 설정");
                    sw.WriteLine("useResources=true # Resources 폴더 사용 여부 (true/false)");
                    sw.WriteLine("resourcesInternalPath=JSON # Resources 내부 경로");
                }

                config["excelDirectoryPath"] = "excel_files";
                config["loaderOutputDirectory"] = "loader_output";
                config["jsonOutputDirectory"] = "json_output";
                config["useAbsolutePath"] = "false";
                config["allowMultipleSheets"] = "false";
                config["useResources"] = "true";
                config["resourcesInternalPath"] = "JSON";
            }

            return config;
        }

        static string ValidateOrCreateDirectory(Dictionary<string, string> config, string key, string defaultDirectory, bool useAbsolutePath, string basePath)
        {
            string path;
            if (config.ContainsKey(key))
            {
                path = config[key];

                // 상대 경로 사용 옵션이 꺼져 있다면 실행 파일 경로를 기준으로 변환
                if (!useAbsolutePath && !Path.IsPathRooted(path))
                {
                    path = Path.Combine(basePath, path);
                }
            }
            else
            {
                path = Path.Combine(basePath, defaultDirectory);
            }

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                Console.WriteLine($"Created directory: {path}");
            }
            return path;
        }

        static Dictionary<string, Dictionary<string, int>> LoadEnumDefinitionsAndGenerateCs(string excelDir, string loaderDir, string logFilePath)
        {
            var enumDefinitions = new Dictionary<string, Dictionary<string, int>>();
            var enumFilePath = Path.Combine(excelDir, "Enum.xlsx");

            if (File.Exists(enumFilePath))
            {
                try
                {
                    var sb = new StringBuilder();
                    sb.AppendLine("using System;");
                    sb.AppendLine();
                    sb.AppendLine("public static class DesignEnums");
                    sb.AppendLine("{");

                    using (var workbook = new XLWorkbook(enumFilePath))
                    {
                        var worksheet = workbook.Worksheet(1);

                        foreach (var row in worksheet.RowsUsed())
                        {
                            var enumName = row.Cell(1).GetValue<string>();
                            if (enumDefinitions.ContainsKey(enumName))
                            {
                                throw new Exception($"Duplicate enum name '{enumName}' found in Enum definitions.");
                            }

                            var enumValues = new Dictionary<string, int>();
                            int index = 0;

                            sb.AppendLine($"    public enum {enumName}");
                            sb.AppendLine("    {");

                            for (int col = 2; col <= row.LastCellUsed().Address.ColumnNumber; col++)
                            {
                                var value = row.Cell(col).GetValue<string>();
                                if (enumValues.ContainsKey(value))
                                {
                                    throw new Exception($"Duplicate value '{value}' found in enum '{enumName}'.");
                                }

                                enumValues[value] = index;
                                sb.AppendLine($"        {value} = {index},");
                                index++;
                            }

                            sb.AppendLine("    }");
                            enumDefinitions[enumName] = enumValues;
                        }
                    }

                    sb.AppendLine("}");

                    var enumOutputPath = Path.Combine(loaderDir, "DesignEnums.cs");
                    File.WriteAllText(enumOutputPath, sb.ToString());
                    Console.WriteLine($"Enum definitions file generated\n");
                }
                catch (Exception ex)
                {
                    LogError(logFilePath, $"Error loading Enum definitions from file {enumFilePath}: {ex.Message}\n{ex.StackTrace}");
                    Console.WriteLine($"Error loading Enum definitions from file {enumFilePath}: {ex.Message}");
                }
            }

            return enumDefinitions;
        }

        static void ProcessExcelFiles(string excelDir, string loaderDir, string jsonDir, string logFilePath, bool allowMultipleSheets, bool useResources, string resourcesInternalPath)
        {
            var excelFiles = Directory.GetFiles(excelDir, "*.xlsx");

            int totalFiles = excelFiles.Length;
            int errorFiles = 0;
            int processedFiles = 0;

            foreach (var excelFilePath in excelFiles)
            {
                if (Path.GetFileName(excelFilePath).StartsWith("~") || Path.GetFileName(excelFilePath).Equals("Enum.xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Skipping file: {excelFilePath}\n");
                    continue;
                }

                bool success = GenerateClassAndJsonFromExcel(excelFilePath, loaderDir, jsonDir, logFilePath, allowMultipleSheets, useResources, resourcesInternalPath);
                if (success)
                {
                    processedFiles++;
                }
                else
                {
                    errorFiles++;
                }
            }

            Console.WriteLine($"Total files processed: {totalFiles}");
            Console.WriteLine($"Successfully processed files: {processedFiles}");
            Console.WriteLine($"Files with errors: {errorFiles}");
        }

        static bool GenerateClassAndJsonFromExcel(string excelPath, string loaderDir, string jsonDir, string logFilePath, bool allowMultipleSheets, bool useResources, string resourcesInternalPath)
        {
            try
            {
                emptyValueDetectedInFile = false;  // 새로운 파일을 처리할 때마다 초기화

                using (var workbook = new XLWorkbook(excelPath))
                {
                    IEnumerable<IXLWorksheet> sheets = allowMultipleSheets ? workbook.Worksheets : new[] { workbook.Worksheet(1) };

                    foreach (var worksheet in sheets)
                    {
                        try
                        {
                            var rows = worksheet.RowsUsed();

                            if (!rows.Any())
                            {
                                continue;
                            }

                            var className = allowMultipleSheets ?
                                MakeValidClassName($"{Path.GetFileNameWithoutExtension(excelPath)}_{worksheet.Name}") :
                                MakeValidClassName(Path.GetFileNameWithoutExtension(excelPath));
                            var sb = new StringBuilder();

                            sb.AppendLine("using System;");
                            sb.AppendLine("using System.Collections.Generic;");
                            sb.AppendLine("using System.IO;");
                            sb.AppendLine("using UnityEngine;");
                            sb.AppendLine();
                            sb.AppendLine($"[Serializable]");
                            sb.AppendLine($"public class {className}");
                            sb.AppendLine("{");

                            var headers = worksheet.Row(1).Cells();
                            var types = worksheet.Row(2).Cells();
                            var descriptions = worksheet.Row(3).Cells();

                            int validColumnCount = 0;
                            for (int i = 0; i < headers.Count(); i++)
                            {
                                string header = headers.ElementAt(i).GetString();
                                string type = types.ElementAt(i).GetString();

                                // 빈 셀이 하나라도 나오면 중단
                                if (string.IsNullOrWhiteSpace(header) || string.IsNullOrWhiteSpace(type))
                                {
                                    break;
                                }

                                validColumnCount++;
                            }

                            if (!headers.ElementAt(0).GetString().Equals("key", StringComparison.OrdinalIgnoreCase))
                            {
                                throw new Exception("The first column must be 'key'.");
                            }

                            if (!types.ElementAt(0).GetString().Equals("int", StringComparison.OrdinalIgnoreCase))
                            {
                                throw new Exception("The type of the first column must be 'int'.");
                            }

                            for (int i = 0; i < validColumnCount; i++)
                            {
                                var variableName = headers.ElementAt(i).GetString();
                                var dataType = types.ElementAt(i).GetString();
                                var description = descriptions.ElementAtOrDefault(i)?.GetString() ?? "No description provided.";

                                if ((dataType.StartsWith("Enum<") || dataType.StartsWith("List<Enum<")) && enumMappings == null)
                                {
                                    throw new Exception($"Enum definition file not found, but type {dataType} requires it.");
                                }

                                if (dataType.StartsWith("Enum<"))
                                {
                                    var enumTypeName = dataType.Substring(5, dataType.Length - 6);
                                    if (!enumMappings.ContainsKey(enumTypeName))
                                    {
                                        throw new Exception($"Enum type '{enumTypeName}' not found in Enum definitions.");
                                    }
                                    dataType = $"DesignEnums.{enumTypeName}";
                                }

                                if (dataType.StartsWith("List<Enum<"))
                                {
                                    var enumTypeName = dataType.Substring(10, dataType.Length - 12);
                                    if (!enumMappings.ContainsKey(enumTypeName))
                                    {
                                        throw new Exception($"Enum type '{enumTypeName}' not found in Enum definitions.");
                                    }
                                    dataType = $"List<DesignEnums.{enumTypeName}>";
                                }

                                sb.AppendLine($"    /// <summary>");
                                sb.AppendLine($"    /// {description}");
                                sb.AppendLine($"    /// </summary>");
                                sb.AppendLine($"    public {dataType} {variableName};");
                                sb.AppendLine();
                            }

                            sb.AppendLine("}");

                            sb.AppendLine($"public class {className}Loader");
                            sb.AppendLine("{");
                            sb.AppendLine($"    public List<{className}> ItemsList {{ get; private set; }}");
                            sb.AppendLine($"    public Dictionary<int, {className}> ItemsDict {{ get; private set; }}");
                            sb.AppendLine();
                            if (useResources)
                            {
                                sb.AppendLine($"    public {className}Loader(string path = \"{resourcesInternalPath}/{className}\")");
                            }
                            else
                            {
                                sb.AppendLine($"    public {className}Loader(string path)");
                            }
                            sb.AppendLine("    {");
                            sb.AppendLine("        string jsonData;");
                            if (useResources)
                            {
                                sb.AppendLine("        jsonData = Resources.Load<TextAsset>(path).text;");
                            }
                            else
                            {
                                sb.AppendLine("        jsonData = File.ReadAllText(path);");
                            }
                            sb.AppendLine("        ItemsList = JsonUtility.FromJson<Wrapper>(jsonData).Items;");
                            sb.AppendLine($"        ItemsDict = new Dictionary<int, {className}>();");
                            sb.AppendLine("        foreach (var item in ItemsList)");
                            sb.AppendLine("        {");
                            sb.AppendLine($"            ItemsDict.Add(item.key, item);");
                            sb.AppendLine("        }");
                            sb.AppendLine("    }");
                            sb.AppendLine();
                            sb.AppendLine($"    [Serializable]");
                            sb.AppendLine($"    private class Wrapper");
                            sb.AppendLine("    {");
                            sb.AppendLine($"        public List<{className}> Items;");
                            sb.AppendLine("    }");
                            sb.AppendLine();

                            // GetByKey 메서드 추가
                            sb.AppendLine($"    public {className} GetByKey(int key)");
                            sb.AppendLine("    {");
                            sb.AppendLine("        if (ItemsDict.ContainsKey(key))");
                            sb.AppendLine("        {");
                            sb.AppendLine("            return ItemsDict[key];");
                            sb.AppendLine("        }");
                            sb.AppendLine("        return null;");
                            sb.AppendLine("    }");

                            // GetByIndex 메서드 추가
                            sb.AppendLine($"    public {className} GetByIndex(int index)");
                            sb.AppendLine("    {");
                            sb.AppendLine("        if (index >= 0 && index < ItemsList.Count)");
                            sb.AppendLine("        {");
                            sb.AppendLine("            return ItemsList[index];");
                            sb.AppendLine("        }");
                            sb.AppendLine("        return null;");
                            sb.AppendLine("    }");
                            sb.AppendLine("}");

                            var jsonArray = new List<Dictionary<string, object>>();
                            var keySet = new HashSet<int>();



                            for (int i = 4; i <= worksheet.LastRowUsed().RowNumber(); i++)
                            {
                                var row = worksheet.Row(i);
                                var rowDict = new Dictionary<string, object>();

                                for (int j = 0; j < validColumnCount; j++)
                                {
                                    var variableName = headers.ElementAt(j).GetString();
                                    var dataType = types.ElementAt(j).GetString();
                                    var cellValue = row.Cell(j + 1).GetValue<string>();

                                    

                                    var convertedValue = ConvertToType(cellValue, dataType, variableName, logFilePath, excelPath, worksheet.Name);

                                    if (variableName == "key" && !keySet.Add((int)convertedValue))
                                    {
                                        throw new Exception($"Duplicate key value '{convertedValue}' found in sheet '{worksheet.Name}' of file '{excelPath}'");
                                    }

                                    rowDict[variableName] = convertedValue;
                                }

                                jsonArray.Add(rowDict);
                            }

                            var classCode = sb.ToString();
                            var loaderOutputPath = Path.Combine(loaderDir, $"{className}.cs");
                            File.WriteAllText(loaderOutputPath, classCode);
                            Console.WriteLine($"Class file generated at {className}");

                            var jsonOutputPath = Path.Combine(jsonDir, $"{className}.json");
                            var wrapper = new { Items = jsonArray };
                            var jsonData = JsonSerializer.Serialize(wrapper, new JsonSerializerOptions
                            {
                                WriteIndented = true,
                                Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
                            });
                            File.WriteAllText(jsonOutputPath, jsonData);
                            Console.WriteLine($"JSON file generated at {className}\n");

                            // 파일 처리 후, 빈 값이 감지되었을 때 메시지 출력
                            if (emptyValueDetectedInFile)
                            {
                                Console.WriteLine($"Warning: Empty values detected in file '{excelPath}'. Default values were used.");
                                LogError(logFilePath, $"Empty values detected in file '{excelPath}'. Default values were used.");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogError(logFilePath, $"Error processing sheet {worksheet.Name} in file {excelPath}: {ex.Message}\n{ex.StackTrace}");
                            Console.WriteLine($"Error processing sheet {worksheet.Name} in file {excelPath}: {ex.Message}");
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                LogError(logFilePath, $"Error processing file {excelPath}: {ex.Message}\n{ex.StackTrace}");
                Console.WriteLine($"Error processing file {excelPath}: {ex.Message}");
                return false;
            }
        }

        static object ConvertToType(string value, string type, string variableName, string logFilePath, string excelPath, string sheetName)
        {
            try
            {
                // 빈 칸일 경우 기본값 설정
                if (string.IsNullOrWhiteSpace(value))
                {
                    // 파일 내에서 처음으로 빈 칸이 발견되면 플래그 설정
                    if (!emptyValueDetectedInFile)
                    {
                        emptyValueDetectedInFile = true;
                    }

                    // 리스트 타입에서 빈 칸을 허용하지 않도록 메시지 출력
                    if (type.StartsWith("List<"))
                    {
                        throw new Exception($"Empty values are not supported for list types. Variable '{variableName}' in sheet '{sheetName}' of file '{excelPath}' contains an empty value.");
                    }

                    // Enum 타입에서 빈 칸을 허용하지 않도록 메시지 출력
                    if (type.StartsWith("Enum<"))
                    {
                        throw new Exception($"Empty values are not supported for Enum types. Variable '{variableName}' in sheet '{sheetName}' of file '{excelPath}' contains an empty value.");
                    }

                    // 타입에 따른 기본값 처리
                    switch (type)
                    {
                        case "int":
                            return 0;  // 기본값 0
                        case "float":
                            return 0.0f;  // 기본값 0.0f
                        case "double":
                            return 0.0;  // 기본값 0.0
                        case "bool":
                            return false;  // 기본값 false
                        case "string":
                            return "";  // 기본값 빈 문자열
                        default:
                            throw new Exception($"Unsupported data type: {type}");
                    }
                }

                if (type.StartsWith("List<Enum<"))
                {
                    var enumTypeName = type.Substring(10, type.Length - 12);
                    if (!enumMappings.ContainsKey(enumTypeName))
                    {
                        throw new Exception($"Enum type '{enumTypeName}' not found in Enum definitions.");
                    }
                    var enumMap = enumMappings[enumTypeName];
                    return value.Split(',').Select(v => enumMap[v.Trim()]).ToList();
                }
                else if (type.StartsWith("List<"))
                {
                    var itemType = type.Substring(5, type.Length - 6);
                    if (itemType == "int")
                    {
                        return value.Split(',').Select(int.Parse).ToList();
                    }
                    else if (itemType == "float")
                    {
                        return value.Split(',').Select(float.Parse).ToList();
                    }
                    else if (itemType == "double")
                    {
                        return value.Split(',').Select(double.Parse).ToList();
                    }
                    else if (itemType == "bool")
                    {
                        return value.Split(',').Select(bool.Parse).ToList();
                    }
                    else
                    {
                        return value.Split(',').Select(v => v.Trim()).ToList();
                    }
                }
                else if (type == "int")
                {
                    return int.Parse(value);
                }
                else if (type == "float")
                {
                    return float.Parse(value);
                }
                else if (type == "double")
                {
                    return double.Parse(value);
                }
                else if (type == "bool")
                {
                    return bool.Parse(value);
                }
                else if (type.StartsWith("Enum<"))
                {
                    var enumTypeName = type.Substring(5, type.Length - 6);
                    if (!enumMappings.ContainsKey(enumTypeName))
                    {
                        throw new Exception($"Enum type '{enumTypeName}' not found in Enum definitions.");
                    }
                    var enumMap = enumMappings[enumTypeName];
                    return enumMap[value];
                }
                else if (type == "string")
                {
                    return value;
                }
                else
                {
                    throw new Exception($"Unsupported data type: {type}");
                }
            }
            catch (Exception ex)
            {
                LogError(logFilePath, $"Error converting value '{value}' for variable '{variableName}' in sheet '{sheetName}' of file '{excelPath}': {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }

        static string MakeValidClassName(string name)
        {
            var sb = new StringBuilder();
            foreach (var c in name)
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        static void LogError(string logFilePath, string message)
        {
            try
            {
                using (StreamWriter sw = File.AppendText(logFilePath))
                {
                    sw.WriteLine($"{DateTime.Now}: {message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }
    }
}
