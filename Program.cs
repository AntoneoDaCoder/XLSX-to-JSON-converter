using System.Reflection.Emit;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;
using OfficeOpenXml;

namespace Excel_to_JSON_converter
{
    internal class Program
    {

        enum TarCodes
        {
            Tar4 = 1, Tar5 = 2, Tar7 = 3, Tar4Sum = 4, Tar5Sum = 5, Tar7Sum = 6, Tar14Sumt = 7, Tar14Sumdiv = 8
        }
        enum Mode
        {
            Create, Read
        }
        struct TarContainer
        {
            private Tar14? tar14;
            private Tar4? tar4;
            private Tar5? tar5;
            private Tar7? tar7;
            public Tar4? Tar4 { get { return tar4; } set { tar4 = value; } }
            public Tar7? Tar7 { get { return tar7; } set { tar7 = value; } }
            public Tar5? Tar5 { get { return tar5; } set { tar5 = value; } }
            public Tar14? Tar14 { get { return tar14; } set { tar14 = value; } }
            public TarContainer()
            {
                tar4 = null;
                tar7 = null;
                tar5 = null;
                tar14 = null;
            }
            public void ParseTars(PckAgent agent, byte month, string key)
            {
                if (tar4 is not null)
                {
                    agent.docagentMap[key].tar4 ??= new LinkedList<Tar4>();
                    tar4.nmonth = month;
                    agent?.docagentMap[key]?.tar4?.AddLast(tar4);
                }
                if (tar5 is not null)
                {
                    agent.docagentMap[key].tar5 ??= new LinkedList<Tar5>();
                    tar5.nmonth = month;
                    agent?.docagentMap[key]?.tar5?.AddLast(tar5);
                }
                if (tar7 is not null)
                {
                    agent.docagentMap[key].tar7 ??= new LinkedList<Tar7>();
                    tar7.nmonth = month;
                    agent?.docagentMap[key]?.tar7?.AddLast(tar7);
                }
                if (tar14 is not null)
                {
                    agent.docagentMap[key].tar14 ??= new LinkedList<Tar14>();
                    tar14.nmonth = month;
                    agent?.docagentMap[key]?.tar14?.AddLast(tar14);
                }
            }
        }
        static bool ReadTarCodes(Dictionary<TarCodes, LinkedList<int>> codes)
        {
            using (StreamReader reader = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "tar_codes.txt"))
            {
                int i = 0;
                try
                {
                    string? line = "";
                    while ((line = reader.ReadLine()) is not null)
                    {
                        int code = int.Parse(line.Split(":")[0]);
                        if (!Enum.IsDefined(typeof(TarCodes), code))
                        {
                            Console.WriteLine($"Error. Unknown key code at line {i} (1st char).");
                            reader.Close();
                            return false;
                        }
                        string?[] data = line.Split(":")[1].Split(",");
                        foreach (string? d in data)
                            codes[(TarCodes)code].AddLast(int.Parse(d));
                        i++;
                    }
                }
                catch (IOException e)
                {
                    Console.WriteLine("File I/O error: " + e.Message);
                    reader.Close();
                    return false;
                }
                catch (FormatException e)
                {
                    Console.WriteLine($"File data error at line {i}: " + e.Message);
                    reader.Close();
                    return false;
                }
                reader.Close();
            }
            return true;
        }
        static bool GetPckAgentInfo(PckAgentInfo pck, string? docPath)
        {
            DateTime currDate = DateTime.Now;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(docPath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[2];
                int rowCount = worksheet.Dimension.Rows, colCount = worksheet.Dimension.Columns;
                pck = new PckAgentInfo(currDate);
                try
                {
                    for (int row = 2; row < rowCount + 1; row++)
                        for (int j = 1; j < colCount + 1; j++)
                        {
                            string? field = worksheet.Cells[1, j].Value.ToString();
                            if (worksheet.Cells[row, j].Value is not null && field is not null)
                            {
                                if (field == "nmns" || field == "nmnsf" || field == "ntype" || field == "ngod" || field == "ndepno")
                                    pck.AdditionalFields[field] = int.Parse(worksheet.Cells[row, j].Value.ToString());
                                else
                                    pck.AdditionalFields[field] = worksheet.Cells[row, j].Value.ToString();
                            }
                        }
                }
                catch (FormatException e)
                {
                    Console.WriteLine("File data error: " + e.Message);
                    package.Dispose();
                    return false;
                }
                catch (IOException e)
                {
                    Console.WriteLine("File I/O error: " + e.Message);
                    package.Dispose();
                    return false;
                }
                package.Dispose();
            }
            return true;
        }
        static string? GetPath(Mode fileMode)
        {
            string? output = "";
            bool isCorrect;
            do
            {
                isCorrect = true;
                Console.WriteLine("Enter the path to the file ");
                output = Console.ReadLine();
                try
                {
                    switch (fileMode)
                    {
                        case Mode.Read:
                            using (StreamReader reader = new StreamReader(output))
                            {
                                reader.ReadLine();
                                reader.Close();
                            }
                            break;
                        case Mode.Create:
                            using (StreamWriter writer = new StreamWriter(output))
                            {
                                writer.WriteLine();
                                writer.Close();
                            }
                            break;
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine("File error: " + e.Message);
                    isCorrect = false;
                }
            } while (!isCorrect);
            return output;
        }

        static bool ReadDocAgentTable(PckAgent pckAgent, string docPath)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(docPath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows, colCount = worksheet.Dimension.Columns;
                if (rowCount == 0)
                {
                    Console.WriteLine("Error. Empty docagent spreadsheet");
                    package.Dispose();
                    return false;
                }
                for (int row = 2; row < rowCount + 1; row++)
                {
                    DocAgentInfo newAgentInfo = new();
                    int i = 1;
                    try
                    {
                        for (int j = 1; j < 5; i++, j++)
                        {
                            string? field = worksheet.Cells[1, j].Value.ToString();
                            if (worksheet.Cells[row, j].Value is not null && field is not null)
                                newAgentInfo.info[field] = worksheet.Cells[row, j].Value.ToString();
                        }
                        DocAgent newAgent = new DocAgent(newAgentInfo);
                        for (int j = 5; j < colCount + 1; j++)
                        {
                            string? field = worksheet.Cells[1, j].Value.ToString();
                            if (worksheet.Cells[row, j].Value is not null && field is not null)
                                newAgent.AdditionalFields[field] = double.Parse(worksheet.Cells[row, j].Value.ToString());
                        }
                        pckAgent.docagentMap.Add(worksheet.Cells[row, 1].Value.ToString(), newAgent);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Docagent data spreadsheet error (row: {row}, column: {i}): {e.Message}");
                        package.Dispose();
                        return false;
                    }
                }
                package.Dispose();
            }
            return true;
        }

        static TarCodes CheckCode(int code, Dictionary<TarCodes, LinkedList<int>> codes)
        {
            foreach (TarCodes tarCode in codes.Keys)
                if (codes[tarCode].Contains(code))
                    return tarCode;
            return 0;
        }

        static bool ReadSpreadSheet(string docPath, PckAgent agent, Dictionary<TarCodes, LinkedList<int>> tarCodes)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(docPath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows, colCount = worksheet.Dimension.Columns;
                if (rowCount == 0)
                {
                    Console.WriteLine("Error. Empty data spreadsheet");
                    package.Dispose();
                    return false;
                }
                for (int i = 2; i < rowCount + 1; i++)
                {
                    TarContainer container = new TarContainer();
                    byte month = byte.Parse(worksheet.Cells[i, 2].Value.ToString());
                    for (int j = 1; j < colCount + 1; j++)
                    {
                        int ncode;
                        if (worksheet.Cells[i, j].Value is null)
                            continue;
                        if (int.TryParse(worksheet.Cells[1, j].Value.ToString(), out ncode))
                        {
                            double temp = double.Parse(worksheet.Cells[i, j].Value.ToString());
                            switch (CheckCode(ncode, tarCodes))
                            {
                                case TarCodes.Tar4:
                                    container.Tar4 ??= new Tar4();
                                    container.Tar4.tar4sum.AddLast(new Dictionary<string, double>() { { "ncode", ncode }, { "nsum", temp } });
                                    break;
                                case TarCodes.Tar5:
                                    container.Tar5 ??= new Tar5();
                                    container.Tar5.tar5sum.AddLast(new Dictionary<string, double>() { { "ncode", ncode }, { "nsum", temp } });
                                    break;
                                case TarCodes.Tar7:
                                    container.Tar7 ??= new Tar7();
                                    container.Tar7.tar7sum.AddLast(new Dictionary<string, double>() { { "ncode", ncode }, { "nsumv", temp } });
                                    break;
                                case TarCodes.Tar4Sum:
                                    container.Tar4 ??= new Tar4();
                                    container.Tar4.nsummonth = temp;
                                    break;
                                case TarCodes.Tar5Sum:
                                    container.Tar5 ??= new Tar5();
                                    container.Tar5.nsummonth = temp;
                                    break;
                                case TarCodes.Tar7Sum:
                                    container.Tar7 ??= new Tar7();
                                    container.Tar7.nsummonth = temp;
                                    break;
                                case TarCodes.Tar14Sumt:
                                    container.Tar14 ??= new Tar14();
                                    container.Tar14.nsumt = temp;
                                    break;
                                case TarCodes.Tar14Sumdiv:
                                    container.Tar14 ??= new Tar14();
                                    container.Tar14.nsumdiv = temp;
                                    break;
                            }
                        }
                    }
                    container.ParseTars(agent, month, worksheet.Cells[i, 1].Value.ToString());
                }
                package.Dispose();
            }
            return true;
        }

        static void WriteJSONStringToFile(string? jsonString, Mode mode)
        {
            string? path = GetPath(mode);
            using (StreamWriter writer = new StreamWriter(path))
            {
                writer.Write(jsonString);
                writer.Close();
            }
            Console.WriteLine($"JSON file created, {path}");
        }
        static void Main()
        {
            Dictionary<TarCodes, LinkedList<int>> tarCodes = new Dictionary<TarCodes, LinkedList<int>>();
            tarCodes.Add(TarCodes.Tar4, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar5, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar7, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar4Sum, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar5Sum, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar7Sum, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar14Sumdiv, new LinkedList<int>());
            tarCodes.Add(TarCodes.Tar14Sumt, new LinkedList<int>());
            var docPath = GetPath(Mode.Read);
            if (!ReadTarCodes(tarCodes))
            {
                Console.WriteLine("An error occurred while reading the tar_codes file.");
                return;
            }
            PckAgentInfo? agentInfo = null;
            if (!GetPckAgentInfo(agentInfo, docPath))
            {
                Console.WriteLine("An error occurred while reading the pckagentinfo file.");
                return;
            }
            PckAgent agent = new(agentInfo);
            if (!ReadDocAgentTable(agent, docPath))
            {
                Console.WriteLine("An error occurred while reading the docagent spreadsheet.");
                return;
            }
            if (!ReadSpreadSheet(docPath, agent, tarCodes))
            {
                Console.WriteLine("An error occurred while reading the data spreadsheet.");
                return;
            }
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            Wrapper wrapper = new(agent);
            string jsonString = JsonSerializer.Serialize(wrapper, options);
            Console.Write("Serialization completed.");
            WriteJSONStringToFile(jsonString, Mode.Create);
        }
    }
}