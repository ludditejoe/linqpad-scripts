<Query Kind="Program">
  <Namespace>System.Data.OleDb</Namespace>
</Query>

void Main()
{
	// Write code to test your extensions here. Press F5 to compile and run.
}

public static class MyExtensions
{
	// Write custom extension methods here. They will be available to all queries.
}

#region StringExtensions

public static class StringOps 
{
    /// Extension for overloading Contains with a supplied StringComparison
    public static bool Contains(this string source, string toCheck, StringComparison comparer)
    {
        return source.IndexOf(toCheck, comparer) >= 0;
    }
    
    /// Implementation of case-insensitive Contains
    public static bool CiContains(this string source, string toCheck)
    {
        return Contains(source, toCheck, StringComparison.InvariantCultureIgnoreCase);
    }
    
    /// Extension for (inefficient) Regex matching, since it creates a new Regex for each string to match
    public static bool IsMatch(this string target, string regExp)
    {
        return (new Regex(regExp).IsMatch(target));
    }
}

#endregion StringExtensions

#region FileOps

public static class FileOps
{
    #region TextReader
    /// Helper function to read all lines in a textfile and return an IEnumerable of strings
    public static IEnumerable<string> ReadText(string file)
    {
        if (!File.Exists(file))
            throw new FileNotFoundException("Provided file doesn't exist!");
        
        using (var reader = new StreamReader(file))
        {
            while (reader.Peek() >= 0)
            {
                yield return reader.ReadLine() ?? String.Empty;
            }
        }
    }
    
   
    /// Helper function to read all lines in a textfile and return an IEnumerable of tuples containing line number and line text
    public static IEnumerable<Tuple<int, string>> ReadText2(string file)
    {
        if (!File.Exists(file))
            throw new FileNotFoundException("Provided file doesn't exist!");
        
        using (var reader = new StreamReader(file))
        {
            int lineNo = 0;
            while (reader.Peek() >= 0)
            {
                lineNo += 1;
                yield return Tuple.Create(lineNo, reader.ReadLine() ?? String.Empty);
            }
        }
    }
    
    /// Helper function to read char-delimited text files and return as IEnumerable of string arrays
    public static IEnumerable<string[]> ReadCsv(string file, char delimiter)
    {
        return ReadText(file).Select(x => x.Split(delimiter));
    }
    
    #endregion TextReader
    
    #region Searching
    
    public static IEnumerable<FileInfo> FindFile(string searchString, string startingFolder)
    {
        if (!Directory.Exists(startingFolder))
            throw new DirectoryNotFoundException("Provided starting folder doesn't exist!");
        
        using (var conn = new OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"))
        {
            conn.Open();

            string query = String.Format("SELECT System.FileName, System.ItemFolderPathDisplay FROM SystemIndex WHERE SCOPE='file:{0}'", startingFolder);
  
            using(var rdr = (new OleDbCommand(query, conn)).ExecuteReader())
            {
                while (rdr.Read())
                {
                    if (rdr[0].ToString().CiContains(searchString))
                        yield return new FileInfo(Path.Combine(rdr[1].ToString(), rdr[0].ToString()));
                }
                rdr.Close();
            }
            conn.Close();
        }
    }
    
    public static IEnumerable<FileInfo> FindFile(string searchString)
    {
        return FindFile(searchString, Path.GetPathRoot(Environment.SystemDirectory));
    }
    
    /// Programmatically search files for certain words or phrases using OleDB and windows search
    public static IEnumerable<Tuple<string, string, string>> Find(string searchString, string startingFolder)
    {
        if (!Directory.Exists(startingFolder))
            throw new DirectoryNotFoundException("Provided starting folder doesn't exist!");
        
        using (var conn = new OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"))
        {
            conn.Open();

            string query = String.Format("SELECT System.FileName, System.ItemFolderPathDisplay FROM SystemIndex WHERE {0} ('{1}') AND SCOPE='file:{2}'", (searchString.Contains(" ") ? "FREETEXT" : "CONTAINS"), searchString, startingFolder);
  
            using(var rdr = (new OleDbCommand(query, conn)).ExecuteReader())
            {
                while (rdr.Read())
                {
                    yield return Tuple.Create(rdr[0].ToString(), rdr[1].ToString(), String.Concat(Path.Combine(rdr[1].ToString(), rdr[0].ToString())));
                }
                rdr.Close();
            }
            conn.Close();
        }
    }
    
    // If no starting folder is supplied use system disk
    public static IEnumerable<Tuple<string, string, string>> Find(string searchString)
    {
        return Find(searchString, Path.GetPathRoot(Environment.SystemDirectory));
    }
    
    
    // Only search (or rather return) files having the supplied file extension
    public static IEnumerable<Tuple<string, string, string>> Find(string searchString, string startingFolder, string fileExtension)
    {
        return Find(searchString, startingFolder)
                .Where(x => String.IsNullOrEmpty(fileExtension) 
                            || (Path.GetExtension(x.Item1)).IndexOf(fileExtension, System.StringComparison.InvariantCultureIgnoreCase) >= 0);
    }
    
    
    /// Searches and replaces trings within a specified text file; commits changes if updateFile is true. Returns a set of changed rows before and after replacement.
    public static IEnumerable<SearchReplaceResult> SearchReplace(string filePath, string searchFor, string replaceWith, bool updateFile, bool createBackup)
    {
        var tempFile = Path.GetTempFileName();
        var encoding = Encoding.GetEncoding(1252); // defaulting to win 1252 encoding if not identified as UTF-8 or Unicode
        
        // Lazy, non-foolproof way of identifying encoding, but should catch the common cases.
        using (var reader = new StreamReader(filePath, encoding, true))
        {
            reader.Peek();
            encoding = reader.CurrentEncoding;
            reader.Close();
        }
        
        using (var writer = new StreamWriter(tempFile, false, encoding))
        {
            foreach (var line in FileOps.ReadText2(filePath))
            {
                var changedLine = line.Item2.Replace(searchFor, replaceWith);
                if (updateFile)
                    writer.WriteLine(changedLine);
                    
                yield return new SearchReplaceResult(line.Item1, line.Item2, changedLine, line.Item2.Contains(searchFor));
            }
            
            writer.Close();
        }
        
        if (updateFile)
        {
            string backupFileName = null;
            
            if (createBackup)
            {
                string backupDirectory = Path.Combine(Path.GetDirectoryName(filePath), "Backup");
                if (!Directory.Exists(backupDirectory))
                   Directory.CreateDirectory(backupDirectory);
                   
                backupFileName = Path.Combine(backupDirectory, Path.GetFileName(filePath));
            }   
            File.Replace(tempFile, filePath, backupFileName);
        }
    }
    
    public static IEnumerable<SearchReplaceResult> SearchReplace(string filePath, string searchFor, string replaceWith, bool updateFile)
    {
        return SearchReplace(filePath, searchFor, replaceWith, updateFile, false);
    }
    
    /// Readonly implementation of SearchReplace
    public static IEnumerable<SearchReplaceResult> SearchReplace(string filePath, string searchFor, string replaceWith)
    {
        return SearchReplace(filePath, searchFor, replaceWith, false, false);
    }
    
    
    /// Only does searching in files and returns only matches
    public static IEnumerable<SearchResult> Search(string filePath, string searchFor)
    {
        return SearchReplace(filePath, searchFor, searchFor)
                .Where(x => x.MatchFound)
                .Select(x => new SearchResult (x.LineNo, x.OriginalLine));
    }
    
    public static IEnumerable<SearchResult> Search(string searchFor)
    {
        return Search(Path.GetPathRoot(Environment.SystemDirectory), searchFor);
    }
    
    /// Open the first file that is found using supplied expression; paths that contain the expression are prioritized.
    public static void Open(string expression, string startingFolder)
    {
        Process.Start(FileOps.Find(expression, startingFolder)
                        .Select(x => x.Item3)
                        .OrderByDescending(x => x.Contains(expression))
                        .First());
    }
    
    public static void Open(string expression)
    {
        Open(expression, Path.GetPathRoot(Environment.SystemDirectory));
    }

    /// Executes a Find followed by a Search for occurrances of the search string in the Find result
    public static IEnumerable<Tuple<String, IEnumerable<SearchResult>>> Find2(string searchString, string startingFolder, string fileExtension)
    {
        return FileOps.Find(searchString, startingFolder, fileExtension).Select(x => Tuple.Create(x.Item3, FileOps.Search(x.Item3, searchString)));
    }
    
    public static IEnumerable<Tuple<String, IEnumerable<SearchResult>>> Find2(string searchString, string startingFolder)
    {
        return FileOps.Find(searchString, startingFolder).Select(x => Tuple.Create(x.Item3, FileOps.Search(x.Item3, searchString)));
    }
    
    public static IEnumerable<Tuple<String, IEnumerable<SearchResult>>> Find2(string searchString)
    {
        return FileOps.Find2(searchString, Path.GetPathRoot(Environment.SystemDirectory));
    }
    
    
    #endregion Searching
}

#endregion FileOps

// You can also define non-static classes, enums, etc.

public class SearchResult {
    public int      LineNo {get; private set;}
    public string   OriginalLine {get; private set;}
    
    public SearchResult(int lineNo, string originalLine)
    {
        LineNo = lineNo;
        OriginalLine = originalLine;
    }
}

public class SearchReplaceResult : SearchResult {
    public string   ChangedLine {get; private set;}
    public bool     MatchFound {get; private set;}
    
    public SearchReplaceResult(int lineNo, string originalLine,string changedLine, bool matchFound) : base(lineNo, originalLine)
    {
        ChangedLine = changedLine;
        MatchFound = matchFound;
    }
}