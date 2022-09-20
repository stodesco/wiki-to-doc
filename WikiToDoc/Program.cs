// See https://aka.ms/new-console-template for more information

using System.Diagnostics;

var arguments = Environment.GetCommandLineArgs();
if (arguments.Length != 2)
{
    Console.WriteLine("WikiToDoc expects one argument, the root path of a wiki");
    return;
}

string path = arguments[1];

Console.WriteLine($"Generating documents from wiki MD files in path {path}");


var wordDir = Path.Combine(path, "WikiToDoc");
Directory.CreateDirectory(wordDir);
foreach (var delFile in Directory.GetFiles(wordDir))
{
    File.Delete(delFile);
}
CopyFiles(Path.Combine(path, ".attachments"), Path.Combine(wordDir, ".attachments"));

string[] files = System.IO.Directory.GetFiles(path, "*.md", SearchOption.AllDirectories);
foreach (string file in files)
{
    var destPath = Path.Combine(wordDir, file.Replace(path, "").Replace('\\', '-').TrimStart('-'));
    File.Copy(file, destPath, true);
    var text = File.ReadAllText(destPath);
    text = text.Replace("![image.png](/.", "![image.png](.");
    File.WriteAllText(destPath, text);
    Console.WriteLine(file);
}

foreach (var md in Directory.GetFiles(wordDir))
{
    if (md.EndsWith(".md"))
    {
        var wordFile = md.Substring(0, md.Length - 3) + ".docx";

        Process p = new Process();

        p.StartInfo.UseShellExecute = false;
        p.StartInfo.RedirectStandardOutput = true;
        p.StartInfo.WorkingDirectory = wordDir;
        p.StartInfo.FileName = "pandoc";
        p.StartInfo.Arguments = $"-o {wordFile} -f markdown -t docx {md}";
        p.Start();
        p.WaitForExit();

        Console.WriteLine($"Successfully converted '{wordFile}'");
    }
}

static void CopyFiles(string sourcePath, string targetPath)
{
    if (!Directory.Exists(targetPath))
        Directory.CreateDirectory(targetPath);
    //Copy all the files & Replaces any files with the same name
    foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
    {
        File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
    }
}