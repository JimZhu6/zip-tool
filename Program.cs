using System;
using System.Collections.Generic;
using System.IO;
using Ionic.Zip; // 需安装DotNetZip包
using System.IO.Compression; // 用于无密码压缩
using System.Reflection;

class Program
{
  static void Main()
  {
    Console.OutputEncoding = System.Text.Encoding.UTF8;
    Console.CursorVisible = false;
    Console.Title = "文件夹压缩工具";
    // 获取并打印版本号和作者信息
    var version = Assembly.GetExecutingAssembly().GetName().Version;
    var companyAttribute = (AssemblyCompanyAttribute?)Attribute.GetCustomAttribute(
        Assembly.GetExecutingAssembly(), typeof(AssemblyCompanyAttribute), false);
    var authors = companyAttribute != null ? companyAttribute.Company : "Unknown";
    // Console.WriteLine("文件夹压缩工具\n");
    // Console.WriteLine($"版本号: {version}\n");
    // Console.WriteLine($"作者: {authors}\n");
    // Console.WriteLine("...");

    string currentDir = Directory.GetCurrentDirectory();

    // 列出文件夹与zip文件
    List<string> items = [.. Directory.GetDirectories(currentDir), .. Directory.GetFiles(currentDir, "*.zip")];

    if (items.Count == 0)
    {
      Console.WriteLine("当前目录无可处理的文件夹或压缩文件。按回车退出。");
      Console.ReadLine();
      return;
    }

    int selectedIndex = 0;
    while (true)
    {
      Console.Clear();
      Console.WriteLine("文件夹压缩工具\n");
      Console.WriteLine($"版本号: {version}\n");
      Console.WriteLine($"作者: {authors}\n");
      Console.WriteLine("...");
      Console.WriteLine("请选择要操作的文件夹或zip文件 (↑↓选择，回车确定)：\n");
      for (int i = 0; i < items.Count; i++)
      {
        Console.WriteLine("{0} {1}",
            i == selectedIndex ? ">" : " ",
            Path.GetFileName(items[i]));
      }

      var key = Console.ReadKey(true).Key;
      switch (key)
      {
        case ConsoleKey.UpArrow:
          selectedIndex = (selectedIndex == 0) ? items.Count - 1 : selectedIndex - 1;
          break;
        case ConsoleKey.DownArrow:
          selectedIndex = (selectedIndex + 1) % items.Count;
          break;
        case ConsoleKey.Enter:
          HandleSelection(items[selectedIndex]);
          return;
        case ConsoleKey.Escape:
          Console.WriteLine("程序已退出，按回车键关闭窗口。");
          Console.ReadLine();
          return;
      }
    }
  }

  static void HandleSelection(string path)
  {
    if (Directory.Exists(path))
    {
      while (true)
      {
        Console.Clear();
        Console.WriteLine($"已选中文件夹：{Path.GetFileName(path)}\n");
        Console.WriteLine("请选择操作 (↑↓选择，回车确认)：\n");
        string[] menu = ["压缩该文件夹", "加密并压缩该文件夹"];
        int idx = 0;

        for (; ; )
        {
          for (int i = 0; i < menu.Length; i++)
          {
            Console.WriteLine("{0} {1}",
                i == idx ? ">" : " ",
                menu[i]);
          }
          var k = Console.ReadKey(true).Key;
          if (k == ConsoleKey.UpArrow) idx = (idx == 0) ? menu.Length - 1 : idx - 1;
          else if (k == ConsoleKey.DownArrow) idx = (idx + 1) % menu.Length;
          else if (k == ConsoleKey.Enter) break;
          else if (k == ConsoleKey.Escape)
          {
            Console.WriteLine("程序已退出。");
            return;
          }
          Console.SetCursorPosition(0, Console.CursorTop - menu.Length);
        }

        if (idx == 0) // 普通压缩
        {
          Console.WriteLine("开始进行普通压缩...\n");
          string zipName = Path.GetFileName(path) + ".zip";
          ZipFileFromFolder(path, zipName);
          Console.WriteLine($"操作完成，压缩包文件名 {zipName} 按回车退出。");

        }
        else // 加密压缩
        {
          string pwd1, pwd2;
          while (true)
          {
            Console.Write("请输入密码（仅限大小写字母和数字）：");
            pwd1 = ReadPassword();
            Console.Write("\n请再次输入密码：");
            pwd2 = ReadPassword();
            Console.WriteLine();
            if (pwd1.Length == 0 || pwd1 != pwd2)
            {
              Console.WriteLine("密码不一致或无效，请重新输入。");
            }
            else
            {
              break;
            }
          }
          Console.WriteLine("开始进行加密压缩...\n");
          string zipName = Path.GetFileName(path) + ".zip";
          EncryptZipFileFromFolder(path, zipName, pwd1);
          Console.WriteLine($"\n操作完成，压缩包文件名 {zipName} 按回车退出。");
        }
        Console.ReadLine();
        break;
      }
    }
    else if (File.Exists(path))
    {
      Console.WriteLine($"\n当前选中 {Path.GetFileName(path)} 是压缩文件，是否需要解压？(Y/n)");
      var key = Console.ReadKey(true).Key;
      if (key == ConsoleKey.Y || key == ConsoleKey.Enter)
      {
        // 解压
        try
        {
          string folderName = Path.GetFileNameWithoutExtension(path);
          Console.Write("\n如果存在密码，请输入密码（仅限数字、字母）（留空则无密码输入）：");
          string pwd = ReadPassword();
          Console.WriteLine("\n开始解压...\n");
          UnzipFile(path, folderName, pwd);
          Console.WriteLine($"\n操作完成，已解压至 {folderName}-unzip 按回车退出。");

        }
        catch (Exception ex)
        {
          Console.WriteLine($"\n解压失败，可能密码错误或文件无效。错误信息: {ex.Message}");
          Console.WriteLine($"\npath: {path}");
          string folderName = Path.GetFileNameWithoutExtension(path);
          Console.WriteLine($"\nfolderName: {folderName}\n");
          // 打印详细错误信息
          Console.WriteLine(ex.ToString());
          if (Directory.Exists(folderName + "-unzip"))
          {
            Directory.Delete(folderName + "-unzip", true);
          }
          Console.WriteLine($"\n操作失败，按回车退出。");
        }
      }
      Console.ReadLine();
    }
  }

  static string ReadPassword()
  {
    var pass = string.Empty;
    while (true)
    {
      var key = Console.ReadKey(true);
      if (key.Key == ConsoleKey.Enter) break;
      if (char.IsLetterOrDigit(key.KeyChar))
      {
        pass += key.KeyChar;
        Console.Write("*");
      }
    }
    return pass;
  }

  static void ZipFileFromFolder(string folderPath, string zipPath)
  {
    if (File.Exists(zipPath)) File.Delete(zipPath);
    using (var zip = new Ionic.Zip.ZipFile())
    {
      zip.AddDirectory(folderPath);

      int totalEntries = zip.Entries.Count;
      int processedEntries = 0;

      zip.SaveProgress += (sender, e) =>
      {
        if (e.EventType == Ionic.Zip.ZipProgressEventType.Saving_AfterWriteEntry)
        {
          processedEntries++;
          double percent = processedEntries * 100.0 / totalEntries;
          Console.Write($"\r压缩进度：{percent:F2}%");
        }
      };

      zip.Save(zipPath);
    }
  }

  static void EncryptZipFileFromFolder(string folderPath, string zipPath, string password)
  {
    if (File.Exists(zipPath)) File.Delete(zipPath);
    using (var zip = new Ionic.Zip.ZipFile())
    {
      zip.Password = password;
      zip.Encryption = Ionic.Zip.EncryptionAlgorithm.WinZipAes256;
      zip.AddDirectory(folderPath);

      int totalEntries = zip.Entries.Count;
      int processedEntries = 0;

      zip.SaveProgress += (sender, e) =>
      {
        if (e.EventType == Ionic.Zip.ZipProgressEventType.Saving_AfterWriteEntry)
        {
          processedEntries++;
          double percent = processedEntries * 100.0 / totalEntries;
          Console.Write($"\r压缩进度：{percent:F2}%");
        }
      };

      zip.Save(zipPath);
    }
    Console.WriteLine("\n压缩完成。");
  }

  static void UnzipFile(string zipPath, string extractFolder, string password)
  {
    extractFolder += "-unzip"; // 添加后缀
    if (!Directory.Exists(extractFolder))
      Directory.CreateDirectory(extractFolder);

    using (var zip = Ionic.Zip.ZipFile.Read(zipPath))
    {
      if (!string.IsNullOrEmpty(password)) zip.Password = password;

      // 创建条目副本以避免修改集合时的枚举问题
      var entries = new List<Ionic.Zip.ZipEntry>(zip.Entries);

      foreach (var entry in entries)
      {
        // 替换条目名称中的非法字符
        entry.FileName = SanitizeFileName(entry.FileName);
      }

      int totalEntries = zip.Entries.Count;
      int processedEntries = 0;

      zip.ExtractProgress += (sender, e) =>
      {
        if (e.EventType == Ionic.Zip.ZipProgressEventType.Extracting_AfterExtractEntry)
        {
          processedEntries++;
          double percent = processedEntries * 100.0 / totalEntries;
          Console.Write($"\r解压进度：{percent:F2}%");
        }
      };

      zip.ExtractAll(extractFolder, ExtractExistingFileAction.OverwriteSilently);
    }
    Console.WriteLine("\n解压完成。");
  }

  static string SanitizeFileName(string fileName)
  {
    foreach (char c in Path.GetInvalidFileNameChars())
    {
      if (c == '/' || c == '\\') continue;
      fileName = fileName.Replace(c, '_');
    }
    return fileName;
  }
}