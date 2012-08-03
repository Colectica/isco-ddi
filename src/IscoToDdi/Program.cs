using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace IscoToDdi
{
	class Program
	{
		static int Main(string[] args)
		{
			if (args.Length != 2)
			{
				ShowUsage();
				return 1;
			}

			string inputFileName = args[0];
			string outputFileName = args[1];

			if (!File.Exists(inputFileName))
			{
				Console.WriteLine("Sorry, the input file does not exist.");
				ShowUsage();
			}

			IscoToDdiConverter converter = new IscoToDdiConverter();
			converter.Convert(inputFileName, outputFileName);

#if DEBUG
			Console.WriteLine("Press enter to end the program.");
			Console.ReadLine();
#endif

			return 0;
		}

		static void ShowUsage()
		{
			Console.WriteLine("The ISCO to DDI tool describes the ISCO classification using the DDI 3.1 XML standard.");
			Console.WriteLine("The tool reads the ISCO classification from a spreadsheet with the following columns:");
			Console.WriteLine("  id,code,level,description EN,description FR,description DE,texte auto EN,texte auto DE,texte auto FR");
			Console.WriteLine();
			Console.WriteLine("Usage:");
			Console.WriteLine();
			Console.WriteLine("  IscoToDdi.exe inputFile outputFile");
		}
	}
}
