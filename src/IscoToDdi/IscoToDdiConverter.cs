using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data;
using Algenta.Colectica.Model.Ddi;
using Algenta.Colectica.Model;
using Algenta.Colectica.Model.Ddi.Serialization;
using System.Xml;

namespace IscoToDdi
{
	/// <summary>
	/// Converts a spreadsheet describing the ISCO classification to DDI 3.1 XML.
	/// </summary>
	/// <remarks>
	/// The columns from the spreadsheet are mapped to the DDI elements as described below:
	/// 
	///   Code:				Code.Value
	///   Level:			Disregarded (can be determined from DDI model by inspecting parents)
	///   Description EN:	Category.Label["en"]
	///   Description FR:	Category.Label["fr"]
	///   Description DE:	Category.Label["de"]
	///   texte auto EN:	Category.Description["en"]
	///   texte auto DE:	Category.Description["de"]
	///   texte auto FR:	Category.Description["fr"]
	/// </remarks>
	public class IscoToDdiConverter
	{
		static string sheetName = "isco";

		DdiInstance instance;
		ResourcePackage resourcePackage;
		CategoryScheme categoryScheme;
		CodeScheme codeScheme;

		public void Convert(string inputFileName, string outputFileName)
		{
			VersionableBase.DefaultAgencyId = "example.org";

			if (string.IsNullOrWhiteSpace(inputFileName))
			{
				throw new ArgumentNullException("inputFileName");
			}
			if (string.IsNullOrWhiteSpace(outputFileName))
			{
				throw new ArgumentNullException("outputFileName");
			}

			InitializeDdiElements();

			// Go through each row of the spreadsheet.
			DataTable iscoTable = GetSpreadsheetContents(inputFileName);
			foreach (DataRow row in iscoTable.Rows)
			{
				// Create a category and map information from each column.
				Category category = new Category();
				category.Label["en"] = row["description EN"].ToString();
				category.Label["fr"] = row["description FR"].ToString();
				category.Label["de"] = row["description DE"].ToString();
				category.Description["en"] = row["texte auto EN"].ToString();
				category.Description["fr"] = row["texte auto FR"].ToString();
				category.Description["de"] = row["texte auto DE"].ToString();
				categoryScheme.Categories.Add(category);

				// Create a code.
				Code code = new Code();
				code.Category = category;
				code.Value = row["code"].ToString();

				// First level codes are added directly the CodeScheme.
				int level = code.Value.Length;
				if (level == 1)
				{
					codeScheme.Codes.Add(code);
				}
				else
				{
					// For child codes, look up the parent code. 
					// The parent's value will be the same as this new code's value,
					// minus the last character.
					var flatCodes = codeScheme.GetFlattenedCodes();
					Code parent = flatCodes.SingleOrDefault(c => c.Value == code.Value.Substring(0, code.Value.Length - 1));
					parent.ChildCodes.Add(code);
				}
				

				Console.WriteLine(string.Format("{0} {1}", code.Value, category.Label["en"]));
			}

			// Save the DDI XML.
			DDIWorkflowSerializer serializer = new DDIWorkflowSerializer();
			serializer.UseConciseBoundedDescription = false;
			XmlDocument doc = serializer.Serialize(this.instance);
			DDIWorkflowSerializer.SaveXml(outputFileName, doc);
		}

		void InitializeDdiElements()
		{
			resourcePackage = new ResourcePackage();
			resourcePackage.DublinCoreMetadata.Title["en"] = "ISCO Classification";
			resourcePackage.Abstract.Strings["en"] = "Not specified";
			resourcePackage.Purpose.Strings["en"] = "Not specified";

			categoryScheme = new CategoryScheme();
			categoryScheme.Label["en"] = "ISCO Categories";

			codeScheme = new CodeScheme();
			codeScheme.Label["en"] = "ISCO Codes";

			resourcePackage.CategorySchemes.Add(categoryScheme);
			resourcePackage.CodeSchemes.Add(codeScheme);

			instance = new DdiInstance();
			instance.ResourcePackages.Add(resourcePackage);
		}

		DataTable GetSpreadsheetContents(string inputFileName)
		{
			// Load the spreadsheet.
			string connectionString = string.Format(
				"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";",
				inputFileName);

			var adapter = new OleDbDataAdapter(string.Format("SELECT * FROM [{0}$]", sheetName), connectionString);
			var ds = new DataSet();

			adapter.Fill(ds, sheetName);

			return ds.Tables[0];
		}
	}
}
