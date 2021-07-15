using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using TSD.API.Remoting;
using TSD.API.Remoting.Loading;
using TSD.API.Remoting.Materials;
using TSD.API.Remoting.Solver;
using TSD.API.Remoting.Structure;

namespace ExtractMemberForces
{
    internal static class ExtractMemberForces
    {

		#region Methods

		/// <summary>
		/// Main entry point to the application, executes on startup
		/// </summary>
		public static async Task Main()
		{
			// Get all instances of TSD running on the local machine
			var tsdRunningInstances = await ApplicationFactory.GetRunningApplicationsAsync();

			if (!tsdRunningInstances.Any())
			{
				Console.WriteLine("No running instances of TSD found!");

				return;
			}

			// Get the first running TSD instance found
			var tsdInstance = tsdRunningInstances.First();

			// Get the active document from the running instance of TSD
			var document = await tsdInstance.GetDocumentAsync();

			if (document == null)
			{
				Console.WriteLine("No document was found in the TSD instance!");

				return;
			}

			// Get the model from the document
			var model = await document.GetModelAsync();

			if (model == null)
			{
				Console.WriteLine("No model was found in the document!");

				return;
			}

			// Request first order linear analysis type
			var analysisType = AnalysisType.FirstOrderLinear;

			// Get the solver models for the requested analysis types
			var solverModels = await model.GetSolverModelsAsync(new[] { analysisType });

			if (!solverModels.Any())
			{
				Console.WriteLine("No solver models found!");

				return;
			}

			// Get the first order linear solver model
			var firstOrderLinearSolverModel = solverModels.FirstOrDefault();

			if (firstOrderLinearSolverModel == null)
			{
				Console.WriteLine("No solver models found for the first order linear analysis type!");

				return;
			}

			// Get the results for the first order linear solver model
			var firstOrderLinearSolverResults = await firstOrderLinearSolverModel.GetResultsAsync();

			if (firstOrderLinearSolverResults == null)
			{
				Console.WriteLine("No results found for the first order linear analysis type!");

				return;
			}

			// Get the first order linear analysis results
			var analysis3DResults = await firstOrderLinearSolverResults.GetAnalysis3DAsync();

			if (analysis3DResults == null)
			{
				Console.WriteLine("No analysis results found!");

				return;
			}

			// Get the Guids of the solved loading
			var solvedLoadingGuids = await analysis3DResults.GetSolvedLoadingAsync();

			if (!solvedLoadingGuids.Any())
			{
				Console.WriteLine("No solved loading guids found!");

				return;
			}

			// Get all loadcases in the model. Pass null as the parameter to get all loadcases; alternatively a sequence of indices can be passed to get specific loadcases
			var loadcases = await model.GetLoadcaseAsync(null);

			// Get the solved loadcases by cross referencing with the sequence of solved guids obtained from the solver model
			var solvedLoadcases = loadcases.Where(l => solvedLoadingGuids.Contains(l.Id));

			// Remove blank loadcases from the list

			if (!solvedLoadcases.Any())
			{
				Console.WriteLine("No solved loadcases found!");

				return;
			}

			// Create a new Excel workbook for storing extracted data
			var workbook = new XLWorkbook();
			var worksheet = workbook.Worksheets.Add("Member Forces");

			// Add headers to the worksheet
			worksheet.Cell("A1").Value = "Name";
			worksheet.Cell("B1").Value = "Type";
			worksheet.Cell("C1").Value = "Section";
			worksheet.Cell("D1").Value = "Total Length";
			worksheet.Cell("E1").Value = "Span";
			worksheet.Cell("F1").Value = "Span Length";
			worksheet.Cell("G1").Value = "Load Case";
			worksheet.Cell("H1").Value = "Position";
			worksheet.Cell("I1").Value = "Fx";
			worksheet.Cell("J1").Value = "Fy";
			worksheet.Cell("K1").Value = "Fz";
			worksheet.Cell("L1").Value = "Mxx";
			worksheet.Cell("M1").Value = "Myy";
			worksheet.Cell("N1").Value = "Mzz";

			var row = 2;

			// Get all members in the model. Pass null as the parameter to get all members; alternatively a sequence of indices can be passed to get specific members
			var members = await model.GetMemberAsync(null);

			if (!members.Any())
			{
				Console.WriteLine("No members found in the model!");

				return;
			}

			// Specify the loading value options
			var forceValueOption = new LoadingValueOptions(LoadingValueType.Force);
			var momentValueOption = new LoadingValueOptions(LoadingValueType.Moment);

			// Specify the loading directions
			var axialLoadingDirection = LoadingDirection.Axial;
			var majorLoadingDirection = LoadingDirection.Major;
			var minorLoadingDirection = LoadingDirection.Minor;

			// Loop through all members
			foreach (var member in members)
			{
				// Define the total length as 0.0 initially
				double totalLength = 0.0;

				// Loop through the span of each member
				for (int span = 0; span < member.SpanCount; span++)
				{
					// Cumulatively add in the length of each individual span within the member
					totalLength += (await member.GetSpanAsync(span)).Length;
				}

				// Create the output data object and set some member properties
				var memberData = new MemberData
				{
					MemberName = member.Name,
					MemberType = member.Type.ToString(),
					MemberTotalLength = totalLength,
				};

				// Loop through the span of each member
				for (int span = 0; span < member.SpanCount; span++)
				{
					memberData.MemberSection = (await member.GetSpanAsync(span)).ExtendedSection.Name;
					memberData.Span = span;
					memberData.SpanLength = (await member.GetSpanAsync(span)).Length;

					// For each span in the member, loop through the collection of solved loadcases
					foreach (var solvedLoadcase in solvedLoadcases)
					{

						memberData.LoadCase = solvedLoadcase.Name;

						// Get the member loading for the solved loadcase
						var memberLoading = await member.GetLoadingAsync(solvedLoadcase.Id, analysisType, LoadingResultType.Base);

						if (memberLoading == null)
						{
							Console.WriteLine("The member loading could not be found!");

							continue;
						}

						// Set step size for position along span
						double step = 0.25;

						for (double positionRatio = 0.0; positionRatio <= 1.0; positionRatio += step)
						{

							memberData.Position = positionRatio;

							// Position is the measured length along the member, not the ratio
							// Covert ratio to length based on length of span
							var position = positionRatio * memberData.SpanLength;

							// Handle case where we are looking at the end of a span - for some reason output is zero when position = length of span
							if (positionRatio > 0.9)
							{
								position--;
							}

							// Get max axial force
							memberData.Fx = (await memberLoading.GetValueAsync(forceValueOption, axialLoadingDirection, span, position)).Max(lv => lv.Value);

							// Get max minor axis shear force
							memberData.Fy = (await memberLoading.GetValueAsync(forceValueOption, minorLoadingDirection, span, position)).Max(lv => lv.Value);

							// Get max major axis shear force
							memberData.Fz = (await memberLoading.GetValueAsync(forceValueOption, majorLoadingDirection, span, position)).Max(lv => lv.Value);

							// Get max torsion force
							memberData.Mxx = (await memberLoading.GetValueAsync(momentValueOption, axialLoadingDirection, span, position)).Max(lv => lv.Value);

							// Get max major axis bending force
							memberData.Myy = (await memberLoading.GetValueAsync(momentValueOption, majorLoadingDirection, span, position)).Max(lv => lv.Value);

							// Get max minor axis bending force
							memberData.Mzz = (await memberLoading.GetValueAsync(momentValueOption, minorLoadingDirection, span, position)).Max(lv => lv.Value);

							// Write data to excel sheet
							worksheet.Cell(row, 1).Value = memberData.MemberName;
							worksheet.Cell(row, 2).Value = memberData.MemberType;
							worksheet.Cell(row, 3).Value = memberData.MemberSection;
							worksheet.Cell(row, 4).Value = memberData.MemberTotalLength;
							worksheet.Cell(row, 5).Value = memberData.Span;
							worksheet.Cell(row, 6).Value = memberData.SpanLength;
							worksheet.Cell(row, 7).Value = memberData.LoadCase;
							worksheet.Cell(row, 8).Value = memberData.Position;
							worksheet.Cell(row, 9).Value = memberData.Fx;
							worksheet.Cell(row, 10).Value = memberData.Fy;
							worksheet.Cell(row, 11).Value = memberData.Fz;
							worksheet.Cell(row, 12).Value = memberData.Mxx;
							worksheet.Cell(row, 13).Value = memberData.Myy;
							worksheet.Cell(row, 14).Value = memberData.Mzz;

							// Increment worksheet row
							row++;
						}

					}

				}

			}

			// Save the workbook to the same directory as the model
			workbook.SaveAs($"{Path.GetDirectoryName(document.Path)}/MemberForceExtraction.xlsx");
		}
		#endregion
	}
}

