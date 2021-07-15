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

			// Instantiate a new string builder, we will use this to help write our output .csv file
			var stringBuilder = new StringBuilder();

			// Append the first line which will contain the column headers
			stringBuilder.AppendLine(
				$"{nameof(MemberData.Guid)}," +
				$"{nameof(MemberData.Name)}," +
				$"{nameof(MemberData.Type)}," +
				$"{nameof(MemberData.Section)}," +
				$"{nameof(MemberData.Material)}," +
				$"{nameof(MemberData.TotalLength)}," +
				$"{nameof(MemberData.Span)}," +
				$"{nameof(MemberData.SpanLength)}," +
				$"{nameof(MemberData.Position)}," +
				$"{nameof(MemberData.LoadCase)}," +
				$"{nameof(MemberData.Axial)}," +
				$"{nameof(MemberData.ShearMajor)}," +
				$"{nameof(MemberData.ShearMinor)}," +
				$"{nameof(MemberData.Torsion)}," +
				$"{nameof(MemberData.MomentMajor)}," +
				$"{nameof(MemberData.MomentMinor)},");
			
			// Append second line which will contain units where applicable
			stringBuilder.AppendLine(
				$"," +
				$"," +
				$"," +
				$"," +
				$"," +
				$"[ft]," +
				$"," +
				$"[ft]," +
				$"," +
				$"," +
				$"[kip]," +
				$"[kip]," +
				$"[kip]," +
				$"[kip-ft]," +
				$"[kip-ft]," +
				$"[kip-ft],");

			// Unit conversion definitions
			// TODO: Factor this out into helper functions for clarity

			double newtonToKip = 0.00022480894387096;
			double millimeterToFoot = 0.00328084;

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
					Guid = member.Id.ToString(),
					Name = member.Name,
					Type = member.Type.ToString(),
					TotalLength = totalLength * millimeterToFoot,
				};

				// Loop through the span of each member
				for (int span = 0; span < member.SpanCount; span++)
				{

					// Get the span object based on index
					var spanObject = await member.GetSpanAsync(span);

					// Collect information about this span of the member
					memberData.Section = spanObject.ExtendedSection.Name;
					memberData.Material = spanObject.Material.Name;
					memberData.Span = span;
					memberData.SpanLength = spanObject.Length * millimeterToFoot;

					// For each span in the member, loop through the collection of solved loadcases
					foreach (var solvedLoadcase in solvedLoadcases)
					{

						// Skip over empty loadcases
						// TODO: Figure out why these are in here anyways
						if (solvedLoadcase.Name == "0 ")
                        {
							continue;
						}

						memberData.LoadCase = solvedLoadcase.Name;

						// Get the member loading for the solved loadcase
						var memberLoading = await member.GetLoadingAsync(solvedLoadcase.Id, analysisType, LoadingResultType.Base);

						if (memberLoading == null)
						{
							Console.WriteLine("The member loading could not be found!");

							continue;
						}

						// Set step size for position along span
						double step = 0.1;

						for (double positionRatio = 0.0; positionRatio <= 1.0; positionRatio += step)
						{

							memberData.Position = positionRatio;

							// Position is the measured length along the member, not the ratio
							// Covert ratio to length based on length of span
							var position = positionRatio * spanObject.Length;

							// Handle case where we are looking at the end of a span - for some reason output is zero when position = length of span
							if (positionRatio > 0.9)
							{
								position--;
							}

							// Get max axial force
							memberData.Axial = (await memberLoading.GetValueAsync(forceValueOption, axialLoadingDirection, span, position)).Max(lv => lv.Value) * newtonToKip;

							// Get max major axis shear force
							memberData.ShearMajor = (await memberLoading.GetValueAsync(forceValueOption, majorLoadingDirection, span, position)).Max(lv => lv.Value) * newtonToKip;

							// Get max minor axis shear force
							memberData.ShearMinor = (await memberLoading.GetValueAsync(forceValueOption, minorLoadingDirection, span, position)).Max(lv => lv.Value) * newtonToKip;

							// Get max torsion force
							memberData.Torsion = (await memberLoading.GetValueAsync(momentValueOption, axialLoadingDirection, span, position)).Max(lv => lv.Value) * newtonToKip * millimeterToFoot;

							// Get max major axis bending force
							memberData.MomentMajor = (await memberLoading.GetValueAsync(momentValueOption, majorLoadingDirection, span, position)).Max(lv => lv.Value) * newtonToKip * millimeterToFoot;

							// Get max minor axis bending force
							memberData.MomentMinor = (await memberLoading.GetValueAsync(momentValueOption, minorLoadingDirection, span, position)).Max(lv => lv.Value) * newtonToKip * millimeterToFoot;

							// Add a new line containing the data for this member to the output file
							stringBuilder.AppendLine(
								$"{memberData.Guid}," +
								$"{memberData.Name}," +
								$"{memberData.Type}," +
								$"{memberData.Section}," +
								$"{memberData.Material}," +
								$"{memberData.TotalLength}," +
								$"{memberData.Span}," +
								$"{memberData.SpanLength}," +
								$"{memberData.Position}," +
								$"{memberData.LoadCase}," +
								$"{memberData.Axial}," +
								$"{memberData.ShearMajor}," +
								$"{memberData.ShearMinor}," +
								$"{memberData.Torsion}," +
								$"{memberData.MomentMajor}," +
								$"{memberData.MomentMinor},");							

						}

					}

				}

			}

			// Write the .csv file with the output data to the same directory as the model
			File.WriteAllText($"{Path.GetDirectoryName(document.Path)}/MemberForces.csv", stringBuilder.ToString());

		}
		
		#endregion
	}
}

