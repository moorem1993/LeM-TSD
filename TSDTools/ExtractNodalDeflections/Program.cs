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

namespace ExtractNodalDeflections
{
	internal static class ExtractNodalDeflections
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
				$"{nameof(NodeData.NodeIndex)}," +
				$"{nameof(NodeData.LoadCase)}," +
				$"{nameof(NodeData.Mx)}," +
				$"{nameof(NodeData.My)}," +
				$"{nameof(NodeData.Mz)}," +
				$"{nameof(NodeData.Rx)}," +
				$"{nameof(NodeData.Ry)}," +
				$"{nameof(NodeData.Rz)},");

			// Append second line which will contain units where applicable
			stringBuilder.AppendLine(
				$"," +
				$"," +
				$"[in]," +
				$"[in]," +
				$"[in]," +
				$"[rad]," +
				$"[rad]," +
				$"[rad],");

			// Unit conversion definitions
			// TODO: Factor this out into helper functions for clarity

			double mmToInch = 0.0393701;

			// Set Loading Result Type
			var loadingResultType = LoadingResultType.Base;

			// Loop through all solved loadcases
			foreach (var solvedLoadcase in solvedLoadcases)
			{

				// Skip over empty loadcases
				if (solvedLoadcase.Name == "0 ")
				{
					continue;
				}

				// Get all nodal displacements in the model for the given loadcase
				var nodalDisplacements = await analysis3DResults.GetNodalDisplacementAsync(solvedLoadcase.Id, loadingResultType, null);

				// Loop through each displacement object and extract data

				foreach (var nodalDisplacement in nodalDisplacements)
                {

					// Add a new line containing the data for this node to the output file
					stringBuilder.AppendLine(
						$"{nodalDisplacement.NodeIndex}," +
						$"{solvedLoadcase.Name}," +
						$"{nodalDisplacement.Displacement.Mx * mmToInch}," +
						$"{nodalDisplacement.Displacement.My * mmToInch}," +
						$"{nodalDisplacement.Displacement.Mz * mmToInch}," +
						$"{nodalDisplacement.Displacement.Rx}," +
						$"{nodalDisplacement.Displacement.Ry}," +
						$"{nodalDisplacement.Displacement.Rz},");
				}

			}

			// Write the .csv file with the output data to the same directory as the model
			File.WriteAllText($"{Path.GetDirectoryName(document.Path)}/NodalDisplacements.csv", stringBuilder.ToString());

		}

		#endregion
	}
}