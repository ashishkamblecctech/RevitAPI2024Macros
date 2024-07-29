/*
 * Created by SharpDevelop.
 * User: Ashish Kamble
 * Date: 13-02-2024
 * Time: 12:09
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Analysis;
using System.Text;

namespace SystemAnalysis
{
	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
	[Autodesk.Revit.DB.Macros.AddInId("DD768D52-1D13-4A43-B817-F877E31FB221")]
	public partial class ThisApplication
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		public void openFile()
		{
			Document doc = this.ActiveUIDocument.Document;
			using (Transaction trans = new Transaction(doc))
			{
				
				string pathName = ".\\E:\\Revit_SystemAnalysis\\Ashish\\input.rvt";
				
				FilePath path = new FilePath(pathName);
				RevitLinkOptions options = new RevitLinkOptions(false);
				// Create new revit link storing absolute path to file
				LinkLoadResult result = RevitLinkType.Create(doc, path, options);
				//result.Dispose();
				trans.Commit();
			}	
		}
		
		public void LocationAPI()
		{
			Document document = this.ActiveUIDocument.Document;
			
			// Get the project location handle
			ProjectLocation projectLocation = document.ActiveProjectLocation;
			
			// Show the information of current project location
			XYZ origin = new XYZ(0, 0, 0);
			ProjectPosition position = projectLocation.GetProjectPosition(origin);
			if (null == position)
			{
				throw new Exception("No project position in origin point.");
			}
			
			// Format the prompt string to show the message.
			String prompt = "Current project location information:\n";
			prompt += "\n\t" + "Origin point position:";
			prompt += "\n\t\t" + "Angle: " + position.Angle;
			prompt += "\n\t\t" + "East to West offset: " + position.EastWest;
			prompt += "\n\t\t" + "Elevation: " + position.Elevation;
			prompt += "\n\t\t" + "North to South offset: " + position.NorthSouth;
			
			// Angles are in radians when coming from Revit API, so we
			// convert to degrees for display
			const double angleRatio = Math.PI / 180;   // angle conversion factor
			
			SiteLocation site = projectLocation.GetSiteLocation();
			prompt += "\n\t" + "Site location:";
			prompt += "\n\t\t" + "Latitude: " + site.Latitude / angleRatio + "��";
			prompt += "\n\t\t" + "Longitude: " + site.Longitude / angleRatio + "��";
			prompt += "\n\t\t" + "TimeZone: " + site.TimeZone;
			
			// Give the user some information
			TaskDialog.Show("Revit",prompt);

		}
		
		public void GetProjectLocation()
		{
			Document document = this.ActiveUIDocument.Document;
			ProjectLocation currentLocation = document.ActiveProjectLocation;

			//get the project position
			XYZ origin = new XYZ(0, 0, 0);
			
			const double angleRatio = Math.PI / 180;   // angle conversion factor
			
			ProjectPosition projectPosition = currentLocation.GetProjectPosition(origin);
			//Angle from True North
			double angle = 0 * angleRatio;   // convert degrees to radian
			double eastWest =73.8567;     //East to West offset
			double northSouth = 18.5204;   //North to South offset
			double elevation = 560.0;    //Elevation above ground level
			
			//create a new project position
			ProjectPosition newPosition =
				document.Application.Create.NewProjectPosition(eastWest, northSouth, elevation, angle);
			
			if (null != newPosition)
			{
				//set the value of the project position
				currentLocation.SetProjectPosition(origin, newPosition);
			}
		}
		
		public void CreateZone(Level level, Phase phase, Document doc)
		{
			Dictionary<ElementId, List<Zone>> m_zoneDictionary = new Dictionary<ElementId, List<Zone>>();
			Zone zone = doc.Create.NewZone(level, phase);
			if (zone != null)
			{
				m_zoneDictionary[level.Id].Add(zone);
			}
		}
		
		public void CreateEnergyAnalysis()
		{
			Document doc = this.ActiveUIDocument.Document;
			
			// Collect space and surface data from the building's analytical thermal model
			EnergyAnalysisDetailModelOptions options = new EnergyAnalysisDetailModelOptions();
			options.Tier = EnergyAnalysisDetailModelTier.Final; // include constructions, schedules, and non-graphical data in the computation of the energy analysis model
			options.EnergyModelType = EnergyModelType.SpatialElement;   // Energy model based on rooms or spaces
			
			using (Transaction trans = new Transaction(doc))
			{
				trans.Start("Add Zone");
				
				EnergyAnalysisDetailModel eadm = EnergyAnalysisDetailModel.Create(doc, options);
				IList<EnergyAnalysisSpace> spaces = eadm.GetAnalyticalSpaces();
				StringBuilder builder = new StringBuilder();
				builder.AppendLine("Spaces: " + spaces.Count);
				foreach (EnergyAnalysisSpace space in spaces)
				{
					SpatialElement spatialElement = doc.GetElement(space.CADObjectUniqueId) as SpatialElement;
					ElementId spatialElementId = spatialElement == null ? ElementId.InvalidElementId : spatialElement.Id;
					builder.AppendLine("   >>> " + space.SpaceName + " related to " + spatialElementId);
					IList<EnergyAnalysisSurface> surfaces = space.GetAnalyticalSurfaces();
					builder.AppendLine("       has " + surfaces.Count + " surfaces.");
					foreach (EnergyAnalysisSurface surface in surfaces)
					{
						builder.AppendLine("            +++ Surface from " + surface.OriginatingElementDescription);
					}
				}
				TaskDialog.Show("EAM", builder.ToString());
				trans.Commit();
			}
		}
	}
}