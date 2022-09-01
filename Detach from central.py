# Load the Python Standard and DesignScript Libraries
import sys
import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *
# Import ToDSType(bool) extension method
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
# Import geometry conversion extension methods
clr.ImportExtensions(Revit.GeometryConversion)
# Import DocumentManager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *
# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import SaveAsOptions

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application


import os
import datetime

open_options = OpenOptions() #Creating a new OpenOptions object
open_options.DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets
#worksetConfig = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
#open_options.SetOpenWorksetsConfiguration(worksetConfig)
#open_options.Audit = True
report = [] #Creating an empty list, which will contain our final report

dateStamp   = datetime.datetime.today().strftime("%d%m%y")
timeStamp   = datetime.datetime.today().strftime("%H:%M:%S")

#path = Document.PathName.GetValue(doc)
folder = ()
if IN[0]:
	folder = IN[0]
	
directory = str(IN[1])

run_nwc = IN[2]

#path,path.rsplit("\\",1)
failed = []
no_models = 0
errorReport = "SCRIPT DID NOT RUN CORRECTLY - CHECK DIRECTORY SETUP"
path = ()	

if IN[0]:
	for f in folder:
		path = str(f)
		filepath = FilePath(path) 
		
		try:
			
			doc = app.OpenDocumentFile(filepath, open_options) 			
			
			t = Transaction(doc, 'Model Clean')
			t.Start()
			
			# --- CLEAN SHEETS --- 
			
			#tdelSheets = Transaction(doc, 'Delete Sheets')
			#tdelSheets.Start()
			
			all_sheet_types = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
			all_sheet_types.WhereElementIsNotElementType()
			all_sheets = [i for i in all_sheet_types]
			number_sheets = len(all_sheets)
			export_sheet = []
			export_sheet_numbers = [i.SheetNumber for i in export_sheet]
			export_sheet_exists = ("NO")	
			del_sheets = []
			sheet_numbers = []
			for sheet in all_sheets:
				try:			
					sheet_no = sheet.SheetNumber
					sheet_numbers.append(sheet_no)
					if sheet_no == "01-E": 
						export_sheet.append(UnwrapElement(sheet))
						export_sheet_exists = "YES"
					else:
						del_sheets.append(UnwrapElement(sheet))
				except:
					pass
			for s in del_sheets:	
				try:
					doc.Delete(s.Id)
				except:
					pass				
			#doc.Regenerate()
			#tdelSheets.Commit()
			export_sheets = ' and '.join(map(str,export_sheet_numbers))
			report_sheets = "All sheets (" + str(number_sheets) + ") were deleted."
			#if export_sheet_exists == ("YES"):
				#report_sheets = "All sheets (" + str(number_sheets) + ") were deleted except for: " + str(export_sheets)
			#if export_sheet_exists == ("NO"):
				#report_sheets = "All sheets (" + str(number_sheets) + ") were deleted."
				
			# --- CLEAN VIEWS --- 
			
			all_view_types = FilteredElementCollector(doc).OfClass(clr.GetClrType(View)).WhereElementIsNotElementType().ToElements()
			all_views=[]
				
			for v in all_view_types:
				#if v.IsTemplate == False:
				all_views.append(UnwrapElement(v))
			v_names = [] 
			del_views = []
			export_view = []
			export_view_exists = ("NO")		
			title_view = ()
			for v in all_views:
				v_name = v.Name
				v_names.append(v_name)
				#if v_name == "3D Export (Revit)" or "Project View" or "System Browser":
				if v_name == "Project View":
					export_view.append(v_name)
				elif v_name == "3D EXPORT (Revit)":
					export_view.append(v_name)
					export_view_exists = "YES"
					title_view = v.Id
				elif v_name == "System Browser":
					export_view.append(v_name)
				#elif v_name == "3D View: 3D EXPORT (Revit)":
					#export_view.append(v_name)	
				else:
					del_views.append(v)			
			number_views = len(all_views)
			number_expviews = len(del_views)
		
			#tdelViews = Transaction(doc, 'Delete Views')
			#tdelViews.Start()
			#TransactionManager.Instance.EnsureInTransaction(doc)
			
			for v in del_views:
				try:
					vid = v.Id
					doc.Delete(vid)
				except:
					pass
			#doc.Regenerate()
			#tdelViews.Commit()		
			#TransactionManager.Instance.TransactionTaskDone()		
			
			report="File {} contains {} views".format(path, number_views)
			export_views = ' and '.join(map(str,export_view))
			report_2 = "All views (" + str(number_views - number_expviews) + ") were deleted except for: " + export_views 
			# --- CLEAN SCHEDULES --- 
			
			all_schedules = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Schedules).WhereElementIsNotElementType().ToElementIds()
			
			for s in all_schedules:
				doc.Delete(s)
			
			length_sched = len([i for i in all_schedules])
			report_sched = "All schedules (" + str(length_sched) + ") were deleted." 
			
			# --- FILTER CAD LINKS --- 
			
			cad_links = FilteredElementCollector(doc).OfClass(clr.GetClrType(ImportInstance)).ToElementIds()
			for c in cad_links:
				doc.Delete(c)
			
			length_cad = len([i for i in cad_links])
			report_cad = "All CAD links (" + str(length_cad) + ") were deleted." 
				
			# --- FILTER REVIT LINKS --- 
			
			rvt_links = FilteredElementCollector(doc).OfClass(clr.GetClrType(RevitLinkInstance)).WhereElementIsNotElementType().ToElements()
			rvt_link_ids = FilteredElementCollector(doc).OfClass(clr.GetClrType(RevitLinkInstance)).ToElementIds()
			link_types = []
			for r in rvt_links:
				link_types.append(doc.GetElement(r.GetTypeId()))
				
			length_rvt = len([i for i in rvt_links])
			
			for r in rvt_link_ids:
				doc.Delete(r)
			for l in link_types:
				doc.Delete(l.Id)
			
			
			report_rvt = "All RVT links (" + str(length_rvt) + ") were deleted." 
			
			# --- FILTER ROOMS ---
			
			all_rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
			placed , unplaced = [],[]
			for room in all_rooms:
				if room.Area > 0:
					placed.append(room)
				else:
					unplaced.append(room.Id)
					
			delrooms = [doc.Delete(i) for i in unplaced]

			length_rms = len([i for i in all_rooms])
			length_unplaced = len(unplaced)
			report_rms = "Of " + str(length_rms) + " rooms, " + str(length_unplaced) + " were not placed, and so deleted."
			
			
			t.Commit()
			
			# --- ENABLE WORKSHARING --- 
			#tEnabWorkSharing = Transaction(doc, 'Enable Worksharing')
			#tEnabWorkSharing.Start()
			#doc.EnableWorksharing("Shared Levels and Grids", "Workset1")
						
			#tEnabWorkSharing.Commit()
			
			# --- SAVING & REPORTING --- 
							
			text_to_be_replaced = ".rvt"
			replacement_text = "_EXPORT_" + dateStamp + ".rvt"
			NewFilePath = path.replace(text_to_be_replaced, replacement_text)
			nav_replacement = ""
			navisworks_path = path.Replace(text_to_be_replaced, nav_replacement)
			navis_split = navisworks_path.Split("\\")
			navis_path = navis_split[-1]
			#navisworks_path = "navexp"
			
			dir_replaced = "/"
			dir_replacement = "//"
			directory_raw = directory.replace(dir_replaced, dir_replacement)
			
			
			save_options = SaveAsOptions()
			save_options.OverwriteExistingFile = True
			#work_save_options = WorksharingSaveAsOptions()
			#work_save_options.SaveAsCentral = True
			#work_save_options.OpenWorksetsDefault = SimpleWorksetConfiguration.AllWorksets
			#save_options.SetWorksharingOptions(work_save_options)
			
			doc.SaveAs(str(NewFilePath),save_options)
			
			#tSave.Commit()
			
			#transclose = TransactionManager.Instance
			#transclose.ForceCloseTransaction()
			
			# --- NAVISWORKS EXPORT --- 
			
			nweOptions = NavisworksExportOptions()
			nweOptions.ExportScope = NavisworksExportScope.Model
			nweOptions.ViewId = title_view
			nweOptions.ConvertLinkedCADFormats = True
			nweOptions.FacetingFactor = 16
			nweOptions.ExportLinks = False
			
			nwc_run = "NO"

			if run_nwc == True:

				doc.Export(directory_raw, navis_path , nweOptions)
				nwc_run = "YES"
			
			# --- CLOSE --- 
			
			doc.Close(False) #(False = without saving)
			
			# --- FEEDBACK--- 
			
			no_models += 1
			
			myLog   =  "M:\9. Dynamo\\2. Dynamo Scripts\\20. Log files\\TestReport.csv"
			dataRow = report
			# Adds new line to log file or creates one if doesn't exist
			try:
				with open(myLog, "a") as file:
					file.writelines(dataRow + "\n")
				result = True
			except:
				result = False
			#Result = "File: " + NewFilePath + " succesfully exported without errors. " + '\n' + report_2 + '\n' + "Export View: \"3D Export (Revit)\" existed?: " + export_view_exists + '\n' + report_sheets + "Export Sheet with number\"01-E\" existed?: " + export_sheet_exists
			Result = "File: " + NewFilePath + " succesfully exported without errors. " + '\n' + report_2 + '\n' + "Export View: \"3D EXPORT (Revit)\" existed?: " + export_view_exists + '\n' + report_sheets + '\n' + report_sched + '\n' + report_cad + '\n' + report_rvt + '\n' + report_rms + '\n' + "Navisworks export success?:" + nwc_run
			if errorReport != "SCRIPT DID NOT RUN CORRECTLY - CHECK DIRECTORY SETUP":
				errorReport = errorReport + '\n' + '\n' + Result
			else:
				errorReport = Result
		except:
			import traceback
			errors = traceback.format_exc()
			#t.Rollback()
			#transclose = TransactionManager.Instance
			#transclose.ForceCloseTransaction()
			#doc.Close(False)
			"""dialog = TaskDialog("*** MERIT AUTOMATION DIALOGUE ***")
			dialog.MainInstruction = "An error has occured!"
			dialog.MainContent = "test"
			dialog.CommonButtons = TaskDialogCommonButtons.Close;
			dialog.DefaultButton = TaskDialogResult.Close;
			dialog.Show"""
			Result = "File: " + "path" + '\n' + "ERROR EXPORTING. Please copy paste this error message and send to Justin: " + '\n' + errors
			
			if errorReport != "SCRIPT DID NOT RUN CORRECTLY - CHECK DIRECTORY SETUP":
				errorReport = errorReport + Result
			else:
				errorReport = Result 
			
else:
	pass
	
errorReport = errorReport + '\n' + '\n' + str(no_models) + " model(s) processed."
#OUT = navisworks_path
OUT = errorReport, no_models, sheet_numbers


	