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

import math 
doc = DocumentManager.Instance.CurrentDBDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument


PAM_no_OG = IN[0]
PAM_nos_string = IN[1]
PAM_nos_New = PAM_nos_string.split(",")

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ PART 1 - VIEW SETUP ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

"""
# --- GET VIEWS ---

all_views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views)
all_views.WhereElementIsNotElementType()

views_OG = []

for v in all_views:
	v_name = v.Name
	if v_name.Contains(str(PAM_no_OG)):
		views_OG.append(UnwrapElement(v))
		
"""
		
# --- MISC DEFINITIONS ---

utid = UnitTypeId.Millimeters

def conv(x):
	c = UnitUtils.ConvertToInternalUnits(x,utid)
	return c
	
def change_name(x,n):
	param_viewname = x.get_Parameter(BuiltInParameter.VIEW_NAME)
	clone_name = x.Name
	new_name_1 = clone_name.replace(str(PAM_no_OG),str(n))
	new_name_2 = new_name_1.replace("Copy 1", " ")
	param_viewname.Set(new_name_2)
	
def sched_change_name(s,n):
	param_viewname = s.get_Parameter(BuiltInParameter.VIEW_NAME)
	clone_name = s.Name
	new_name_1 = clone_name.replace("Copy 1", " - " + n)
	param_viewname.Set(new_name_1)
	
def prefix(assembly_id):
	el = doc.GetElement(assembly_id)
	n = el.Name
	return n
	
# --- GET TITLEBLOCKS ---
"""
titleblocks = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_TitleBlocks)

t_names = []
t_search = []
#t_test =t_search[0]

for t in titleblocks:
	UnwrapElement(t)
	t_id = t.Id
	ele = doc.GetElement(t_id).FamilyName
	if ele.Contains("MRT_ANO_A0_TB_2019"):
		t_search.append(t_id)
	t_names.append(ele)
	
"""

# --- GET SHEET ---

sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets)
sheets.WhereElementIsNotElementType()


# sheet_names = []
sheet_no_list = []
sheet = ()

for s in sheets:
	sheet_no = s.SheetNumber
	if sheet_no.Contains(str(PAM_no_OG)):
		unwrapped = UnwrapElement(s)
		sheet = unwrapped
		
def sheet_name(s,n):
	param_name = s.get_Parameter(BuiltInParameter.SHEET_NAME)
	param_number = s.get_Parameter(BuiltInParameter.SHEET_NUMBER)
	number_OG = sheet.SheetNumber
	new_number = str(number_OG).replace(str(PAM_no_OG),str(n))
	param_number.Set(new_number)
	new_name = str(sheet.Name)
	param_name.Set(new_name)
	return new_number

# --- GET TITLEBLOCK ---
	
titleblock = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks)

tblock= []
for t in titleblock:
	UnwrapElement(t)
	sy = t.Symbol
	s_id = sy.Id
	tblock.append(s_id)
	
# --- GET REFERENCE COORDINATES ---

viewport_ids = sheet.GetAllViewports()
views_OG = []
viewport_names = []

for i in viewport_ids:	
	v = doc.GetElement(i)
	v_id = v.ViewId
	v_el = doc.GetElement(v_id)
	v_name = v_el.Name
	if v_el.ViewType != ViewType.Legend:
		views_OG.append(v)
		viewport_names.append(v_name)

center_points = []

for v in views_OG:
	c = v.GetBoxCenter()
	center_points.append(c)

# --- GET SCHEDULES ---

schedules_OG = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_ScheduleGraphics)
schedules_OG.ToElements()
schedules_OG_names = []
schedules_OG_unwrapped = []
master_Id = []

for s in schedules_OG:
	name = s.Name
	if name.Contains("Revision Schedule"):
		pass
	else:
		unwrapped = UnwrapElement(s)
		schedules_OG_unwrapped.append(unwrapped)
		schedules_OG_names.append(name)
		master_sched_Id = s.ScheduleId
		master_Id.append(master_sched_Id)
	
def get_ymax(sched,sh):
	bb = sched.get_BoundingBox(sh)
	bb_min = bb.Max
	y = bb_min.Y
	return y

def get_ymin(sched, sh):
	bb = sched.get_BoundingBox(sh)
	bb_min = bb.Min
	y = bb_min.Y
	return y

def get_xmin(sched, sh):
	bb = sched.get_BoundingBox(sh)
	bb_min = bb.Min
	x = bb_min.X
	return x
	
sched_0 = schedules_OG_unwrapped[0]
no1_sched = master_Id[0]


x = get_xmin(sched_0, sheet)
y1 = get_ymin(sched_0, sheet)
y2 = get_ymax(sched_0, sheet)

point_0 = XYZ(x,y2,0)

flexi_schedules = master_Id[1:]
#flexi_schedules = schedules_OG_unwrapped[2:]

# --- DUPLICATE SHEETS AND VIEWS ---

TransactionManager.Instance.EnsureInTransaction(doc)

new_sheets = []

result = []
#y_flexi = get_ymin(result)
#flexi_point = XYZ(x,y_flexi,0)

for n in PAM_nos_New:	
	scheds_on_sheet = []
	schedule_no = 0
	s = Autodesk.Revit.DB.ViewSheet.Create(doc, tblock[0])
	new_sheet = sheet_name(s,n)
	new_sheets.append(s)
	first_sched_clone = doc.GetElement(no1_sched).Duplicate(ViewDuplicateOption.Duplicate)
	sched_change_name(doc.GetElement(first_sched_clone),n)
	first_sched = ScheduleSheetInstance.Create(doc,s.Id,first_sched_clone,point_0)
	scheds_on_sheet.append(first_sched)
	clones = []
		
	#while schedule_no <= len(master_Id)-1:
	for f in flexi_schedules:
		y_flexi = get_ymin(scheds_on_sheet[schedule_no], s)
		flexi_point = XYZ(x,y_flexi,0)
		sched_clone = doc.GetElement(f).Duplicate(ViewDuplicateOption.Duplicate)
		sched_change_name(doc.GetElement(sched_clone),n)
		sched_create = ScheduleSheetInstance.Create(doc,s.Id,sched_clone,flexi_point)
		scheds_on_sheet.append(sched_create)
		schedule_no += 1
		
	for v in views_OG:
		view_id = v.ViewId
		view = doc.GetElement(view_id)
		clone = view.Duplicate(ViewDuplicateOption.WithDetailing)
		change_name(doc.GetElement(clone),n)
		clones.append(doc.GetElement(clone))
	for v,c in zip(clones,center_points):
		if Viewport.CanAddViewToSheet(doc,s.Id,v.Id) == True:
			Viewport.Create(doc,s.Id,v.Id,c)
	

TransactionManager.Instance.TransactionTaskDone()


OUT= PAM_no_OG, PAM_nos_string