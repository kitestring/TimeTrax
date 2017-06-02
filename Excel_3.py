import xlrd
import win32com.client

class xlrd_fx():
	def __init__(self, filepath, sheet_index):
		self.wkbk = xlrd.open_workbook(filepath)
		self.sheet = self.wkbk.sheet_by_index(sheet_index)
		
	def get_cell_value(self, row, column, empty):
		'''note row and columns are represented by integers starting a zero
		example: A1: row = 0 column = 0 C3: row = 2 column = 2
		If the cell is empty then it will return what is defined by empty.''' 

		cell = self.sheet.cell(row,column)
		
		if cell.value == "":
			return empty
		else:
			return cell.value
		
	def convert_excel_date(self, excel_date, format):
		'''Takes the excel date format and returns a tuple 
		with the following date format (year, month, day, hours, minutes, seconds)
		or converts the tuple into str(yyyy-mm-dd)'''
		date_tuple =  xlrd.xldate_as_tuple(excel_date, self.wkbk.datemode)
		
		if format == 'date_tuple':
			return date_tuple
		elif format == 'yyyy-mm-dd':
			return self.date_format_into_yyyy_mm_dd(date_tuple)
		
		
	def get_cell_range_values(self, row, col_start, col_end, empty):
		#This method has not been tested as of yet
		'''Returns a range of cell values across a given row, from a
		staring column to an ending column.  Because the row_slice method
		applies the slice from col_start:col_end -1, the col_end +1 has been applied
		to simplify.  Recall that rows and columns are represented by integers starting 
		a zero.  It returns the cell values in a list as float that can be iterated.  
		Here's an example: "[number:8.0, number:6.0, number:8.0]"  This format is 
		converted to a list of floats with empty cells converted to what is defined by empty.'''
		vals = []
		cells = sheet.row_slice(row, col_start, col_end+1)
		
		for cell in cells:
			if cell.value == "":
				vals.append(empty)
			else:
				vals.append(float(cell.value))
		
		return vals
		
	def date_format_into_yyyy_mm_dd(self, date_tuple):
		'''Converts the date format (year, month, day, hours, minutes, seconds) tuple
		into yyyy-mm-dd string'''
		y = date_tuple[0]
		m = str(date_tuple[1])
		d = str(date_tuple[2])
		
		if len(m) == 1:
			m = "0%d" % date_tuple[1]
			
		if len(d) == 1:
			d = "0%d" % date_tuple[2]
			
		return '%s-%s-%s' % (y, m, d)
	
class Macros():

	def __init__(self):
		self.ConverterPath = 'C:\\ProgramData\\TimeTrax\\docs\\XLManipulation.xlsm'
		self.Macro_Prefix = "XLManipulation.xlsm!XL_2013."
		self.xl = win32com.client.Dispatch("Excel.Application")
		self.xl.Visible = False
	
	def OldFileUpgrade(self, OldXL_File_Path):
		Macro_Name = self.Macro_Prefix + "FileUpgrader"
		self.xl.Workbooks.Open(Filename=self.ConverterPath)
		self.xl.Application.Run(Macro_Name, OldXL_File_Path)
		self.xl.Quit()
		
	def BadValueMarker(self, XL_File_Path, Valid_Date, Valid_Clock_Number):
		Macro_Name = self.Macro_Prefix + "MarkBadData"
		self.xl.Workbooks.Open(Filename=self.ConverterPath)
		self.xl.Application.Run(Macro_Name, XL_File_Path, Valid_Date, Valid_Clock_Number)
		self.xl.Quit()
		
	def TransferToV2(self, XL_File_Path):
		Macro_Name = self.Macro_Prefix + "TransferTo_Version_2"
		self.xl.Workbooks.Open(Filename=self.ConverterPath)
		self.xl.Application.Run(Macro_Name, XL_File_Path)
		self.xl.Quit()
		
	def TransferToV2_FromOriginal(self, XL_File_Path):
		Macro_Name = self.Macro_Prefix + "TransferTo_Version_2_From_1"
		self.xl.Workbooks.Open(Filename=self.ConverterPath)
		self.xl.Application.Run(Macro_Name, XL_File_Path)
		self.xl.Quit()
		
class XL_Constants():
	def __init__(self):
		self.clk_no_row = 2
		self.monday_row = 1
		self.clk_no_and_monday_column = 9
		self.day_date_row = 5
		self.start_column_D = 3
		self.version_column = 11
		self.version_row = 2

	def Get_Category_Start_Row_Dict(self):
		return {
			'SS_Installations': 8,
			'GDS_Installations': 17,	
			'SS_PM_Site_Visits': 26,	
			'GDS_PM_Site_Visits': 35,		
			'SS_Rpr_Maint_Site_Visits': 44,		
			'GDS_Rpr_Maint_Site_Visits': 53,		
			'SS_Rmt_Hrdwr_Spt': 62,
			'GDS_Rmt_Hrdwr_Spt': 70,	
			'SS_Rmt_Sftwr_Spt' : 78,		
			'GDS_Rmt_Sftwr_Spt': 86,		
			'SS_Rpr_Mant_RFB_in_House': 94,		
			'GDS_Rpr_Mant_RFB_in_House': 100,	
			'Miscellaneous' : 106,		
			'SS_Doc_Gen': 111,		
			'GDS_Doc_Gen': 121,	
			'SS_Inter_Dep_Spt': 131,	
			'GDS_Inter_Dep_Spt': 139,	
			'SS_Online_Training': 147,	
			'GDS_Online_Training': 158,	
			'SS_Onsite_Training': 165,	
			'GDS_Onsite_Training': 180,	
			'SS_In_House_Training': 194,	
			'GDS_In_House_Training': 210,	
			'Validation_Duties': 225
			}