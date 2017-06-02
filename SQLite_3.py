import sqlite3
import os

class Database():
	def __init__(self, filepath):
		self.filepath = filepath
		self.conn = sqlite3.connect(self.filepath)
		self.cur = self.conn.cursor()
		
	def Get_Tables(self):
		#returns a tuple list with all the table names from a given db connection
		self.cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
		return self.cur.fetchall()
		
	def Get_Tables_Dict(self):
		#returns a library of category table names as keys and 
		#category tables names with no "_" as values
		tables = self.Get_Tables()
		table_lib = {}
		for table in tables:
			key = "%s" % table
			value = key.replace('_',' ')
			value = value.replace('and','&')
			table_lib[str(key)] = str(value)
		return table_lib
	
	def Get_Columns(self, table):
		#returns a tuple list with all the column names from a given db connection
		column_query = self.conn.execute('SELECT * from %s' % table)
		return [description[0] for description in column_query.description]
	
	def Select_Query(self, columns, table, conditions):
		return self.conn.execute("SELECT %s FROM %s %s" % (columns, table, conditions))
		
	def Select_Query_Column_Summation(self, column, table, condition, condition_value, date1, date2):
		select_query_statement = "SELECT SUM(%s) FROM %s WHERE %s %s AND data_date BETWEEN %s AND \
			'%s'" % (column, table, condition, tuple(condition_value), tuple(date1), tuple(date2))
		return self.conn.execute(select_query_statement)
		
	def Insert_Query_No_Conditions(self, table, columns, values):
		self.conn.execute("INSERT INTO %s %s VALUES %s" % (table, tuple(columns), tuple(values)))
		#self.conn.commit()
		
	def Delete_Query_With_Conditions(self, table, conditions):
		self.conn.execute("DELETE FROM %s WHERE %s" % (table, conditions))
		self.conn.commit()
			
	def Create_Employee_Tables(self):
		self.conn.execute('''
				CREATE TABLE employee(
				clk_no TEXT PRIMARY KEY,
				employee_name TEXT,
				employee_type TEXT)
				''')
				
		self.conn.execute('''
				CREATE TABLE upload_dates(
				employee_clk_no TEXT,
				mondays_date TEXT)
				''')
	
	def Update_Employee_Tables(self, department):
		table = 'employee'
		columns = ['employee_name', 'clk_no', 'employee_type']
		if department == "SepSci":
			values = [['Jones, Tom', '0743','In-House'],['Leetz, Edward','0791','Field'],['Rowe, Christopher','0227','In-House'],
					['Card, Wayne','0767','In-House'],['Clifford, Ray','0798','Field'],['Frazee, Joseph','0196','In-House'],
					['Hasty, Matthew','0938','Field'],['Immoos, Chris','1076','In-House'],['Jeka, Angelia','3012','N/A'],
					['Kite, Kenneth','0912','N/A'],['Norlock, Samantha','3013','N/A'],['Paulaski, James','0923','Field'],
					['Rich, Kiely','0192','In-House'],['Richards, Todd','0889','N/A'],['Toms, Anthony','0888','In-House'],
					['Miniard, Brittany','0198','N/A']]

		elif department == "GDS":
			values = [['Jones, Tom', '0743','In-House'],['Rowe, Christopher','0227','In-House'],
					['Bergen, Joseph','0953','Field'],['Johns, Ben','0362','Field']]
				
		for v in values:
			self.Insert_Query_No_Conditions(table, columns, v)
	
	def Create_Category_Tables(self):
		
		self.conn.execute('''
					CREATE TABLE SS_Installations(
					data_date TEXT,
					employee_clk_no TEXT,
					Installations_day_sum INTEGER,
					Installations_Prep INTEGER,
					Documentation_Admin INTEGER,
					Travel INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Training_by_Engineer INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Installations(
					data_date TEXT,
					employee_clk_no TEXT,
					Installations_day_sum INTEGER,
					Installations_Prep INTEGER,
					Documentation_Admin INTEGER,
					Travel INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Training_by_Engineer INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_PM_Site_Visits(
					data_date TEXT,
					employee_clk_no TEXT,
					PM_Site_Visits_day_sum INTEGER,
					PM_Prep INTEGER,
					Documentation_Admin INTEGER,
					Travel INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Training_by_Engineer INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_PM_Site_Visits(
					data_date TEXT,
					employee_clk_no TEXT,
					PM_Site_Visits_day_sum INTEGER,
					PM_Prep INTEGER,
					Documentation_Admin INTEGER,
					Travel INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Training_by_Engineer INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Rpr_Maint_Site_Visits(
					data_date TEXT,
					employee_clk_no TEXT,
					Inst_Repair_or_Maintenance_on_Site_day_sum INTEGER,
					Site_Visit_Prep INTEGER,
					Documentation_Admin INTEGER,
					Travel INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Training_by_Engineer INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Rpr_Maint_Site_Visits(
					data_date TEXT,
					employee_clk_no TEXT,
					Inst_Repair_or_Maintenance_on_Site_day_sum INTEGER,
					Site_Visit_Prep INTEGER,
					Documentation_Admin INTEGER,
					Travel INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Training_by_Engineer INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Rmt_Hrdwr_Spt(
					data_date TEXT,
					employee_clk_no TEXT,
					Rmt_Hardware_Support_day_sum INTEGER,
					Email INTEGER,
					Phone INTEGER,
					Rmt_PC_Access INTEGER,
					Troubleshooting INTEGER,
					Documentation_Admin INTEGER,
					International_Support INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Rmt_Hrdwr_Spt(
					data_date TEXT,
					employee_clk_no TEXT,
					Rmt_Hardware_Support_day_sum INTEGER,
					Email INTEGER,
					Phone INTEGER,
					Rmt_PC_Access INTEGER,
					Troubleshooting INTEGER,
					Documentation_Admin INTEGER,
					International_Support INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Rmt_Sftwr_Spt(
					data_date TEXT,
					employee_clk_no TEXT,
					Rmt_Software_Support_day_sum INTEGER,
					Email INTEGER,
					Phone INTEGER,
					Rmt_PC_Access INTEGER,
					Troubleshooting INTEGER,
					Documentation_Admin INTEGER,
					International_Support INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Rmt_Sftwr_Spt(
					data_date TEXT,
					employee_clk_no TEXT,
					Rmt_Software_Support_day_sum INTEGER,
					Email INTEGER,
					Phone INTEGER,
					Rmt_PC_Access INTEGER,
					Troubleshooting INTEGER,
					Documentation_Admin INTEGER,
					International_Support INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Rpr_Mant_RFB_in_House(
					data_date TEXT,
					employee_clk_no TEXT,
					Inst_Repair_Maint_Rfb_In_House_day_sum INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Documentation_Admin INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Rpr_Mant_RFB_in_House(
					data_date TEXT,
					employee_clk_no TEXT,
					Inst_Repair_Maint_Rfb_In_House_day_sum INTEGER,
					Hands_On_Instrument INTEGER,
					Troubleshooting INTEGER,
					Documentation_Admin INTEGER,
					Other INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE Miscellaneous(
					data_date TEXT,
					employee_clk_no TEXT,
					Miscellaneous_day_sum INTEGER,
					Meetings_Internal_Communication INTEGER,
					PTO_Sick_Dr_Apts INTEGER,
					SAP INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Doc_Gen(
					data_date TEXT,
					employee_clk_no TEXT,
					Document_Generation_day_sum INTEGER,
					Pegasus_HT_4D_Hardware INTEGER,
					Pegasus_HT_4D_Software INTEGER,
					Pegasus_BT_Hardware INTEGER,
					Pegasus_BT_Software INTEGER,
					Pegasus_HRT_Hardware INTEGER,
					Pegasus_HRT_Software INTEGER,
					Other_Hardware INTEGER,
					Other_Software INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Doc_Gen(
					data_date TEXT,
					employee_clk_no TEXT,
					Document_Generation_day_sum INTEGER,
					GDS500A_Hardware INTEGER,
					GDS500A_Software INTEGER,
					GDS850A_Hardware INTEGER,
					GDS850A_Software INTEGER,
					GDS900_Hardware INTEGER,
					GDS900_Software INTEGER,
					GDS950_Hardware INTEGER,
					GDS950_Software INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Inter_Dep_Spt(
					data_date TEXT,
					employee_clk_no TEXT,
					Inter_Dep_Spt_day_sum INTEGER,
					R_and_D INTEGER,
					Applications INTEGER,
					Validation INTEGER,
					Marketing INTEGER,
					Test INTEGER,
					Shipping INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Inter_Dep_Spt(
					data_date TEXT,
					employee_clk_no TEXT,
					Inter_Dep_Spt_day_sum INTEGER,
					R_and_D INTEGER,
					Applications INTEGER,
					Validation INTEGER,
					Marketing INTEGER,
					Test INTEGER,
					Shipping INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Online_Training(
					data_date TEXT,
					employee_clk_no TEXT,
					Online_Training_day_sum INTEGER,
					Scheduling INTEGER,
					Peg_HT_New_Customer INTEGER,
					Peg_4D_New_Customer INTEGER,
					Peg_BT_New_Customer INTEGER,
					Peg_HRT_New_Customer INTEGER,
					Peg_HT_Paid INTEGER,
					Peg_4D_Paid INTEGER,
					Peg_BT_Paid INTEGER,
					Peg_HRT_Paid INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Online_Training(
					data_date TEXT,
					employee_clk_no TEXT,
					Online_Training_day_sum INTEGER,
					Scheduling INTEGER,
					GDS900_New_Customer INTEGER,
					GDS950_New_Customer INTEGER,
					GDS900_Paid INTEGER,
					GDS950_Paid INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_Onsite_Training(
					data_date TEXT,
					employee_clk_no TEXT,
					Onsite_Training_day_sum INTEGER,
					Scheduling INTEGER,
					Travel INTEGER,
					Documentation_Admin INTEGER,
					Customer_Entertainment INTEGER,
					Peg_HT_4D_New_Customer INTEGER,
					Peg_BT_New_Customer INTEGER,
					Peg_HRT_New_Customer INTEGER,
					GCxGC_FID_New_Customer INTEGER,
					TruTOF_Paid INTEGER,
					Peg_HT_4D_Paid INTEGER,
					Peg_BT_Paid INTEGER,
					Peg_HRT_Paid INTEGER,
					GCxGC_FID_Paid INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_Onsite_Training(
					data_date TEXT,
					employee_clk_no TEXT,
					Onsite_Training_day_sum INTEGER,
					Scheduling INTEGER,
					Travel INTEGER,
					Documentation_Admin INTEGER,
					Customer_Entertainment INTEGER,
					GDS500A_New_Customer INTEGER,
					GDS850A_New_Customer INTEGER,
					GDS900_New_Customer INTEGER,
					GDS950_New_Customer INTEGER,
					GDS500A_Paid INTEGER,
					GDS850A_Paid INTEGER,
					GDS900_Paid INTEGER,
					GDS950_FID_Paid INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE SS_In_House_Training(
					data_date TEXT,
					employee_clk_no TEXT,
					In_House_Training_day_sum INTEGER,
					Classroom_Setup INTEGER,
					System_Checkout INTEGER,
					Customer_Entertainment INTEGER,
					Peg_HT_4D_New_Customer INTEGER,
					Peg_BT_New_Customer INTEGER,
					Peg_HRT_New_Customer INTEGER,
					GCxGC_FID_New_Customer INTEGER,
					TruTOF_Paid INTEGER,
					Peg_HT_4D_Paid INTEGER,
					Peg_BT_Paid INTEGER,
					Peg_HRT_Paid INTEGER,
					GCxGC_FID_Paid INTEGER,
					Internal_Domestic INTEGER,
					Internal_International INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE GDS_In_House_Training(
					data_date TEXT,
					employee_clk_no TEXT,
					In_House_Training_day_sum INTEGER,
					Classroom_Setup INTEGER,
					System_Checkout INTEGER,
					Customer_Entertainment INTEGER,
					GDS500A_New_Customer INTEGER,
					GDS850A_New_Customer INTEGER,
					GDS900_New_Customer INTEGER,
					GDS950_New_Customer INTEGER,
					GDS500A_Paid INTEGER,
					GDS850A_Paid INTEGER,
					GDS900_Paid INTEGER,
					GDS950_Paid INTEGER,
					Internal_Domestic INTEGER,
					Internal_International INTEGER)
					''')
					
		self.conn.execute('''
					CREATE TABLE Validation_Duties(
					data_date TEXT,
					employee_clk_no TEXT,
					Validation_Duties_day_sum INTEGER,
					Data_Collection INTEGER,
					Data_Handling INTEGER,
					Report_Generation INTEGER,
					Instrument_Troubleshooting INTEGER,
					Instrument_Maintenance INTEGER,
					Experiment_Planning_and_Method_Development INTEGER,
					Sample_Prep INTEGER,
					Remote_Instrument_Interface INTEGER,
					Administrative INTEGER)
					''')
					
	def Get_Query_Category_Table_List(self):
		return [
			'SS_Installations',
			'GDS_Installations',	
			'SS_PM_Site_Visits',	
			'GDS_PM_Site_Visits',		
			'SS_Rpr_Maint_Site_Visits',		
			'GDS_Rpr_Maint_Site_Visits',		
			'SS_Rmt_Hrdwr_Spt',
			'GDS_Rmt_Hrdwr_Spt',	
			'SS_Rmt_Sftwr_Spt',		
			'GDS_Rmt_Sftwr_Spt',		
			'SS_Rpr_Mant_RFB_in_House',		
			'GDS_Rpr_Mant_RFB_in_House',	
			'Miscellaneous',		
			'SS_Doc_Gen',		
			'GDS_Doc_Gen',	
			'SS_Inter_Dep_Spt',	
			'GDS_Inter_Dep_Spt',	
			'SS_Online_Training',	
			'GDS_Online_Training',	
			'SS_Onsite_Training',	
			'GDS_Onsite_Training',	
			'SS_In_House_Training',	
			'GDS_In_House_Training',	
			'Validation_Duties'
			]
			
	def Get_SS_Query_Category_Table_List(self):
		return [
			'SS_Installations',
			'SS_PM_Site_Visits',			
			'SS_Rpr_Maint_Site_Visits',				
			'SS_Rmt_Hrdwr_Spt',	
			'SS_Rmt_Sftwr_Spt',			
			'SS_Rpr_Mant_RFB_in_House',			
			'Miscellaneous',		
			'SS_Doc_Gen',			
			'SS_Inter_Dep_Spt',		
			'SS_Online_Training',		
			'SS_Onsite_Training',		
			'SS_In_House_Training',		
			'Validation_Duties'
			]
			
	def Get_GDS_Query_Category_Table_List(self):
		return [
			'GDS_Installations',	
			'GDS_PM_Site_Visits',			
			'GDS_Rpr_Maint_Site_Visits',		
			'GDS_Rmt_Hrdwr_Spt',		
			'GDS_Rmt_Sftwr_Spt',			
			'GDS_Rpr_Mant_RFB_in_House',	
			'Miscellaneous',				
			'GDS_Doc_Gen',		
			'GDS_Inter_Dep_Spt',		
			'GDS_Online_Training',		
			'GDS_Onsite_Training',	
			'GDS_In_House_Training',	
			]
			
	def Get_Diplay_Category_Table_Dict(self):
		return {
			'SS_Installations': 'Installations',
			'GDS_Installations': 'Installations',	
			'SS_PM_Site_Visits': 'PM Site Visits',	
			'GDS_PM_Site_Visits': 'PM Site Visits',		
			'SS_Rpr_Maint_Site_Visits' :'Site Visits',		
			'GDS_Rpr_Maint_Site_Visits' :'Site Visits',		
			'SS_Rmt_Hrdwr_Spt': 'Remote Hardware Support',
			'GDS_Rmt_Hrdwr_Spt': 'Remote Hardware Support',	
			'SS_Rmt_Sftwr_Spt' : 'Remote Software Support',		
			'GDS_Rmt_Sftwr_Spt': 'Remote Software Support',		
			'SS_Rpr_Mant_RFB_in_House': 'In House Maintenance',		
			'GDS_Rpr_Mant_RFB_in_House': 'In House Maintenance',	
			'Miscellaneous' : 'Miscellaneous',		
			'SS_Doc_Gen': 'Document Creation',		
			'GDS_Doc_Gen': 'Document Creation',	
			'SS_Inter_Dep_Spt': 'Inter Departmental Spt.',	
			'GDS_Inter_Dep_Spt': 'Inter Departmental Spt.',	
			'SS_Online_Training': 'Online Training',	
			'GDS_Online_Training': 'Online Training',	
			'SS_Onsite_Training': 'Onsite Training',	
			'GDS_Onsite_Training': 'Onsite Training',	
			'SS_In_House_Training': 'In House Training',	
			'GDS_In_House_Training': 'In House Training',	
			'Validation_Duties': 'Validation Duties'
			}
	
	def Get_Category_Indexes(self, type):
		if type == "AllSepSci":
			return [0,2,4,6,8,10,12,13,15,17,19,21,23]
		elif type == "AllGDS":
			return [1,3,5,7,9,11,12,14,16,18,20,22]