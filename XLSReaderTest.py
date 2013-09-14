from XLSReader import XLSReader

test_path = 'sample/un-country-codes.xls'

# =====================================
#  Creating and configure an XLSReader
# =====================================

#Instantiate a new XLSReader object, passing it the path of our test file
xlsr = XLSReader(test_path)
#Store a reference to a specific sheet inside of the file
test_sheet = xlsr.get_sheet_by_name('data')
#This sheet must be configured before it can become the active sheet
xlsr.set_sheet_config(test_sheet, {
	'FIELDS_ROW_INDEX': 0,   #The fields are labeled in the first row of the file
	'DATA_LOWER_INDEX': 1,   #The dataset begins on the second row of the file
	'DATA_UPPER_INDEX': 241, #The dataset ends on the 242nd row of the file
	'UNIQUE_ID_FIELD': 'un_country_code'
})

# ======================
#  Set the active sheet
# ======================

#Check that the sheet has been properly configured
print 'test_sheet is configured:', xlsr.is_configured(test_sheet)

#Set the 'data' sheet as the active sheet, and verify
xlsr.set_active_sheet(test_sheet)
print 'Active sheet name:', xlsr.src_wb_active_sheet.name

# ====================================
#  Locate the column index of a field
# ====================================

#Test that we can return column indexes for a specific field
print 'Country codes are in column', xlsr.get_col_index_by_field('un_country_code')

# =================================
#  Locate a row index by unique ID
# =================================

#Test that we can locate a row (index) given a specific unique id value
print 'Data for country no. 368 can be found in row', xlsr.get_row_index_by_uid(368)

# =====================================
#  Run a data query and return a value
# =====================================

#Finally, check that the XLSReader.query method works correctly
print 'Returning information on country no. 368:'
query = {
	'uid': 368,
	'fields': ['iso_alpha3_code', 'country_or_area_name']
}
print xlsr.query(**query)


