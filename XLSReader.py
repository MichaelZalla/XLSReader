from xlrd import open_workbook

'''
This class may be used to read values from datasets recorded inside of Excel
workbook files (XLS). A single XLSReader object may be used to read values from
each sheet contained in the file.

To read data intelligently from a sheet, the user of this class must specify
a number of variables, including FIELDS_ROW_INDEX, DATA_LOWER_INDEX, and 
DATA_UPPER_INDEX. These values provide the XLSReader object with enough information
to locate specific values given a column (field) name and a unique row ID.
'''
class XLSReader(object):

    #Static constants
    CONFIG_PROPERTIES = [
        'FIELDS_ROW_INDEX',
        'DATA_LOWER_INDEX',
        'DATA_UPPER_INDEX',
        'UNIQUE_ID_FIELD'
    ]

    #Reader data
    src_wb = None
    src_wb_sheets = [ ]
    src_wb_sheets_config = { }
    src_wb_active_sheet = None

    '''
    Create a new XLSReader object by passing a file path to its constructor. The
    constructor will attempt to create an XLRD Workbook object with the given path.
    If this process fails, the XLSReader object will output an exception message
    and destroy itself.
    '''
    def __init__(self, filepath):
        #Attempt to create an XLRD Workbook object using the given filepath
        try:
            self.src_wb = open_workbook(filepath)
            self.src_wb_sheets = self.src_wb.sheets()
            #Populate src_wb_sheets_config with config objects
            for sheet in self.src_wb_sheets:
                #Create a new config object for the sheet
                self.src_wb_sheets_config[sheet.name] = { }
                #Set all required config properties to 'None'
                for prop in self.CONFIG_PROPERTIES:
                    self.src_wb_sheets_config[sheet.name][prop] = None
            #Set the first sheet in the file as the active sheet by deafult
            self.src_wb_active_sheet = self.src_wb_sheets[0]
        #Destruct if no workbook could be created
        except (NameError, IOError, XLRDError) as e:
            #print e.strerror
            raise Exception("Error: No readable Excel file at '" + filepath + "'.")
            del self

    '''
    Returns an XLRD sheet object matching the name provided. If no matching sheet
    is found, this method will return None.
    '''
    def get_sheet_by_name(self, name):
        #For each XLRD sheet object within src_wb_sheets
        for sheet in self.src_wb_sheets:
            if sheet.name.lower() == name.lower():
                return sheet
        #At this point, if no match has been found, raise an Exception
        raise Exception("Error: Sheet '" + sheet.name + "' not found in workbook.")

    '''
    Allows the user to pass in a sheet and an associated 'config' object containing
    a number of properties describing the sheet. These properties include:

        FIELDS_ROW_INDEX    The index of the row containing descriptive field labels
        DATA_LOWER_INDEX    The lower boundary (first row) index of the data set
        DATA_UPPER_INDEX    The upper boundary (last row) index of the data set
        UNIQUE_ID_FIELD     The name of a field containing only unique values
    
    If one or more of these properties are missing from the config object,
    the method will raise an exception, and no config data is written to the
    self.src_wb_sheets_config object.
    '''
    def set_sheet_config(self, sheet, config):
        #Check that the config object contains all necessary values
        for prop in self.CONFIG_PROPERTIES:
            if prop not in config:
                error_msg  = 'Error: Config object is missing the following property: '
                error_msg += prop
                raise Exception(error_msg)
            else:
                #We can assume that the src_wb_sheets_config object contains a key for
                # the given sheet, as these are created in the object's constructor method
                # These values must be set inside a sheet's config object before values
                # or queries can be returned.
                self.src_wb_sheets_config[sheet.name][prop] = config[prop]

    '''
    Returns a boolean value indicating whether a specified sheet has been
    properly configured.
    '''
    def is_configured(self, sheet):
        for prop in self.CONFIG_PROPERTIES:
            #Check that each of the properties has been set
            if prop not in self.src_wb_sheets_config[sheet.name]:
                return False
        return True;

    '''
    Sets src_wb_active_sheet to the given sheet object, if it can be found within
    the XLSReader object's workbook file. Can be used in conjunction with the
    XLSReader.get_sheet_by_name function.
    '''
    def set_active_sheet(self, sheet):
        #Only allow configured sheets to become active
        if sheet in self.src_wb_sheets and self.is_configured(sheet):
            self.src_wb_active_sheet = sheet
        else:
            raise Exception("Error: The specific sheet could not be found in the workbook.")

    '''
    Takes a single sheet and returns a list of column (field) names. This method requires that
    the specified sheet has already been configured.
    '''
    def get_fields(self, sheet = None):
        if sheet is None:
            sheet = self.src_wb_active_sheet
        if self.is_configured(sheet):
            fields = [ ]
            #By now, we can assume that the sheet's FIELDS_ROW_INDEX value has been set
            for column_index in range(sheet.ncols):
                fri = self.src_wb_sheets_config[sheet.name]['FIELDS_ROW_INDEX']
                fields.append(str(sheet.cell(fri, column_index).value))
            return fields
        else:
            raise Exception("Error: A sheet must be configured before attempting to read it!")

    '''
    Technically, a sheet will have to be configured before this method may be called,
    as required by XLSReader.get_fields(), which is called within this method. However,
    this method allows the user to re-defined the current 'UNIQUE_ID_FIELD' for a sheet
    between queries (in the case that multiple possible unique fields are available).
    If no sheet is specified, the method will use the current active sheet.
    '''
    def set_unique_id_field(self, field_name, sheet = None):        
        if sheet is None:
            sheet = self.src_wb_active_sheet
        if field_name.lower() in [f.lower() for f in self.get_fields()]:
            self.src_wb_sheets_config[sheet.name]['UNIQUE_ID_FIELD'] = field_name.lower()

    '''
    Returns a column index associated with a given field name. If no sheet
    is specified, the method will use the current active sheet.
    '''
    def get_col_index_by_field(self, field_name, sheet = None):
        if sheet is None:
            sheet = self.src_wb_active_sheet        
        #In case the target field is passed in as an int (e.g. - 1997)
        field_name = str(field_name)
        #Iterate over each field and return the correct index
        for column_index in range(sheet.ncols):
            fri = self.src_wb_sheets_config[sheet.name]['FIELDS_ROW_INDEX']
            if sheet.cell(fri, column_index).value.lower() == field_name.lower():
                return column_index
        return None

    '''
    Returns a row index associated with a given value contained in the current
    'unique id' column (field). Because the 'unique id' column should only contain
    unique values, there should never be more than one matching row. If no sheet
    is specific, the method will use the current active sheet.
    '''
    def get_row_index_by_uid(self, target_id, sheet = None):
        #Get the index of the 'unique id' column
        if sheet is None:
            sheet = self.src_wb_active_sheet
        #Return the name of the current 'unique id' field
        uid_field = self.src_wb_sheets_config[sheet.name]['UNIQUE_ID_FIELD']
        #Store the index of the 'unique id' column
        uid_column_index = self.get_col_index_by_field(uid_field)
        #Determine the data range that we'll be looking in
        lower_bound = self.src_wb_sheets_config[sheet.name]['DATA_LOWER_INDEX']
        upper_bound = self.src_wb_sheets_config[sheet.name]['DATA_UPPER_INDEX']
        #Locate the matching row index
        for row_index in range(lower_bound, upper_bound + 1):
            value = sheet.cell(row_index, uid_column_index).value
            if value == target_id:
                return row_index
        return None

    '''
    Allows the user to return specific data values for the given fields and a given
    row (as defined by a unique ID). If no matching row is found, this method will
    return None. Queries can be expressed as dictionaries, and passed to the method
    using Python's built-in argument-unpacking idiom (**kwargs).

    Example:

    query = {
        'uid': 231,
        'fields': ['iso_alpha3_code', 'country_or_area_name'],
        'sheet': self.src_wb_active_sheet,
    }
    print reader.query(**data_query)
    '''

    def query(self, uid, fields = None, sheet = None):
        ret = { }
        if sheet is None:
            sheet = self.src_wb_active_sheet
        #Defaults to returning all field values for the row
        if fields is None:
            fields = self.get_fields(sheet)
        #Allows the user to pass in a single field as a string
        if type(fields) is 'str':
            fields = [fields]
        #Determine the row index according to the UID provided
        row_index = self.get_row_index_by_uid(uid)
        if row_index is not None:
            #Iterate over each requested field and add its value to ret
            for field in fields:
                column_index = self.get_col_index_by_field(field)
                if column_index is not None:
                    ret[field] = str(sheet.cell(row_index, column_index).value)
            #Return the resulting data object
            return ret
        return None

    '''
    Allows the user to specify a partial 'view' of the data that will be 'visible'
    when the XLSReader executes a value query or searches for a row index.
    '''
    def set_visible_rows(self, lower_bound, upper_bound, sheet = None):
        if sheet is None:
            sheet = self.src_wb_active_sheet
        if lower_bound > 0 and upper_bound > lower_bound:
            self.src_wb_sheets_config[sheet.name]['DATA_LOWER_INDEX'] = lower_bound
            self.src_wb_sheets_config[sheet.name]['DATA_UPPER_INDEX'] = upper_bound
        
