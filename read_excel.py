from xlrd import open_workbook

class Ta_info(object):
    def __init__(self, student_id, TA_HRS,HRS_OWED):
        self.student_id = student_id
        self.TA_HRS = TA_HRS
        self.HRS_OWED = HRS_OWED


    def __str__(self):
        return("TA_info:\n"
               "  student_id = {0}\n"
               "  TA_HRS = {1}\n"
               "  HRS_OWED = {2}\n"

               .format(self.student_id, self.TA_HRS, self.HRS_OWED))
def TA_Info():
    wb = open_workbook('TA Student Info.xls')
    items = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                values.append(value)
            item = Ta_info(*values)
            items.append(item)
    return items

class Course_info(object):
    def __init__(self, Course, Course_TA_total, TA_Units, Lab_Units):
        self.Course = Course
        self.Course_TA_total = Course_TA_total
        self.TA_Units = TA_Units
        self.Lab_Units = Lab_Units

    def __str__(self):
        return("Course_Info:\n"
               "  Course = {0}\n"
               "  Course_TA_total = {1}\n"
               "  TA_Units = {2}\n"
               "  Lab_Units = {3}\n"

               .format(self.Course,self.Course_TA_total, self.TA_Units, self.Lab_Units))

def Course_Info():
    wb = open_workbook('Course Info.xls')
    items = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if sheet.cell_type(row,col) == 0:
                    value = 0
                values.append(value)
            item = Course_info(*values)
            items.append(item)
    return items

class TA_Input(object):
    def __init__(self, student_id, first_pref,role_1, TAd1,Taken1, second_pref, role_2,TAd2,Taken2, third_pref, role_3,TAd3,Taken3):
        self.student_id = student_id
        self.first_pref = first_pref
        self.role_1 = role_1
        self.TAd1 = TAd1
        self.Taken1 = Taken1
        self.second_pref = second_pref
        self.role_2 = role_2
        self.TAd2 = TAd2
        self.Taken2 = Taken2
        self.third_pref = third_pref
        self.role_3 = role_3
        self.TAd3 = TAd3
        self.Taken3 = Taken3


    def __str__(self):
        return("TA_input:\n"
               "  student_id = {0}\n"
               "  first_pref = {1}\n"
               "  role_1 = {2}\n"
               "  TAd1 = {3}\n"
               "  Taken1 = {4}\n"
               "  second_pref = {5}\n"
               "  role_2 = {6}\n"
               "  TAd2 = {7}\n"
               "  Taken2 = {8}\n"
               "  third_pref = {9}\n"
               "  role_3 = {10}\n"
               "  TAd3 = {11}\n"
               "  Taken3 = {12}"


               .format(self.student_id, self.first_pref, self.role_1,self.TAd1, self.Taken1,
                       self.second_pref, self.role_2, self.TAd2, self.Taken2,
                       self.third_pref, self.role_3, self.TAd3, self.Taken3))
def TA_input():
    wb = open_workbook('TA Input.xls')
    items = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                if sheet.cell_type(row,col) == 0:
                    value = 0
                values.append(value)
            item = TA_Input(*values)
            items.append(item)
    return items

class Prof_Input(object):
    def __init__(self, course, first_pref,role_1, second_pref, role_2, third_pref, role_3):
        self.course = course
        self.first_pref = first_pref
        self.role_1 = role_1
        self.second_pref = second_pref
        self.role_2 = role_2
        self.third_pref = third_pref
        self.role_3 = role_3

    def __str__(self):
        return("Profs Input:\n"
               "  course = {0}\n"
               "  first_pref = {1}\n"
               "  role_1 = {2}\n"
               "  second_pref = {3}\n"
               "  role_2 = {4}\n"
               "  third_pref = {5}\n"
               "  role_3 = {6}\n"

               .format(self.course, self.first_pref, self.role_1, self.second_pref, self.role_2, self.third_pref, self.role_3))
def Prof_input():
    wb = open_workbook('Profs Input.xls')
    items = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        for row in range(1, number_of_rows):
            values = []
            for col in range(number_of_columns):
                value  = (sheet.cell(row,col).value)
                if sheet.cell_type(row,col) == 0:
                    value = 0
                values.append(value)
            item = Prof_Input(*values)
            items.append(item)
    return items




if __name__ == "__main__":
    items = TA_Info()
    for item in items:
        print item
        print("Accessing one single value (eg. student_id): {0}".format(item.student_id))
        print (format(item.student_id))
    blurns = Course_Info()
    for blurn in blurns:
        print blurn
        print
    balls = TA_input()
    for ball in balls:
        print ball
        print
    benders = Prof_input()
    for bender in benders:
        print bender
        print
