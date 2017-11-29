from read_excel import *
import numpy as np
import xlwt
import re

# change the weight of evalutaion
pref1 = 1
pref2 = .5
pref3 = .25
base = .1
TA_Weight = 1
Prof_Weight = 1.1
# convert ta units to hours
UNITS_TO_HOURS= 1



# create a class to store all of the class info for determining fitness
class course():
    def __init__(self,course, TA_Units, Lab_Units):
        self.course = course
        self.TA_Units = TA_Units*UNITS_TO_HOURS
        self.Lab_Units = Lab_Units*UNITS_TO_HOURS
        #import professor inputs
        PROF_INPUT= Prof_input()

        self.P_P1 = 0
        self.P_P2 = 0
        self.P_P3 = 0
        #check the professors tutorial preference for the given course
        for p in PROF_INPUT:
            if self.course == p.course and p.role_1 == "Course TA":
                self.P_P1 = p.first_pref
                self.P_P2 = p.second_pref
                self.P_P3 = p.third_pref
                break

        self.LP_P1 = 0
        self.LP_P2 = 0
        self.LP_P3 = 0
        #check the professors Lab preference for the give course
        if self.Lab_Units > 0:
            for p in PROF_INPUT:
                if self.course == p.course and p.role_1 == "Lab TA":
                    self.LP_P1 = p.first_pref
                    self.LP_P2 = p.second_pref
                    self.LP_P3 = p.third_pref
                    break

        self.T_P1 = []#T_P1
        self.T_P2 = []#T_P2
        self.T_P3 = []#T_P3
        self.LT_P1 = []#LT_P1
        self.LT_P2 = []#LT_P2
        self.LT_P3 = []#LT_P3
        #import the TA inputs
        TA_INPUT= TA_input()
        #aggregate all of the TAs that prefer this course
        for ta in TA_INPUT:
            if self.TA_Units > 0:
                if ta.first_pref == self.course and re.match('^Course',ta.role_1) :
                    self.T_P1.append(ta)

                elif ta.second_pref == self.course and re.match('^Course',ta.role_2):

                    self.T_P2.append(ta)

                elif ta.third_pref == self.course and re.match('^Course',ta.role_3):
                    self.T_P3.append(ta)

            if self.Lab_Units > 0:
                if ta.first_pref == self.course and re.match('^Lab',ta.role_1):
                    self.LT_P1.append(ta)

                elif ta.second_pref == self.course and re.match('^Lab',ta.role_2):
                    self.LT_P2.append(ta)

                elif ta.third_pref == self.course and re.match('^Lab',ta.role_3):
                    self.LT_P3.append(ta)

        self.fitness = 0
        self.TutorialChoice = 0
        self.LabChoice = 0

    def find_fitness(self,Avail_TA):
        best_fitness = 0
        best_index = np.random.randint(len(Avail_TA.availible_TAs))
        #check every Availible TA for their fit to the course
        for i in range(len(Avail_TA.availible_TAs)):
            temp_fit = 0
            if Avail_TA.availible_TAs[i].Student_ID == self.P_P1 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:

                temp_fit += 1*pref1*Prof_Weight

            elif Avail_TA.availible_TAs[i].Student_ID == self.P_P2 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                temp_fit += 1*pref2*Prof_Weight

            elif Avail_TA.availible_TAs[i].Student_ID == self.P_P3 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                temp_fit += 1*pref3*Prof_Weight
            for j in range(len(self.T_P1)):
                if Avail_TA.availible_TAs[i].Student_ID == self.T_P1[j].student_id and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                    temp_fit += 1*pref1*TA_Weight
                    break

            for j in range(len(self.T_P2)):
                if Avail_TA.availible_TAs[i].Student_ID == self.T_P2[j].student_id and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                    temp_fit += 1*pref2*TA_Weight
                    break

            for j in range(len(self.T_P3)):
                if Avail_TA.availible_TAs[i].Student_ID == self.T_P3[j].student_id and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                    temp_fit += 1*pref3*TA_Weight
                    break

            if temp_fit > best_fitness:
                best_fitness = temp_fit
                best_index = i

        if best_fitness == 0:
            while(True):
                best_index = np.random.randint(len(Avail_TA.availible_TAs))
                if Avail_TA.availible_TAs[best_index].Hours_owed > self.TA_Units:
                    best_fitness += 1*base
                    break
        #take the best fit and choose that TA
        self.TutorialChoice = Avail_TA.availible_TAs[best_index]
        self.fitness += best_fitness
        #remove that student from the list if they have been used twice or they dont owe hours
        Avail_TA.Use_Student(best_index, self.TA_Units)
        best_fitness = 0
        #same thing for labs
        if self.Lab_Units > 0:
            for i in range(len(Avail_TA.availible_TAs)):
                temp_fit = 0
                if Avail_TA.availible_TAs[i].Student_ID == self.LP_P1 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                    temp_fit += 1 * pref1 * Prof_Weight

                elif Avail_TA.availible_TAs[i].Student_ID == self.LP_P2 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                    temp_fit += 1 * pref2 * Prof_Weight

                elif Avail_TA.availible_TAs[i].Student_ID == self.LP_P3 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                    temp_fit += 1 * pref3 * Prof_Weight

                for j in range(len(self.LT_P1)):
                    if Avail_TA.availible_TAs[i].Student_ID == self.LT_P1[j] and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                        temp_fit += 1 * pref1 * TA_Weight
                        break

                for j in range(len(self.LT_P2)):
                    if Avail_TA.availible_TAs[i].Student_ID == self.LT_P2[j] and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                        temp_fit += 1 * pref2 * TA_Weight
                        break

                for j in range(len(self.LT_P3)):
                    if Avail_TA.availible_TAs[i].Student_ID == self.LT_P3 and Avail_TA.availible_TAs[i].Hours_owed > self.TA_Units:
                        temp_fit += 1 * pref3 * TA_Weight
                        break

                if temp_fit > best_fitness:
                    best_fitness = temp_fit
                    best_index = i

            if best_fitness == 0:
                while (True):
                    best_index = np.random.randint(len(Avail_TA.availible_TAs))
                    if Avail_TA.availible_TAs[best_index].Hours_owed > self.TA_Units:
                        best_fitness += 1 * base
                        break

            self.LabChoice = Avail_TA.availible_TAs[best_index]
            self.fitness += best_fitness
            Avail_TA.Use_Student(best_index, self.TA_Units)

        else:
            #if there is no lab then double the fitness of the TA match
            self.fitness *= 2


class TA():
    #save info for output and control of available tas
    def __init__(self, Student_ID, Hours_owed):
        self.Student_ID = Student_ID
        self.Hours_owed = Hours_owed
        self.times_used = 0

# a class that holds all of the tas used or available
class Availible_TA():

    def __init__(self):
        self.availible_TAs = []
        all_TAs = TA_Info()
        for i in range(len(all_TAs)):
            temp = TA(all_TAs[i].student_id, all_TAs[i].HRS_OWED)
            self.availible_TAs.append(temp)
        self.TAs_Used = []
    #update ta after use and remove from list if has been used twice or no hours are owed
    def check_used(self):
        for i in range(len(self.availible_TAs)):
            if self.availible_TAs[i].times_used == 2:
                self.TAs_Used.append(self.availible_TAs[i])
                del self.availible_TAs[i]
                break
            if self.availible_TAs[i].Hours_owed == 0:
                self.TAs_Used.append(self.availible_TAs[i])
                del self.availible_TAs[i]
                break

    def Use_Student(self,i,TA_Units):
        self.availible_TAs[i].Hours_owed -= TA_Units
        self.availible_TAs[i].times_used += 1
        self.check_used()

#run algorithm and write to file
if __name__ == "__main__":
    solutions = []
    avail_TA = Availible_TA()
    courses = Course_Info()
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Results")
    ws.write(0, 0, "Course")
    ws.write(0, 1, "TA_Units")
    ws.write(0, 2, "Lab_Units")
    ws.write(0, 3, "Fitness")
    ws.write(0, 4, "TA_Choice id")
    ws.write(0, 5, "TA_Choice hours owed")
    ws.write(0, 6, "Lab Choice")
    ws.write(0, 7, "Lab_Choice hours owed")
    for i in range(0,len(courses),1):
        temp = course(courses[i].Course, courses[i].TA_Units, courses[i].Lab_Units)
        temp.find_fitness(avail_TA)
        solutions.append(temp)
        ws.write(i+1,0, solutions[i].course)
        ws.write(i + 1, 1, solutions[i].TA_Units)
        ws.write(i + 1, 2, solutions[i].Lab_Units)
        ws.write(i + 1, 3, solutions[i].fitness)
        ws.write(i + 1, 4, solutions[i].TutorialChoice.Student_ID)
        ws.write(i + 1, 5, solutions[i].TutorialChoice.Hours_owed)
        if solutions[i].LabChoice != 0:
            ws.write(i + 1, 6, solutions[i].LabChoice.Student_ID)
            ws.write(i + 1, 7, solutions[i].LabChoice.Hours_owed)
        else:
            ws.write(i + 1, 6, 0)
            ws.write(i + 1, 7, 0)
    wb.save("TA_assignment.xls")








