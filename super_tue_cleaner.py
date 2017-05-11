#!/usr/bin/env python
# Bike Count Data Cleaner
# Convert the Moreland Super Tuesday Excel Data into a CSV file, for use in mapping
# By: Simon Stainsby
# Github Username: FunOnTheUpfield
# On behalf of Moreland City Council Transport Unit
# Github Usersname: MorelandTransport
# Created: 25 March 2017
# Last updated: 10 May 2017
# -----------------------------

# -----------------------------
# Step 1
# Scrape excel spreadsheet to collect observational Data
# Format as 'clean data' in portable low overhead data format (.csv)
# Seperate observations from calcuated data

# Desired output
# One file per count location saved in  ./script_output/count_observations
# In .csv format with filename /[Sheet_Name]/[Count_Date_YYYY_MM_DD]
# Count location details in several header rows

# ______________________________

import os
import collections
from datetime import date,datetime,time

try:
    from xlrd import open_workbook,xldate_as_tuple
except:
    print("Install python module xlrd.  Available from https://pypi.python.org/pypi/xlrd")
    exit()

# - functions  -
def sum_with_na (dic):
    na_sum = 0
    summation = 0
    for k in dic:
        try:
            summation = summation + int(dic[k])
        except:
            na_sum = na_sum + 1
        if na_sum == len(dic):
            # If all cells contain NA, return NA, rather than 0
            summation = 'NA'
        else:
            pass
    return summation


def sum_observations (direction_dic, countdic):
    # 'direction' is a string, other imputs are dictionaries
    # Returns a dictionary (direction[gender][timestamp], directionsum)

    wholecount = {}
    for gender in countdic:
        genderdic = {}
        for starttime in countdic[gender]:
            turn_nacount = 0
            turnsum = 0
            # Count up the eight 15min bins in the 7 - 9 am observation period.
            for v in direction_dic:
                #Directions in a given 15min observation bin.
                try:
                    turnsum = turnsum + int(countdic[gender][starttime][v])
                except:
                    # bin contents is not a number (str 'NA')
                    turn_nacount = turn_nacount +1

                if turn_nacount == len(direction_dic):
                    # The NAs are only intersting if every value = NA
                    turnsum = 'NA'
                else:
                    pass
                #turnsum here = total of all relevant turns in a 15min bin for a given gender (either an int or str 'NA')
            genderdic[starttime] = turnsum

            timesum = sum_with_na(genderdic)
            wholecount[gender] = timesum
    counttotal = sum_with_na(wholecount)
    return counttotal

# -  functions end --





# Create a directory (if needed) for script output files
resultsdir = "./script_output/count_observations/"
if not os.path.exists(resultsdir):
    os.makedirs(resultsdir)

# Open the source data excel spreadsheet
inputfilename    = "Traffic Count - Bicycle Count - Bike count - \
Morning Peak 7am to 9am - Weekday - Super ~ 2017.XLSX"
workbook	= open_workbook(inputfilename, on_demand=True)
print ' Opening ', inputfilename

# Source file is a multiple worksheet excel file. One count site per sheet, mulitple counts on each sheet.
# Count observations are recorded on work sheets (pythonic)6 'BW-CityLinkBrunswickRd' to 100 'MerriCrkTrailWestRingRdTrail'
for worksheet_num in range(6, 101):  # From 6 to 101
    sheet = workbook.sheet_by_index(worksheet_num)
    countsite = sheet.name
    print countsite

    # Collect location information.
    # Each worksheet contains a 'site details block' in (excel) rows 1 to 11

    # site_description stored in excel cell C1 (pythonic 1,2). A text string that may contain commas
    site_description = sheet.cell(1,2).value
    site_description = site_description.replace(',', '')

    # suburb stored in excel cell C2 (pythonic 2,2). A text string without commas
    suburb = sheet.cell(2,2).value

    # Distance from CBD is stored in excel cell C4 (pythonic 4,2). A decimal number with no more than 2 significant figures.
    dist_from_cbd = str(sheet.cell(4,2).value)

    #  GIS reference (Coordinate Reference System = GDA 94 MGA Zone 55 http://spatialreference.org/ref/epsg/gda94-mga-zone-55/)
    # Easting is stored in excel cell H4 (pythonic 3,7). A decimal number with no more than 2 significant figures.
    easting = str(sheet.cell(3,7).value)
    # Northing (GDA 94 MGA Zone 55 Coordinate Reference System) is stored in excel cell L4 (pythonic 3,11) . A decimal number with no more than 2 significant figures.
    northing = str(sheet.cell(3,11).value)

    # Melway Map Grid Reference is stored in excel cell C4 (pythonic 3,2). A sting no longer than 7 characters.
    # Note: Some spreadsheets may attempt to render a Melway grid reference like "24 E10" as a number in exponential notation.
    melway_ref = sheet.cell(3,2).value

    # The primary road is stored in excel cell D6 (pythonic 5,3). A string that may contain commas
    primary_road = sheet.cell(5,3).value
    primary_road = primary_road.replace(',', '')

    # The secondary road is stored in excel cell M6 (pythonic 5,12). A string that may contain commas
    secondary_road = sheet.cell(5,12).value
    secondary_road = secondary_road.replace(',', '')

    sitedic = collections.OrderedDict()

    # Collect details from each count.
    for count_row in (92, 124, 156, 188, 221, 253, 285):
        # The counts are recorded in blocks commencing on (excel)rows  93, 125, 157, 189, 221, 253, 285

        # Collect count date
        excel_format_count_date = ""
        try:
            #    The Count Date field is (first row of data block, column c)
            excel_format_count_date = float(sheet.cell(count_row,2).value)
        except:
            pass

        if excel_format_count_date != "":
            #   A count data block only contains bin data if the Count Date field is populated
            #   Skip a block if there is not value in Count Date

            # Excel has its own date format, convert to YYYY-MM-DD
            preformatted_date = xldate_as_tuple(excel_format_count_date,workbook.datemode)
            formatted_date = date(*preformatted_date[0:3])
            str_formatted_date = str(formatted_date)
            survey_year = preformatted_date[0]

            # Create directory (if needed) and empty text file for details of this day's count.
            savepoint = resultsdir + countsite + "/" + str(formatted_date) + "/"

            if not os.path.exists(savepoint):
                os.makedirs(savepoint)
            output_file = open(savepoint + countsite + str_formatted_date + ".csv", "w")

            # First line of results file, count site and date of count
            output_file.write(countsite + "-" + str_formatted_date + '\n')
            output_file.write('\n')


            # Third line of results file, a header row for site information
            output_file.write('countsite, site_description, suburb, \
dist_from_cbd, easting, northing, melway_ref, primary_road, secondary_road' + '\n')

            # Fourth line of results file, the site information.
            output_file.write(countsite + "," + site_description + "," + suburb \
            + "," + dist_from_cbd + "," + easting + "," + northing + "," \
            + melway_ref + "," + primary_road + "," + secondary_road + '\n')
            output_file.write('\n')


            # Collect details specific to an given count date.
            # Collect bin duration. Stored in second row, colunn K. An integer.
            try:
                bin_duration = int(sheet.cell(count_row+1,10).value)
            except:
                # Values should be either 15 or 120. Data source contains errors.
                # Some bin_duration fields that should contain the value 15 have been left blank.
                # Specify bin_duration as 15 if data is missing
                bin_duration = 15

            # Collect gender split. Stored in second row, column O. A Booleen string, either Y or N
            gender_split = sheet.cell(count_row+1,14).value

            # Collect Counter details (where provided).  Stored in first row column U. A string, may contain non alpha characters including commas.
            counter_name = sheet.cell(count_row,20).value
            counter_name = counter_name.replace(',', '')

            # Specify that you are counting bicycles, other counts condcuted by Council record a mix of bicycles and pedestrians.
            counting = "bicycle riders"

            # Sixth line of results file, a header for count specific information
            output_file.write('countsite, count_date, bin_duration, counting, \
gender_split, counter' + '\n')

            # Seventh line of results file, the count specific information
            output_file.write(
                countsite + "," + str(formatted_date) + "," \
                + str(bin_duration) + "," + counting + ","\
                + gender_split + "," + counter_name + '\n'
                )
            output_file.write('\n')

            countdic = {}

            # Collect movement bin and turn details from this year's count

            # A full data block has bin_duration = "15", gender_split = "Y"
            # There are no counts that have have a gender count without a 15min breakdown
            # However, a few 15min counts do not have gender breakdowns.

            # Desired Data output order female cyclists, 7:00 to 9:00 a line break, new header row then male cyclists, 7:00 to 9:00
            if gender_split == 'Y':
                genders = ('F','M')
            else:
                genders = ('NA',)

            for gender in genders:

                genderdic = {}

                output_file.write('\n')
                output_file.write('countsite, time, bin_duration, counting, \
gender, north_turn_left, north_through, north_turn_right, east_turn_left, \
east_through, east_turn_right, south_turn_left, south_through, \
south_turn_right, west_turn_left, west_through, west_turn_right' + '\n')


                # Collect each of the 15 minute observations
                for obs_row in range(5,13):
                    # Data in a full block has movement observations recorded in the sixth to thirteen rows

                    # The first column contains the start time in excel time format.
                    excel_start_time = sheet.cell(count_row + obs_row,0).value

                    # Convert start time to YYYY-MM-DD HH:MM:SS format  TODO: Can we change the formating to lose the seconds?
                    preformatted_start_time	= xldate_as_tuple(excel_start_time,workbook.datemode)
                    formatted_time = time(*preformatted_start_time[3:5])
                    start_datetime = datetime.combine(formatted_date,formatted_time)

                    # Collect observation data (how many people made what turn)
                    # The movements of male cyclists (or NA gender) are recorded in columns C, E, G, I, K, M, O, Q, S, U, W, Y


                    male_movements = [
                                    ('north_turn_right', 2), ('north_through', 4), ('north_turn_left',6),\
                                    ('east_turn_right',8),('east_through',10),('east_turn_left',12),\
                                    ('south_turn_right',14),('south_through',16),('south_turn_left',18),\
                                    ('west_turn_right',20),('west_through',22),('west_turn_left',24)\
                                    ]

                    turndic = {}
                    for (turn, turn_col) in male_movements:
                        if gender == 'F':
                            turn_col = turn_col+1
                            # The movments of female cyclists are recorded in next column (i,e columns D, F, H, J, L, N, P, R, T, V, X, Z)

                        try:
                            turnscrape = int(sheet.cell(count_row + obs_row,turn_col).value)
                        except:
                            turnscrape = "NA"

                        turndic[turn] = turnscrape

                    output_file.write(
                        countsite + "," + str(start_datetime) + "," \
                        + str(bin_duration) + "," + counting + "," + gender + "," \
                        + str(turndic['north_turn_left']) + "," + str(turndic['north_through']) + "," + str(turndic['north_turn_right']) + "," \
                        + str(turndic['east_turn_left']) + "," + str(turndic['east_through']) + "," + str(turndic['east_turn_right']) + "," \
                        + str(turndic['south_turn_left']) + "," + str(turndic['south_through']) + "," + str(turndic['south_turn_right']) + "," \
                        + str(turndic['west_turn_left']) + "," + str(turndic['west_through']) + ","   + str(turndic['west_turn_right']) + '\n' \
                        )


                    # ------------------------------------------------------------------------------
                    # Step 2
                    # Sum observations to develop useful information
                    # Also, scrape excel spreadsheet for old super tuesday ( bin_duration = 120 counts ) data.

                    # Create a .csv file reporting each metric suitable for plotting change over time (and calcuating annualised growth for each site.
                    # ------------------------------------------------------------------------------

                    # Store all the count observations made on a specified count date for calculations
                    genderdic[start_datetime] = turndic
                countdic[gender] = genderdic

            countsummary = {}
            countsummary['countsite'] = countsite
            countsummary['dist_from_cbd'] = dist_from_cbd
            countsummary['time'] = min(countdic[gender])
            countsummary['bin_duration'] = 120 # Hard coded, it would be better if it were summed from consituent rows.
            countsummary['counting'] = 'bicycle riders'
            countsummary['gender'] = 'NA'

            all_moves_dic = ['north_turn_right', 'north_through', 'north_turn_left', \
                            'east_turn_right', 'east_through', 'east_turn_left',\
                            'south_turn_right','south_through','south_turn_left',\
                            'west_turn_right','west_through','west_turn_left']
            all_moves = sum_observations(all_moves_dic, countdic)
            countsummary['total'] = all_moves
#            print countsite, all_moves, 'All bike movements'

            from_north_dic  = ['north_turn_left', 'north_through','north_turn_right']
            from_north = sum_observations(from_north_dic, countdic);
            countsummary['from_north'] = from_north
#            print countsite, from_north, "From North"

            from_east_dic   = ['east_turn_left', 'east_through', 'east_turn_right']
            from_east = sum_observations(from_east_dic, countdic)
            countsummary['from_east'] = from_east
#            print countsite, from_east, "From East"

            from_south_dic  = ['south_turn_left', 'south_through', 'south_turn_right']
            from_south = sum_observations(from_south_dic, countdic)
            countsummary['from_south'] = from_south
#            print countsite, from_south, "From South"

            from_west_dic   = ['west_turn_left', 'west_through', 'west_turn_right']
            from_west = sum_observations(from_west_dic, countdic)
            countsummary['from_west'] = from_west
#            print countsite, from_west, "From West"

            to_north_dic    = ['east_turn_right', 'south_through', 'west_turn_left']
            to_north = sum_observations(to_north_dic, countdic)
            countsummary['to_north'] = to_north
#            print countsite, to_north, "To North"

            to_east_dic     = ['north_turn_left', 'west_through', 'south_turn_right']
            to_east  = sum_observations(to_east_dic, countdic)
            countsummary['to_east'] = to_east
#            print countsite, to_east, "To East"

            to_south_dic    = ['north_through', 'east_turn_left', 'west_turn_right']
            to_south  = sum_observations(to_south_dic, countdic)
            countsummary['to_south'] = to_east
#            print countsite, to_south, "To South"

            to_west_dic     = ['north_turn_right', 'east_through', 'south_turn_left']
            to_west   = sum_observations(to_west_dic, countdic)
            countsummary['to_west'] = to_west
#            print countsite, to_west, "To West"

            sitedic[str_formatted_date] = countsummary
        # ------------------------------------------------------------------------------
        # Old Super Tuesday counts

        else:
            #   Old Super Tuesday counts contain a value in 'Count Year' but nothing in 'Count Date'

            try:
                countyear_test = int(sheet.cell(count_row,13).value)
            except:
                countyear_test = 0

            #   The earlist super tuesday count in this dataset was conducted in 2006.
            #   If the Count Date field contains a date value less than 2005, the data block is blank

            if countyear_test > 2000:
                print 'Countyear_test =', countyear_test, 'Historic Super Tue Data'
                countsummary = {}
                countsummary['countsite'] = countsite
                countsummary['dist_from_cbd'] = dist_from_cbd
                countsummary['bin_duration'] = 120
                countsummary['counting'] = 'bicycle riders'
                countsummary['gender'] = 'NA'

                #   Since no count date is specified we will need to add this data.
                #   Count assumed to occur on First tuesday of March.
                #   First Tuesday: 1 March 2005; 7 March 2006; 6 March 2007; 4 March 2008; 3 March 2009;
                first_tue = {
                    2005 : date(2005, 3, 1), 2006 : date(2006, 3, 7), \
                    2007 : date(2007, 3, 6), 2008 : date(2008, 3, 4), \
                    2009 : date(2009, 3, 3)
                            }

                countsummary['time'] = datetime.combine(first_tue[countyear_test], time(07,00,00))

                #   In a historic super tuesday count results are recorded in the 28th row of the data block:
                #   7-9am all bicycle movements                         column C
                all_moves = int(sheet.cell(count_row+27,2).value)
                countsummary['total'] = all_moves
#                print "Total all movements 7-9am", formatted_date, "=", all_moves

                #   7-9am all bicycle movements entering from North,    column G
                try:
                    from_north = int(sheet.cell(count_row+27,6).value)
                except:
                    from_north = "NA"
                countsummary['from_north'] = from_north
#                print "From North 7-9am", formatted_date, "=", from_north

                #   7-9am all bicycle movements entering from East,     column H
                try:
                    from_east = int(sheet.cell(count_row+27,7).value)
                except:
                    from_east = "NA"
                countsummary['from_east'] = from_east
#                print "From East 7-9am", formatted_date, "=", from_east

                #   7-9am all bicycle movements entering from South,    column I
                try:
                    from_south = int(sheet.cell(count_row+27,8).value)
                except:
                    from_south = "NA"
                countsummary['from_south'] = from_south
#                print "From South 7-9am", formatted_date, "=", from_south

                #   7-9am all bicycle movements entering from West,     column J
                try:
                    from_west = int(sheet.cell(count_row+27,9).value)
                except:
                    from_west = "NA"
                countsummary['from_west'] = from_west
#                print "From West 7-9am", formatted_date, "=", from_west

                #   7-9am all bicycle movements departing via North,    column K
                try:
                    to_north = int(sheet.cell(count_row+27,10).value)
                except:
                    to_north = "NA"
                countsummary['to_north'] = to_north
#                print "To North 7-9am", formatted_date, "=", to_north

                #   7-9am all bicycle movements departing via East,     column L
                try:
                    to_east = int(sheet.cell(count_row+27,11).value)
                except:
                    to_east = "NA"
                countsummary['to_east'] = to_east
#                print "To East 7-9am", formatted_date, "=", to_east

                #   7-9am all bicycle movements departing via South,    column M
                try:
                    to_south = int(sheet.cell(count_row+27,12).value)
                except:
                    to_south = "NA"
                countsummary['to_south'] = to_south
#                print "To South 7-9am", formatted_date, "=", to_south

                #   7-9am all bicycle movements departing via West,     column N
                try:
                    to_west = int(sheet.cell(count_row+27,13).value)
                except:
                    to_west = "NA"
                countsummary['to_west'] = to_west
#                print "To West 7-9am", formatted_date, "=", to_west

                sitedic[str_formatted_date] = countsummary

        print sitedic
        summary_list = ['countsite', 'dist_from_cbd', 'time', 'bin_duration', 'counting', 'gender', 'total', \
                        'from_north', 'from_east', 'from_south', 'from_west', \
                        'to_north', 'to_east', 'to_south', 'to_west']

        output_file = open(resultsdir + countsite + "/" + countsite + "_summary7am-9am.csv", "w")
        for field in summary_list:
            output_file.write(field + ', ')
        output_file.write('\n')

        for countdate in sitedic:
            for field in summary_list:
                output_file.write(str(sitedic[countdate][field]) + ', ')
            output_file.write('\n')

        # TODO: Output Summarised count to a file one file for each count site
        # [countdate 07:00:00][bin_duration = 120][gender = NA][total][from north][from east]\
        # [from south][from west][to north][to east][to south][to west]






                # ------------------------------------------------------------------------------


# Step 3
#
# For each of the metrics in step 2, calculate annualised growth, to report on change over time
