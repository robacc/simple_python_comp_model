from simple_python_comp_model import CompUtils

__author__ = 'sujeet'

import openpyxl
from openpyxl_helpers import *
import numpy
from openpyxl.style import NumberFormat

# Open workbook
wb = openpyxl.load_workbook('/home/sujeet/Desktop/plan_modeling.xlsx')

# Read plan parameters
payout_table_range = read_table_from_named_range(wb,'PayoutTable')
infinity_rate = read_value_from_named_range(wb,'InfinityRate')
TIC = read_value_from_named_range(wb,'TIC')

# Read simulation parameters
average_achievement = read_value_from_named_range(wb,'AverageAchievement')
std_achievement = read_value_from_named_range(wb,'StdDevAchievement')
sample_size = read_value_from_named_range(wb,'SampleSize')

# Initialize Rate Table
payout_table = CompUtils.Rate_Table()
payout_table.add_row_list(payout_table_range,startKey='achievement',baseKey='attainment')
payout_table.set_infinity_rate(infinity_rate)

# Populate the achievement sample
achievement_distribution = numpy.random.normal(average_achievement,std_achievement,sample_size)

# Calculate Attainment based on the Payout Table
attainment = CompUtils.calculate_lookup_attainment(achievement_distribution,payout_table.rows)

# Calculate Earnings as Attainment times Target Incentive
earnings = attainment * TIC

# Create an output sheet for us to write to
output_ws = wb.create_sheet(title='Output')

# Write Achievement, Attainment and Earnings into columns
write_list_of_values(output_ws,'A1',list(achievement_distribution),'Achievement',number_format=NumberFormat.FORMAT_PERCENTAGE_00)
write_list_of_values(output_ws,'B1',list(attainment),'Attainment',number_format=NumberFormat.FORMAT_PERCENTAGE_00)
write_list_of_values(output_ws,'C1',list(earnings),'Earnings',number_format=NumberFormat.FORMAT_CURRENCY_USD_SIMPLE)

# Create a Distribution Table
create_distribution_table(output_ws,'E1',earnings,number_format=NumberFormat.FORMAT_CURRENCY_USD_SIMPLE)

# Create a Stats Table
#create_stats_table(output_ws,'I1',earnings,number_format=NumberFormat.FORMAT_CURRENCY_USD_SIMPLE)

# Write the output to a new file
wb.save(filename='/home/sujeet/Desktop/plan_modeling_simulated.xlsx')