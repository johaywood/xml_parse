import xml.etree.ElementTree as ET
from openpyxl import Workbook
import sys

class AutoVivification(dict):
    """Implementation of perl's autovivification feature"""
    def __getitem__(self, item):
        try:
            return dict.__getitem__(self, item)
        except KeyError:
            value = self[item] = type(self)()
            return value

#Generate tree in memory from XML file
tree = ET.parse(sys.argv[1])
grp_tree = ET.parse(sys.argv[2])
root = tree.getroot()
grp_root = grp_tree.getroot()

#Initialize lists and dicts
animal_uars = []
meas_ids = []
grp_ids = []
grp_summary_ids = []
group_factor_results = []
group_grp_ids = []
meas_by_day = []
animal_key = AutoVivification()
an_grp_key = AutoVivification()
grp_key = AutoVivification()
group_grp_key = AutoVivification()
meas_key = AutoVivification()
grp_summary_key = AutoVivification()
results = AutoVivification()
grp_results = AutoVivification()

study_number = str()
study_title = str()

#Collect study info
for sn in root.iter('STUDY'):
  study_number = str(sn.find('STUDY_REFERENCE').text)
  study_title = str(sn.find('STUDY_TITLE').text)

#Collect animal ID info
for anim in root.iter('ANIMAL'):
  anim_uar = str(anim.find('ANIMAL_UAR').text)
  anim_num = str(anim.find('ANIMAL_REFERENCE').text)
  anim_grp = str(anim.find('GROUP_ID').text)
  animal_uars.append(anim_uar)
  animal_key[anim_uar] = anim_num
  an_grp_key[anim_uar] = anim_grp

#Collect measurement ID info
for meas in root.iter('MEASUREMENT'):
  meas_id = str(meas.find('MEASUREMENT_ID').text)
  meas_name = str(meas.find('MEASUREMENT_DESCR').text)
  meas_ids.append(meas_id)
  meas_key[meas_id] = meas_name
  
#Collect individual file group ID info
for grp in root.iter('GROUP'):
  grp_id = str(grp.find('GROUP_ID').text)
  grp_name = str(grp.find('GROUP_LONG_NAME').text)
  grp_ids.append(grp_id)
  grp_key[grp_id] = grp_name
  
#Collect group file group ID info
for grp in grp_root.iter('GROUP'):
  grp_id = str(grp.find('GROUP_ID').text)
  grp_name = str(grp.find('GROUP_LONG_NAME').text)
  group_grp_ids.append(grp_id)
  group_grp_key[grp_id] = grp_name
  
#Collect group summary ID info
for summ in grp_root.iter('GROUP_SUMMARY'):
  summ_id = str(summ.find('GROUP_SUMMARY_ID').text)
  summ_name = str(summ.find('GROUP_SUMMARY_DESCR').text)
  grp_summary_ids.append(summ_id)
  grp_summary_key[summ_id] = summ_name

#Collect group factor results (timepoints)
for gft in grp_root.iter('GROUP_FACTOR_RESULT'):
  for tp_from in gft.iter('TIME_PERIOD_FROM'):
    tpf = tp_from.text
  for tp_to in gft.iter('TIME_PERIOD_TO'):
    tpt = tp_to.text
  if tpf == tpt:
    tp = 'Day ' + tpf
  else:
    tp = 'Day ' + tpf + ' - Day ' + tpt
  group_factor_results.append(tp)
  
#Collect animal results by measurement and UAR
for result in root.iter('ANIMAL_RESULT'):
  for uar_result in result.iter('ANIMAL_UAR'):
    an = uar_result.text
  for mid_result in result.iter('MEASUREMENT_ID'):
    mid = mid_result.text
  for str_result in result.iter('RESULT_STRING'):
    val = str_result.text
    results[an][mid]['RESULT_STRING'] = val
    
#Collect group results by measurement, timepoint, group ID
for result in grp_root.iter('GROUP_SUMMARY_RESULT'):
  for grp_id_result in result.iter('GROUP_ID'):
    gn = grp_id_result.text
    group_grp_ids.append(gn)
  for mid_result in result.iter('MEASUREMENT_ID'):
    mid = mid_result.text
  for gsid_result in result.iter('GROUP_SUMMARY_ID'):
    gsid = gsid_result.text
  for tp_from in result.iter('TIME_PERIOD_FROM'):
    tpf = tp_from.text
  for tp_to in result.iter('TIME_PERIOD_TO'):
    tpt = tp_to.text
  if tpf == tpt:
    tp = meas_key[mid] + ' Day ' + tpf
  else:
    tp = meas_key[mid] + ' Day ' + tpf + ' - Day ' + tpt
  meas_by_day.append(tp)  
  for grp_str_result in result.iter('GROUP_RESULT_STRING'):
    val = grp_str_result.text
    grp_results[tp][gn][gsid]['RESULT_STRING'] = val
meas_by_day = set(meas_by_day) 
meas_by_day = list(meas_by_day)
group_grp_ids = set(group_grp_ids)
group_grp_ids = list(group_grp_ids)
group_grp_ids.sort()

#Initialize openpyxl workbook in memory
wb = Workbook()
ws = wb.active
ws.title = study_number + " Ind. Animal Data"
ws2 = wb.create_sheet()
ws2.title = study_number + " Group Summary"

#Add animal # and measurement column headers from lookup of measurement descriptions
ws.cell(row=1, column=1).value = "Animal #"
ws.cell(row=1, column=2).value = "Group"
for j in range (1, len(meas_ids) + 1):
  ws.cell(row=1, column=j+2).value = meas_key[meas_ids[j-1]]

#Insert animal number and group name via lookup of UAR
for i in range(1, len(animal_uars) + 1): 
  ws.cell(row=i+1, column=1).value = animal_key[animal_uars[i-1]]
  ws.cell(row=i+1, column=2).value = grp_key[an_grp_key[animal_uars[i-1]]]
  #If a measurement exists for the animal and is not an empty dictionary
  #lookup the result string by UAR and measurement ID and insert into spreadsheet
  for j in range (1, len(meas_ids) + 1):
    if results[animal_uars[i-1]][meas_ids[j-1]] != None and results[animal_uars[i-1]][meas_ids[j-1]] != {}:
      ws.cell(row=i+1, column=j+2).value = results[animal_uars[i-1]][meas_ids[j-1]]['RESULT_STRING']

#Insert row/column headers and format the group summary table layout
for x in range (0, len(meas_by_day)):
  ws2.merge_cells(start_row=(((x+1)*7+x)-6),start_column=2,end_row=(((x+1)*7+x)-6),end_column=len(grp_ids)+1)
  ws2.cell(row=(((x+1)*7+x)-6), column=2).value = meas_by_day[x]
  for j in range (0, len(grp_ids)):
    ws2.cell(row=(((x+1)*7+x)-5), column=j+2).value = group_grp_key[group_grp_ids[j]]
  for i in range (0, len(grp_summary_ids)):
    ws2.cell(row=(((x+1)*7+x)-4)+i, column=1).value = grp_summary_key[grp_summary_ids[i]]

#Put data into the group layout
for x in range (0, len(meas_by_day)):
  for y in range (0, len(group_grp_ids)):
    for z in range (0, len(grp_summary_ids)):
      if grp_results[meas_by_day[x]][group_grp_ids[y]][grp_summary_ids[z]]['RESULT_STRING'] == {}:
        ws2.cell(row=(((x+1)*7+x+z)-4), column=y+2).value = ''
      else:
        ws2.cell(row=(((x+1)*7+x+z)-4), column=y+2).value = grp_results[meas_by_day[x]][group_grp_ids[y]][grp_summary_ids[z]]['RESULT_STRING']


save_name = (sys.argv[1].split("."))[0]
      
wb.save('%s.xlsx' % save_name)