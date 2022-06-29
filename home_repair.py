from cmath import sqrt
import pandas as pd
import numpy as np 

#Read in second file with square footage and num of stories tied to survey ID
file_name = "EXCEL_Survey 1 Returns Dataset_ALL.xlsx"
sheet = "S1_583_RETURNS"
answerkey = pd.read_excel(f"data/{file_name}", dtype=str, index_col="Survey_ID", sheet_name=sheet)
costs = pd.read_csv("data/costs.csv", index_col=0, squeeze=True)

#df=pd.DataFrame(answerkey, columns=['Survey_ID','building_square_footage', 'building_stories', 'Combined', 'EstCost'])
df = answerkey
df["building_square_footage"] = pd.to_numeric(df["building_square_footage"])

#Convert dataframes to a single dimensional array
arr = df.to_numpy()

#Convert arrays to type string to process by character later in code
stringarr2 = arr.astype(str)

count = 0
ec = 0

#introduce cost figures
averagesize = 1390
TotalCost = 0

serviceheatcount = 0
replaceheatcount = 0
weatherizationcount = 0
upgradeelectriccount = 0
replacebreakercount = 0
electricplugcount = 0
replacepipingcount = 0
snakeaugercount = 0
exterminatorcount = 0
repaintingcount=0
repairinteriorwallcount = 0
repairfloorcount = 0
serviceacequipmentcount = 0

def check_serviceheatingequipment(x):
    if x["Q1"] == "1": return True
    if x["Q2"] == "1" and x["Q3"] == "3": return True
    if x["Q2"] == "1" and x["Q3"] == "4": return True
    return False


def check_replaceheatingeuipment(x):
    if x["Q1"] == "2": return True
    if x["Q2"] == "2" and x["Q3"] == "3": return True
    if x["Q2"] == "2" and x["Q3"] == "4": return True
    return False


def check_weatherization1(x):
    if x["Q3"] == "1": return True
    if x["Q3"] == "2": return True
    return False


def check_serviceacequipment(x):
    if x["Q4"] == "1": return True
    if x["Q5"] == "1" and x["Q3"] == "3": return True
    if x["Q5"] == "1" and x["Q3"] == "4": return True
    return False


def check_replaceacequipment(x):
    if x["Q4"] == "2": return True
    if x["Q4"] == "4": return True
    if x["Q5"] == "2" and x["Q3"] == "3": return True
    if x["Q5"] == "2" and x["Q3"] == "4": return True
    return False


def check_wireforelectric(x):
    if x["Q6"] == "1": return True
    if x["Q6"] == "2": return True
    return False


def check_installelectricplugs(x):
    if x["Q7"] == '1': return True
    if x["Q7"] == '2': return True
    if x["Q10"] == '1': return True
    if x["Q10"] == '2': return True
    return False


def check_concealwiring(x):
    if x["Q8"] == '1': return True # multiplier
    if x["Q8"] == '2': return True
    return False


def check_minorelectrical(x):
    if x["Q9"] == '1': return True # ec = ec + (costs.minorelectrical * 1.5)
    if x["Q9"] == '2': return True # ec = ec + (costs.minorelectrical * 3)
    return False


def check_upgradeelectricservice(x):
    if x["Q11"] == '1' and x["Q12"] == '1': return True
    if x["Q11"] == '2' and x["Q12"] == '1': return True
    return False


def check_replacebreaker(x):
    if x["Q11"] == '1' and x["Q12"] == '2': return True
    if x["Q11"] == '2' and x["Q12"] == '2': return True
    return False


def check_moldremediation(x):
    if x["Q13"] == '1': return True
    if x["Q13"] == '2': return True
    return False


def check_repairtoilet(x):
    if x["Q14"] == '1': return True
    if x["Q14"] == '2': return True
    return False


def check_replacepiping(x):
    if x["Q15"] == '1': return True
    if x["Q15"] == '2': return True
    if x["Q18"] == '2': return True
    if x["Q19"] == '2': return True
    return False


def check_replacewaterheater(x):
    if x["Q16"] == '1': return True
    if x["Q16"] == '2': return True
    return False


def check_snakeaugerline(x):
    if x["Q17"] == '1': return True
    return False


def check_repairsewerline(x):
    if x["Q17"] == '2': return True
    return False


def check_snakeaugerdrane(x):
    if x["Q18"] == '1': return True
    if x["Q19"] == '1': return True
    return False


def check_exterminatetermite(x):
    if x["Q20"] == '1': return True
    return False


if x["Q20"] == '2':
    ec = ec + costs.exterminatetermite
    exterminatorcount = exterminatorcount + 1
if x["Q21"] == '1':
    ec = ec + costs.exterminaterodent
if x["Q21"] == '2':
    ec = ec + costs.exterminateseal

# Question Block 4
if x["Q22"] == '1':
    ec = ec + costs.sealbasement
if x["Q22"] == '2':
    ec = ec + costs.sealbasement
if x["Q23"] == '1':
    ec = ec + costs.sealwindow
if x["Q23"] == '2':
    ec = ec + costs.sealwindow
if x["Q24"] == '1':
    roofsquarefootage = .05*(1.25*squarefootage)
    ec = ec + (roofsquarefootage * costs.sealroofmultiplier)
if x["Q24"] == '2':
    roofsquarefootage = .05*(1.25*squarefootage)
    ec = ec + (roofsquarefootage * costs.sealroofmultiplier)
if x["Q25"] == '1' and replacepipingcount == 0:
    ec = ec +costs.replacepiping
    replacepipingcount = replacepipingcount + 1
if x["Q25"] == '2' and replacepipingcount == 0:
    ec = ec + costs.replacepiping
    replacepipingcount = replacepipingcount + 1
if x["Q26"] == '1':
    ec = ec + costs.repairwallsreplacepiping
if x["Q26"] == '2':
    ec = ec + costs.repairwallsreplacepiping
if x["Q27"] == '1':
    ec = ec + costs.replacewaterheater
if x["Q27"] == '2':
    ec = ec + costs.replacewaterheater
if x["Q28"] == '1' and replacepipingcount == 0:
    ec = ec + costs.replacepiping
    replacepipingcount = replacepipingcount + 1
if x["Q28"] == '2' and replacepipingcount == 0:
    ec = ec + costs.replacepiping
    replacepipingcount = replacepipingcount + 1
if x["Q29"] == '1' and snakeaugercount == 0:
    ec = ec + costs.snakeaugerdrane
    snakeaugercount = snakeaugercount + 1
if x["Q29"] == '2' and snakeaugercount == 0:
    ec = ec + costs.snakeaugerdrane
    snakeaugercount = snakeaugercount + 1

# Question Block 5
if x["Q30"] == '1':
    basementwallarea = (sqrt(squarefootage/numfloors)) * 4
    ec = ec + basementwallarea * costs.repairfoundationmultiplier
if x["Q30"] == '2':
    basementwallarea = (sqrt(squarefootage/numfloors)) * 4
    ec = ec + basementwallarea * costs.repairfoundationmultiplier
if x["Q31"] == '1':
    ec = ec + costs.repairroof
if x["Q31"] == '2':
    ec = ec + costs.repairroof
if x["Q32"] == '1':
    roofarea = 1.25*(squarefootage/numfloors)
    ec = ec + (roofarea * costs.replaceroofmaterialsmultiplier)
if x["Q32"] == '2':
    roofarea = 1.25*(squarefootage/numfloors)
    ec = ec + (roofarea * costs.replaceroofmaterialsmultiplier)
if x["Q33"] == '1':
    roofarea = 1.25*(squarefootage/numfloors)
    ec = ec + (roofarea * costs.replaceroofmultiplier)
if x["Q33"] == '2':
    roofarea = 1.25*(squarefootage/numfloors)
    ec = ec + (roofarea * costs.replaceroofmultiplier)
if x["Q34"] == '1':
    ec = ec + costs.replacegutters
if x["Q34"] == '2':
    ec = ec + costs.replacegutters
if x["Q35"] == '1':
    ec = ec + costs.tuckpointing
if x["Q35"] == '2':
    ec = ec + costs.tuckpointing
if x["Q36"] == '1':
    ec = ec + costs.repainting
    repaintingcount = repaintingcount+1
if x["Q36"] == '2':
    ec = ec + costs.repainting
    repaintingcount = repaintingcount+1
if x["Q37"] == '1':
    ec = ec + (squarefootage * costs.replaceexteriorwallmultiplier)
if x["Q37"] == '2':
    ec = ec + (squarefootage * costs.replaceexteriorwallmultiplier)
if x["Q38"] == '1':
    ec = ec + costs.replacewindow
if x["Q38"] == '2':
    ec = ec + costs.replacewindow

# Question Block 6
if x["Q39"] == '1':
    ec = ec + costs.repairfloor
    repairfloorcount = repairfloorcount + 1
if x["Q39"] == '2':
    ec = ec + costs.repairfloor
    repairfloorcount = repairfloorcount + 1
if x["Q40"] == '1':
    ec = ec + costs.repairinteriorwall
    repairinteriorwallcount = repairinteriorwallcount + 1
if x["Q40"] == '2':
    ec = ec + costs.repairinteriorwall
    repairinteriorwallcount = repairinteriorwallcount + 1
if x["Q41"] == '1' and repaintingcount == 0:
    ec = ec + costs.repainting
    repaintingcount = repaintingcount+1
if x["Q41"] == '2' and repaintingcount == 0:
    ec = ec + costs.repainting
    repaintingcount=repaintingcount+1
if x["Q42"] == '1' and repairinteriorwallcount == 0:
    ec = ec + costs.repairinteriorwall
    repairinteriorwallcount = repairinteriorwallcount + 1
if x["Q42"] == '2' and repairinteriorwallcount == 0:
    ec = ec + costs.repairinteriorwall
    repairinteriorwallcount = repairinteriorwallcount + 1
if x["Q43"] == '1':
    ec = ec + costs.repaintingwindowsashes
if x["Q43"] == '2':
    ec = ec + costs.repaintingwindowsashes
if x["Q44"] == '1' and repairfloorcount == 0:
    ec = ec + costs.repairfloor
    repairfloorcount = repairfloorcount + 1
if x["Q44"] == '2' and repairfloorcount == 0:
    ec = ec + costs.repairfloor
    repairfloorcount = repairfloorcount + 1
if x["Q45"] == '1':
    ec = ec + costs.repairfloortoiletsink
if x["Q45"] == '2':
    ec = ec + costs.repairfloortoiletsink
if x["Q46"] == '1':
    ec = ec + costs.replaceinstalllocks
if x["Q46"] == '2':
    ec = ec + costs.replaceinstalllocks

# Question Block 7
if x["Q47"] == '1':
    ec = ec + costs.treeremovalpruning
if x["Q47"] == '2':
    ec = ec + costs.treeremovalpruning
if x["Q48"] == '1':
    ec = ec + costs.repairexternalwalkway
if x["Q48"] == '2':
    ec = ec + costs.repairexternalwalkway
if x["Q48"] == '1':
    ec = ec + costs.repairpatio
if x["Q48"] == '2':
    ec = ec + costs.repairpatio
if x["Q50"] == '1':
    ec = ec + costs.repairporchdeck
if x["Q50"] == '2':
    ec = ec + costs.repairporchdeck
if x["Q51"] == '1':
    ec = ec + costs.repairexternalstairway
if x["Q51"] == '2':
    ec = ec + costs.repairexternalstairway
if x["Q52"] == '1':
    ec = ec + costs.repairadditionalstructure
if x["Q52"] == '2':
    ec = ec + costs.repairadditionalstructure

    serviceheatcount = 0
    replaceheatcount = 0
    weatherizationcount = 0
    upgradeelectriccount = 0
    replacebreakercount = 0
    electricplugcount = 0
    replacepipingcount = 0
    snakeaugercount = 0
    exterminatorcount = 0
    repaintingcount=0
    repairinteriorwallcount = 0
    repairfloorcount = 0

    #Leave out the imaginary number portion of the estimated cost with.real
    if np.isnan(ec):
        ec = 0
    x[4] = float(ec.real)
    TotalCost = TotalCost + ec
    ec = 0

outputdf = pd.DataFrame(stringarr2)

## Set Columns
df["serviceheatingequipment"] = df.apply(check_serviceheatingequipment, axis=1)
df["replaceheatingequipment"] = df.apply(check_replaceheatingeuipment, axis=1)
df["weatherization1"] = df.apply(check_weatherization1, axis=1)

xlsxfile = 'CEOutput.xls'
excel_writer = pd.ExcelWriter(xlsxfile, engine='xlsxwriter')
outputdf.columns=['SurveyID','SquareFootage','NumStories','SurveyResponses','Est. Cost']
outputdf.to_excel(excel_writer, sheet_name="SurveyIDandCE",startrow=0, startcol=0, header = True, index=True)
excel_writer.save()