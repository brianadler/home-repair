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
    if x["Q25"] == '1': return True
    if x["Q25"] == '2': return True
    if x["Q28"] == '1': return True
    if x["Q28"] == '2': return True
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
    if x["Q29"] == '1': return True    
    if x["Q29"] == '2': return True
    return False


def check_exterminatetermite(x):
    if x["Q20"] == '1': return True
    if x["Q20"] == '2': return True
    if x["Q21"] == '2': return True
    return False


def check_exterminateseal(x):
    if x["Q21"] == '2': return True
    return False

def check_sealbasement(x):
    if x["Q22"] == '1': return True
    if x["Q22"] == '2': return True
    return False

def check_sealwindow(x):
    if x["Q23"] == '1': return True
    if x["Q23"] == '2': return True
    return False

#Must add multiplier: roofsquarefootage = .05*(1.25*squarefootage)
#ec = ec + (roofsquarefootage * costs.sealroofmultiplier

def check_sealroof(x):
    if x["Q24"] == '1': return True
    if x["Q24"] == '2': return True
    return False

def check_repairwallsreplacepiping(x):
    if x["Q26"] == '1': return True
    if x["Q26"] == '2': return True
    return False

def check_replacewaterheater(x):
    if x["Q27"] == '1': return True
    if x["Q27"] == '2': return True
    return False

#multiplier: basementwallarea = (sqrt(squarefootage/numfloors)) * 4
    #ec = ec + basementwallarea * costs.repairfoundationmultiplier
def check_repairfoundation(x):
    if x["Q30"] == '1': return True
    if x["Q30"] == '2': return True
    return False


def check_repairroof(x):
    if x["Q31"] == '1': return True
    if x["Q31"] == '2': return True
    return False

#Replace Roof Multiplier roofarea = 1.25*(squarefootage/numfloors)
   # ec = ec + (roofarea * costs.replaceroofmaterialsmultiplier)
def check_replaceroof(x):
    if x["Q32"] == '1': return True
    if x["Q32"] == '2': return True
    if x["Q33"] == '1': return True
    if x["Q33"] == '2': return True
    return False

def check_replacegutters(x):
    if x["Q34"] == '1': return True
    if x["Q34"] == '2': return True
    return False

def check_tuckpointing(x):
    if x["Q35"] == '1': return True
    if x["Q35"] == '2': return True
    return False

def check_repainting(x):
    if x["Q36"] == '1': return True
    if x["Q36"] == '2': return True
    if x["Q41"] == '1': return True
    if x["Q41"] == '2': return True
    return False

#Multiplier:     ec = ec + (squarefootage * costs.replaceexteriorwallmultiplier)
def check_replaceexteriorwall(x):
    if x["Q37"] == '1': return True
    if x["Q37"] == '2': return True
    return False

def check_replacewindow(x):
    if x["Q38"] == '1': return True
    if x["Q38"] == '2': return True
    return False

def check_repairfloor(x):
    if x["Q39"] == '1': return True
    if x["Q39"] == '2': return True
    return False

def check_repairinteriorwall(x):
    if x["Q40"] == '1': return True
    if x["Q40"] == '2': return True
    if x["Q42"] == '1': return True
    if x["Q42"] == '2': return True
    return False

def check_repaintingwindowsashes(x):
    if x["Q43"] == '1': return True
    if x["Q43"] == '2': return True
    return False

def check_repairfloor(x):
    if x["Q44"] == '1': return True
    if x["Q44"] == '2': return True
    return False

def check_repairfloortoiletsink(x):
    if x["Q45"] == '1': return True
    if x["Q45"] == '2': return True
    return False

def check_replaceinstalllocks(x):
    if x["Q46"] == '1': return True
    if x["Q46"] == '2': return True
    return False

def check_treeremovalpruning(x):
    if x["Q47"] == '1': return True
    if x["Q47"] == '2': return True
    return False

def check_repairexternalwalkway(x):
    if x["Q48"] == '1': return True
    if x["Q48"] == '2': return True
    return False

def check_repairpatio(x):
    if x["Q48"] == '1': return True
    if x["Q48"] == '2': return True
    return False

def check_repairporchdeck(x):
    if x["Q50"] == '1': return True
    if x["Q50"] == '2': return True
    return False

def check_repairexternalstairway(x):
    if x["Q51"] == '1': return True
    if x["Q51"] == '2': return True
    return False

def check_repairadditionalstructure(x):
    if x["Q52"] == '1': return True
    if x["Q52"] == '2': return True
    return False

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