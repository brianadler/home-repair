from math import sqrt
import pandas as pd
import numpy as np 

file_name = "EXCEL_Survey 1 Returns Dataset_ALL.xlsx"
sheet = "S1_583_RETURNS"
answerkey = pd.read_excel(f"data/{file_name}", dtype=str, index_col="Survey_ID", sheet_name=sheet)
costs = pd.read_csv("data/costs.csv", index_col=0, squeeze=True)

df = answerkey
df["building_square_footage"] = pd.to_numeric(df["building_square_footage"])
df["building_stories"] = pd.to_numeric(df["building_stories"])
df["year_built"] = pd.to_numeric(df["year_built"])
df["birthdate"] = pd.to_numeric(df["birthdate"].astype(str).str[0:4])
df["census_tract"] = pd.to_numeric(df["census_tract"])

def check_censustract(x):
    north_city_tracts = [107600, 127000, 108300, 108200, 108100,107300,107200,107400,107500,109600,109700,126700, 110500, 110200, 103000, 111300, 110100, 110300, 126900, 106200,106300,106400,106700,106500,106100]
    north_central_tracts = [105500,105400,105300,105200,105100,105198,106600,112200,112300,111200,111400,111500,127100,110400,120200,126600,125700]
    central_corridor_tracts = [112100,112400,118100,118600,119101,119102,119200,111100,119300,121200,121100,118400,127500,127400,125500,125600]
    south_central_tracts = [104200,104500,117100,117200,127300,123100,123200,127600]
    south_city_tracts = [126800,103600,113500,103400,103700,127200,103800,103100,114102,114101,114200,127200,114300,102200,102100,102300,102500,102400,115100,115200,116200,116100,115300,115400,101300,101200,101100,101500,101800,101400,115500,115600,115700,116301,116302,116400,116500,117400,124200,124100,123300,124300,124600]
    if x["census_tract"] in north_city_tracts: return "North City"
    if x["census_tract"] in north_central_tracts: return "North Central Corridor"
    if x["census_tract"] in central_corridor_tracts: return "Central Corridor" 
    if x["census_tract"] in south_central_tracts: return "South Central Corridor"
    if x["census_tract"] in south_city_tracts: return "South City"

def check_centralair(x):
    if x["Q4"] == '4': return "No Central Air"
    if x["Q5"] == '4': return "No Central Air"
    return "NA"

def check_timeinhome(x):
    if x["Q55"] == '1': return "26+"
    if x["Q55"] == '2': return "11-25"
    if x["Q55"] == '3': return "10-"

def check_surveyedrace(x):
    if x["Q62"] == '1': return "American Indian / Alaskan Native"
    if x["Q62"] == '2': return "Asian / Pacific Islander"
    if x["Q62"] == '3': return "Black / African American"
    if x["Q62"] == '4': return "White / Caucasian"
    if x["Q62"] == '5': return "Other"

def check_surveyedethnicity(x):
    if x["Q61"] == '1': return "Hispanic or Latino"
    if x["Q61"] == '2': return "Not Hispanic or Latino"
    if x["Q61"] == '3': return "Unknown or Prefer not to Say"
    
def check_surveyedincome(x):
    if x["Q59"] == '1': return "17,400 or less"
    if x["Q59"] == '2': return "Between 17,401 and 29,050"
    if x["Q59"] == '3': return "Between 29,051 and 46,450"
    if x["Q59"] == '4': return "Between 46,451 and 69,650"
    if x["Q59"] == '5': return "69,651 or more"    

def check_buildingage(x):
    return 2022 - x["year_built"]

def check_respondentage(x):
    return 2022 - x["birthdate"]

def check_uncomfortablywarmorcold(x):
    if x["Q2"] == "1": return "Uncomfortably Cold"
    if x["Q2"] == "2": return "Uncomfortably Cold"
    if x["Q3"] == "1": return "Uncomfortably Cold"
    if x["Q3"] == "2": return "Uncomfortably Cold"
    return "NA"

def check_serviceheatingequipment(x):
    if x["Q1"] == "1": return costs.serviceheatingequipment
    if x["Q2"] == "1" and x["Q3"] == "3": return costs.serviceheatingequipment
    if x["Q2"] == "1" and x["Q3"] == "4": return costs.serviceheatingequipment
    return 0

def check_replaceheatingeuipment(x):
    if x["Q1"] == "2": return costs.replaceheatingequipment
    if x["Q2"] == "2" and x["Q3"] == "3": return costs.replaceheatingequipment
    if x["Q2"] == "2" and x["Q3"] == "4": return costs.replaceheatingequipment
    return 0

def check_weatherization1(x):
    if x["Q3"] == "1": return costs.weatherization1
    if x["Q3"] == "2": return costs.weatherization1
    return 0

def check_serviceacequipment(x):
    if x["Q4"] == "1": return costs.serviceacequipment
    if x["Q5"] == "1" and x["Q3"] == "3": costs.serviceacequipment
    if x["Q5"] == "1" and x["Q3"] == "4": costs.serviceacequipment
    return 0

def check_replaceacequipment(x):
    if x["Q4"] == "2": return costs.replaceacequipment
    if x["Q4"] == "4": return costs.replaceacequipment
    if x["Q5"] == "2" and x["Q3"] == "3": return costs.replaceacequipment
    if x["Q5"] == "2" and x["Q3"] == "4": return costs.replaceacequipment
    return 0

def check_wireforelectric(x):
    if x["Q6"] == "1": return costs.wireforelectric
    if x["Q6"] == "2": return costs.wireforelectric
    return 0

def check_installelectricplugs(x):
    if x["Q7"] == '1': return costs.installelectricplugs
    if x["Q7"] == '2': return costs.installelectricplugs
    if x["Q10"] == '1': return costs.installelectricplugs
    if x["Q10"] == '2': return costs.installelectricplugs
    return 0

def check_concealwiring(x):
    if x["Q8"] == '1': return 1.5 * costs.concealwiring
    if x["Q8"] == '2': return 3 * costs.concealwiring
    return 0

def check_minorelectrical(x):
    if x["Q9"] == '1': return 1.5*costs.minorelectrical 
    if x["Q9"] == '2': return 1.5*costs.minorelectrical
    return 0

def check_upgradeelectricservice(x):
    if x["Q11"] == '1' and x["Q12"] == '1': return costs.upgradelectricservice
    if x["Q11"] == '2' and x["Q12"] == '1': return costs.upgradelectricservice
    return 0

def check_replacebreaker(x):
    if x["Q11"] == '1' and x["Q12"] == '2': return costs.replacebreaker
    if x["Q11"] == '2' and x["Q12"] == '2': return costs.replacebreaker
    return 0

def check_moldremediation(x):
    if x["Q13"] == '1': return costs.moldremediation
    if x["Q13"] == '2': return costs.moldremediation
    return 0

def check_repairtoilet(x):
    if x["Q14"] == '1': return costs.repairtoilet
    if x["Q14"] == '2': return costs.repairtoilet
    return 0

def check_replacepiping(x):
    if x["Q15"] == '1': return costs.replacepiping
    if x["Q15"] == '2': return costs.replacepiping
    if x["Q18"] == '2': return costs.replacepiping
    if x["Q19"] == '2': return costs.replacepiping
    if x["Q25"] == '1': return costs.replacepiping
    if x["Q25"] == '2': return costs.replacepiping
    if x["Q28"] == '1': return costs.replacepiping
    if x["Q28"] == '2': return costs.replacepiping
    return 0

def check_replacewaterheater(x):
    if x["Q16"] == '1': return costs.replacewaterheater
    if x["Q16"] == '2': return costs.replacewaterheater
    if x["Q27"] == '1': return costs.replacewaterheater
    if x["Q27"] == '2': return costs.replacewaterheater
    return 0

def check_snakeaugerline(x):
    if x["Q17"] == '1': return costs.snakeaugerline
    return 0

def check_repairsewerline(x):
    if x["Q17"] == '2': return costs.repairsewerline
    return 0

def check_snakeaugerdrane(x):
    if x["Q18"] == '1': return costs.snakeaugerdrane
    if x["Q19"] == '1': return costs.snakeaugerdrane
    if x["Q29"] == '1': return costs.snakeaugerdrane    
    if x["Q29"] == '2': return costs.snakeaugerdrane
    return 0

def check_exterminatetermite(x):
    if x["Q20"] == '1': return costs.exterminatetermite
    if x["Q20"] == '2': return costs.exterminatetermite
    if x["Q21"] == '2': return costs.exterminatetermite
    return 0

def check_exterminateseal(x):
    if x["Q21"] == '2': return costs.exterminateseal
    return 0

def check_sealbasement(x):
    if x["Q22"] == '1': return costs.sealbasement
    if x["Q22"] == '2': return costs.sealbasement
    return 0

def check_sealwindow(x):
    if x["Q23"] == '1': return costs.sealwindow
    if x["Q23"] == '2': return costs.sealwindow
    return 0

def check_sealroof(x):
    if x["Q24"] == '1': return x["building_square_footage"] * 1.25 * 0.05 * costs.sealroofmultiplier
    if x["Q24"] == '2': return x["building_square_footage"] * 1.25 * 0.05 * costs.sealroofmultiplier
    return 0

def check_repairwallsreplacepiping(x):
    if x["Q26"] == '1': return costs.repairwallsreplacepiping
    if x["Q26"] == '2': return costs.repairwallsreplacepiping
    return 0

def check_repairfoundation(x):
    if x["Q30"] == '1': return sqrt(x["building_square_footage"]/x["building_stories"])*4*costs.repairfoundationmultiplier
    if x["Q30"] == '2': return sqrt(x["building_square_footage"]/x["building_stories"])*4*costs.repairfoundationmultiplier
    return 0

def check_repairroof(x):
    if x["Q31"] == '1': return costs.repairroof
    if x["Q31"] == '2': return costs.repairroof
    return 0

def check_replaceroof(x):
    if x["Q32"] == '1': return 1.25*(x["building_square_footage"]/x["building_stories"])*costs.replaceroofmaterialsmultiplier
    if x["Q32"] == '2': return 1.25*(x["building_square_footage"]/x["building_stories"])*costs.replaceroofmaterialsmultiplier
    if x["Q33"] == '1': return 1.25*(x["building_square_footage"]/x["building_stories"])*costs.replaceroofmaterialsmultiplier
    if x["Q33"] == '2': return 1.25*(x["building_square_footage"]/x["building_stories"])*costs.replaceroofmaterialsmultiplier
    return 0

def check_replacegutters(x):
    if x["Q34"] == '1': return costs.replacegutters
    if x["Q34"] == '2': return costs.replacegutters
    return 0

def check_tuckpointing(x):
    if x["Q35"] == '1': return costs.tuckpointing
    if x["Q35"] == '2': return costs.tuckpointing
    return 0

def check_repainting(x):
    if x["Q36"] == '1': return costs.repainting
    if x["Q36"] == '2': return costs.repainting
    if x["Q41"] == '1': return costs.repainting
    if x["Q41"] == '2': return costs.repainting
    return 0

def check_replaceexteriorwall(x):
    if x["Q37"] == '1': return x["building_square_footage"]*costs.replaceexteriorwallmultiplier
    if x["Q37"] == '2': return x["building_square_footage"]*costs.replaceexteriorwallmultiplier
    return 0

def check_replacewindow(x):
    if x["Q38"] == '1': return costs.replacewindow
    if x["Q38"] == '2': return costs.replacewindow
    return 0

def check_repairfloor(x):
    if x["Q39"] == '1': return costs.repairfloor
    if x["Q39"] == '2': return costs.repairfloor
    if x["Q44"] == '1': return costs.repairfloor
    if x["Q44"] == '2': return costs.repairfloor
    return 0

def check_repairinteriorwall(x):
    if x["Q40"] == '1': return costs.repairinteriorwall
    if x["Q40"] == '2': return costs.repairinteriorwall
    if x["Q42"] == '1': return costs.repairinteriorwall
    if x["Q42"] == '2': return costs.repairinteriorwall
    return 0

def check_repaintingwindowsashes(x):
    if x["Q43"] == '1': return costs.repaintingwindowsashes
    if x["Q43"] == '2': return costs.repaintingwindowsashes
    return 0

def check_repairfloortoiletsink(x):
    if x["Q45"] == '1': return costs.repairfloortoiletsink
    if x["Q45"] == '2': return costs.repairfloortoiletsink
    return 0

def check_replaceinstalllocks(x):
    if x["Q46"] == '1': return costs.replaceinstalllocks
    if x["Q46"] == '2': return costs.replaceinstalllocks
    return 0

def check_treeremovalpruning(x):
    if x["Q47"] == '1': return costs.treeremovalpruning
    if x["Q47"] == '2': return costs.treeremovalpruning
    return 0

def check_repairexternalwalkway(x):
    if x["Q48"] == '1': return costs.repairexternalwalkway
    if x["Q48"] == '2': return costs.repairexternalwalkway
    return 0

def check_repairpatio(x):
    if x["Q48"] == '1': return costs.repairpatio
    if x["Q48"] == '2': return costs.repairpatio
    return 0

def check_repairporchdeck(x):
    if x["Q50"] == '1': return costs.repairporchdeck
    if x["Q50"] == '2': return costs.repairporchdeck
    return 0

def check_repairexternalstairway(x):
    if x["Q51"] == '1': return costs.repairexternalstairway
    if x["Q51"] == '2': return costs.repairexternalstairway
    return 0

def check_repairadditionalstructure(x):
    if x["Q52"] == '1': return costs.repairadditionalstructure
    if x["Q52"] == '2': return costs.repairadditionalstructure
    return 0

## Set Columns
df["serviceheatingequipment"] = df.apply(check_serviceheatingequipment, axis=1)
df["replaceheatingequipment"] = df.apply(check_replaceheatingeuipment, axis=1)
df["weatherization1"] = df.apply(check_weatherization1, axis=1)
df["serviceacequipment"] = df.apply(check_serviceacequipment, axis=1)
df["replaceacequipment"] = df.apply(check_replaceacequipment, axis=1)
df["wireforelectric"] = df.apply(check_wireforelectric, axis=1)
df["installelectricplugs"] = df.apply(check_installelectricplugs, axis=1)
df["concealwiring"] = df.apply(check_concealwiring, axis=1)
df["minorelectrical"] = df.apply(check_minorelectrical, axis=1)
df["upgradeelectricservice"] = df.apply(check_upgradeelectricservice, axis = 1)
df["replacebreaker"] = df.apply(check_replacebreaker, axis=1)
df["moldremediation"] = df.apply(check_moldremediation, axis=1)
df["repairtoilet"] = df.apply(check_repairtoilet, axis=1)
df["replacepiping"] = df.apply(check_replacepiping, axis=1)
df["replacewaterheater"] = df.apply(check_replacewaterheater, axis=1)
df["snakeaugerline"] = df.apply(check_snakeaugerline, axis=1)
df["repairsewerline"] = df.apply(check_repairsewerline, axis=1)
df["snakeaugerdrane"] = df.apply(check_snakeaugerdrane, axis=1)
df["exterminatetermite"] = df.apply(check_exterminatetermite, axis=1)
df["exterminateseal"] = df.apply(check_exterminateseal, axis=1)
df["sealbasement"] = df.apply(check_sealbasement, axis=1)
df["sealwindow"] = df.apply(check_sealwindow, axis=1)
df["sealroof"] = df.apply(check_sealroof, axis=1)
df["repairwallsreplacepiping"] = df.apply(check_repairwallsreplacepiping, axis=1)
df["repairfoundation"] = df.apply(check_repairfoundation, axis=1)
df["repairroof"] = df.apply(check_repairroof, axis=1)
df["replaceroof"] = df.apply(check_replaceroof, axis=1)
df["replacegutters"] = df.apply(check_replacegutters, axis=1)
df["tuckpointing"] = df.apply(check_tuckpointing, axis=1)
df["repainting"] = df.apply(check_repainting, axis=1)
df["replaceexteriorwall"] = df.apply(check_replaceexteriorwall, axis=1)
df["replacewindow"] = df.apply(check_replacewindow, axis=1)
df["repairfloor"] = df.apply(check_repairfloor, axis=1)
df["repairinteriorwall"] = df.apply(check_repairinteriorwall, axis=1)
df["repaintingwindowsashes"] = df.apply(check_repaintingwindowsashes, axis=1)
df["repairfloortoiletsink"] = df.apply(check_repairfloortoiletsink, axis=1)
df["replaceinstalllocks"] = df.apply(check_replaceinstalllocks, axis=1)
df["treeremovalpruning"] = df.apply(check_treeremovalpruning, axis=1)
df["repairexternalwalkway"] = df.apply(check_repairexternalwalkway, axis=1)
df["repairpatio"] = df.apply(check_repairpatio, axis=1)
df["repairporchdeck"] = df.apply(check_repairporchdeck, axis=1)
df["repairexternalstairway"] = df.apply(check_repairexternalstairway, axis=1)
df["repairadditionalstructure"] = df.apply(check_repairadditionalstructure, axis=1)
df["TotalCost"] = df["serviceheatingequipment"] + df["replaceheatingequipment"] + df["weatherization1"] + df["serviceacequipment"] + df["replaceacequipment"] + df["wireforelectric"] + df["installelectricplugs"] + df["concealwiring"] + df["minorelectrical"] + df["upgradeelectricservice"] + df["replacebreaker"] + df["moldremediation"] + df["repairtoilet"] + df["replacepiping"] + df["replacewaterheater"] +  df["snakeaugerline"] + df["repairsewerline"] + df["snakeaugerdrane"] + df["exterminatetermite"] + df["exterminateseal"] + df["sealbasement"] + df["sealwindow"] + df["sealroof"] + df["repairwallsreplacepiping"] + df["repairfoundation"] + df["repairroof"] + df["replaceroof"] + df["replacegutters"] + df["tuckpointing"] + df["repainting"] + df["replaceexteriorwall"] + df["replacewindow"] + df["repairfloor"] + df["repairinteriorwall"] + df["repaintingwindowsashes"] + df["repairfloortoiletsink"] + df["replaceinstalllocks"] + df["treeremovalpruning"] + df["repairexternalwalkway"] + df["repairpatio"] + df["repairporchdeck"] + df["repairexternalstairway"] + df["repairadditionalstructure"]
df["buildingage"] = df.apply(check_buildingage, axis=1)
df["respondentage"] = df.apply(check_respondentage, axis=1)
df["surveyedrace"] = df.apply(check_surveyedrace, axis = 1)
df["surveyedincome"] = df.apply(check_surveyedincome, axis = 1)
df["surveyedethnicity"] = df.apply(check_surveyedethnicity, axis = 1)
df["timeinhome"] = df.apply(check_timeinhome, axis=1)
df["CityDistrct"] = df.apply(check_censustract, axis=1)
df["CentralAir?"] = df.apply(check_centralair, axis=1)
df["UncomfortablyWarmOrCold"] = df.apply(check_uncomfortablywarmorcold, axis=1)

outputdf = pd.DataFrame(df)
xlsxfile = 'CEOutput.xls'
excel_writer = pd.ExcelWriter(xlsxfile, engine='xlsxwriter')
outputdf.to_excel(excel_writer, sheet_name="SurveyIDandCE",startrow=0, startcol=0, header = True, index=True)
excel_writer.save()