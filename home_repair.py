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


df["serviceheatingequipment"] = ((df.Q1 == "1") | ((df.Q2 == "1") & ((df.Q3 == "3") | (df.Q3 == "4"))))
df["replaceheatingequipment"] = ((df.Q1 == "2") | ((df.Q2 == "2") & ((df.Q3 == "3") | (df.Q3 == "4"))))
df["weatherization1"] = ((df.Q2 == "1") | ((df.Q2 == "2") & ((df.Q3 == "1") | (df.Q3 == "2"))))

#Below is the loop and all of the if conditions that will procure a cost estimate for each survey respondent
for x in stringarr2:
    #The line below looks at the second string entry in x. if we did x[0], we would look at the identifier instead
    item = x[3]
    squarefootage = float(x[1])
    numfloors = float (x[2])
    # if item[0] == '1':
    #     ec = ec + costs.serviceh
    # eatingequipment
    #     serviceheatcount = serviceheatcount + 1
    # if item[0] == '2':
    #     ec = ec + (costs.btu*squarefootage)
    #     replaceheatcount = replaceheatcount + 1
    # #Doing a multiple check
    # if item[1] == '1' and item[2] == 3 and serviceheatcount == 0:
    #     ec = ec + costs.serviceheatingequipment
    #     serviceheatcount = serviceheatcount + 1
    # if item[1] == '1' and item[2] == 4 and serviceheatcount == 0:
    #     ec = ec + costs.serviceheatingequipment
    #     serviceheatcount = serviceheatcount + 1
    #if item[1] == '2' and item[2] == 3 and replaceheatcount == 0:
    #    ec = ec + (costs.btu*squarefootage)
    #    replaceheatcount = replaceheatcount + 1
    #if item[1] == '2' and item[2] == 4 and replaceheatcount == 0:
    #   ec = ec + (costs.btu*squarefootage)
    #    replaceheatcount = replaceheatcount + 1
    #if item[1] == '1' and item[2] == 1 and weatherizationcount == 0:
    #    ec = ec + costs.weatherization1
    #    weatherizationcount = weatherizationcount + 1
    #if item[1] == '1' and item[2] == 2 and weatherizationcount == 0:
    #    ec = ec + costs.weatherization1
    #    weatherizationcount = weatherizationcount + 1
    #if item[1] == '2' and item[2] == 1 and weatherizationcount == 0:
    #    ec = ec + costs.weatherization1
    #    weatherizationcount = weatherizationcount + 1
    #if item[1] == '2' and item[2] == 2 and weatherizationcount == 0:
    #    ec = ec + costs.weatherization1
    #    weatherizationcount = weatherizationcount + 1
    #if item[2] == '1' and weatherizationcount == 0:
    #    ec = ec + costs.weatherization1
    #if item[2] == '2' and weatherizationcount == 0:
    #    ec = ec + costs.weatherization1

    if item[3] == '1' and serviceacequipmentcount == 0:
        #print ("Service AC Equipment")
        ec = ec + costs.serviceacequipment
        serviceacequipmentcount = serviceacequipmentcount + 1
    if item[3] == '2' and replaceacequipment == 0:
       # print ("Replace/Install AC Equipment")
        ec = ec + costs.replaceinstallacequipment
        replaceacequipment = replaceacequipment + 1
    if item[3] == '4' and replaceacequipment == 0:
       # print ("Replace/Install AC Equipment")
        ec = ec + costs.replaceinstallacequipment
        replaceacequipment = replaceacequipment + 1
    if item[4] == '1' and item[2] == '1' and weatherizationcount == 0:
        ec = ec + costs.weatherization1
        weatherizationcount = weatherizationcount + 1
    if item[4] == '1' and item[2] == '2' and weatherizationcount == 0:
        ec = ec + costs.weatherization1
        weatherizationcount = weatherizationcount + 1
    if item[4] == '1' and item[2] == '3' and serviceheatcount == 0:
        ec = ec + costs.serviceheatingequipment
        serviceheatcount = serviceheatcount + 1
    if item[4] == '1' and item[2] == '4' and serviceheatcount == 0:
        ec = ec + costs.serviceheatingequipment
        serviceheatcount = serviceheatcount + 1
    if item[4] == '2' and item[2] == '1' and weatherizationcount == 0:
        ec = ec + costs.weatherization1
        weatherizationcount = weatherizationcount + 1
    if item[4] == '2' and item[2] == '2' and weatherizationcount == 0:
        ec = ec + costs.weatherization1
        weatherizationcount = weatherizationcount + 1
    if item[4] == '2' and item[2] == '3':
        ec = ec + replaceacequipment
    if item[4] == '2' and item[2] == '4' and serviceheatcount == 0:
        ec = ec + replaceacequipment

    # Moving on to question bunch #2
    if item[5] == '1':
      #  print ("Wire Rooms for Electricity")
        ec = ec + costs.wireforelectric
    if item[5] == '2':
       # print ("Wire Rooms for Electricity")
        ec = ec + costs.wireforelectric
    if item[6] == '1':
        ec = ec + costs.installelectricplugs
        electricplugcount = electricplugcount + 1
    if item[6] == '2':
        ec = ec + costs.installelectricplugs
        electricplugcount = electricplugcount + 1
    if item[7] == '1':
        ec = ec + (costs.concealwiring * 1.5)
    if item[7] == '2':
        ec = ec + (costs.concealwiring * 3)
    if item[8] == '1':
        ec = ec + (costs.minorelectrical * 1.5)
    if item[8] == '2':
        ec = ec + (costs.minorelectrical * 3)
    if item[9] == '1' and electricplugcount == 0:
        ec = ec + costs.installelectricplugs
        electricplugcount = electricplugcount + 1
    if item[9] == '2' and electricplugcount == 0:
        ec = ec + costs.installelectricplugs
        electricplugcount = electricplugcount + 1
    if item[10] == '1' and item[11] == '1' and upgradeelectriccount == 0:
        ec = ec + costs.upgradelectricservice
        upgradeelectriccount = upgradeelectriccount + 1
    if item[10] == '1' and item[11] == '2' and replacebreakercount == 0:
        ec = ec + costs.replacebreaker
        replacebreakercount = replacebreakercount + 1
    if item[10] == '2' and item[11] == '1' and upgradeelectriccount == 0:
        ec = ec + costs.upgradelectricservice
        upgradeelectriccount = upgradeelectriccount + 1
    if item[10] == '2' and item[11] == '2' and replacebreakercount == 0:
        ec = ec + costs.replacebreaker
        replacebreakercount = replacebreakercount + 1
 
    # Moving on to question bunch #3
    if item[12] == '1':
        ec = ec + costs.moldremediation
    if item[12] == '2':
        ec = ec + costs.moldremediation
    if item[13] == '1':
        ec = ec + costs.repairtoilet
    if item[13] == '2':
        ec = ec + costs.repairtoilet
    if item[14] == '1':
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[14] == '2':
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[15] == '1':
        ec = ec + costs.replacewaterheater
    if item[15] == '2':
        ec = ec + costs.replacewaterheater
    if item[16] == '1':
        ec = ec + costs.snakeraugerline
        snakeaugercount = snakeaugercount + 1
    if item[16] == '2':
        ec = ec + costs.repairsewerline
    if item[17] == '1':
        ec = ec + costs.snakeaugerdrane
    if item[17] == '2' and replacepipingcount == 0:
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[18] == '1' and snakeaugercount == 0:
        ec = ec + costs.snakeaugerdrane
        snakeaugercount = snakeaugercount + 1
    if item[18] == '2' and replacepipingcount == 0:
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[19] == '1':
        ec = ec + costs.exterminatetermite
        exterminatorcount = exterminatorcount + 1
    if item[19] == '2':
        ec = ec + costs.exterminatetermite
        exterminatorcount = exterminatorcount + 1
    if item[20] == '1':
        ec = ec + costs.exterminaterodent
    if item[20] == '2':
        ec = ec + costs.exterminateseal

    # Question Block 4
    if item[21] == '1':
        ec = ec + costs.sealbasement
    if item[21] == '2':
        ec = ec + costs.sealbasement
    if item[22] == '1':
        ec = ec + costs.sealwindow
    if item[22] == '2':
        ec = ec + costs.sealwindow
    if item[23] == '1':
        roofsquarefootage = .05*(1.25*squarefootage)
        ec = ec + (roofsquarefootage * costs.sealroofmultiplier)
    if item[23] == '2':
        roofsquarefootage = .05*(1.25*squarefootage)
        ec = ec + (roofsquarefootage * costs.sealroofmultiplier)
    if item[24] == '1' and replacepipingcount == 0:
        ec = ec +costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[24] == '2' and replacepipingcount == 0:
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[25] == '1':
        ec = ec + costs.repairwallsreplacepiping
    if item[25] == '2':
        ec = ec + costs.repairwallsreplacepiping
    if item[26] == '1':
        ec = ec + costs.replacewaterheater
    if item[26] == '2':
        ec = ec + costs.replacewaterheater
    if item[27] == '1' and replacepipingcount == 0:
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[27] == '2' and replacepipingcount == 0:
        ec = ec + costs.replacepiping
        replacepipingcount = replacepipingcount + 1
    if item[28] == '1' and snakeaugercount == 0:
        ec = ec + costs.snakeaugerdrane
        snakeaugercount = snakeaugercount + 1
    if item[28] == '2' and snakeaugercount == 0:
        ec = ec + costs.snakeaugerdrane
        snakeaugercount = snakeaugercount + 1

    # Question Block 5
    if item[29] == '1':
        basementwallarea = (sqrt(squarefootage/numfloors)) * 4
        ec = ec + basementwallarea * costs.repairfoundationmultiplier
    if item[29] == '2':
        basementwallarea = (sqrt(squarefootage/numfloors)) * 4
        ec = ec + basementwallarea * costs.repairfoundationmultiplier
    if item[30] == '1':
        ec = ec + costs.repairroof
    if item[30] == '2':
        ec = ec + costs.repairroof
    if item[31] == '1':
       roofarea = 1.25*(squarefootage/numfloors)
       ec = ec + (roofarea * costs.replaceroofmaterialsmultiplier)
    if item[31] == '2':
       roofarea = 1.25*(squarefootage/numfloors)
       ec = ec + (roofarea * costs.replaceroofmaterialsmultiplier)
    if item[32] == '1':
       roofarea = 1.25*(squarefootage/numfloors)
       ec = ec + (roofarea * costs.replaceroofmultiplier)
    if item[32] == '2':
        roofarea = 1.25*(squarefootage/numfloors)
        ec = ec + (roofarea * costs.replaceroofmultiplier)
    if item[33] == '1':
        ec = ec + costs.replacegutters
    if item[33] == '2':
        ec = ec + costs.replacegutters
    if item[34] == '1':
        ec = ec + costs.tuckpointing
    if item[34] == '2':
        ec = ec + costs.tuckpointing
    if item[35] == '1':
        ec = ec + costs.repainting
        repaintingcount = repaintingcount+1
    if item[35] == '2':
        ec = ec + costs.repainting
        repaintingcount = repaintingcount+1
    if item[36] == '1':
        ec = ec + (squarefootage * costs.replaceexteriorwallmultiplier)
    if item[36] == '2':
        ec = ec + (squarefootage * costs.replaceexteriorwallmultiplier)
    if item[37] == '1':
        ec = ec + costs.replacewindow
    if item[37] == '2':
        ec = ec + costs.replacewindow

    # Question Block 6
    if item[38] == '1':
        ec = ec + costs.repairfloor
        repairfloorcount = repairfloorcount + 1
    if item[38] == '2':
        ec = ec + costs.repairfloor
        repairfloorcount = repairfloorcount + 1
    if item[39] == '1':
        ec = ec + costs.repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if item[39] == '2':
        ec = ec + costs.repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if item[40] == '1' and repaintingcount == 0:
        ec = ec + costs.repainting
        repaintingcount = repaintingcount+1
    if item[40] == '2' and repaintingcount == 0:
        ec = ec + costs.repainting
        repaintingcount=repaintingcount+1
    if item[41] == '1' and repairinteriorwallcount == 0:
        ec = ec + costs.repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if item[41] == '2' and repairinteriorwallcount == 0:
        ec = ec + costs.repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if item[42] == '1':
        ec = ec + costs.repaintingwindowsashes
    if item[42] == '2':
        ec = ec + costs.repaintingwindowsashes
    if item[43] == '1' and repairfloorcount == 0:
        ec = ec + costs.repairfloor
        repairfloorcount = repairfloorcount + 1
    if item[43] == '2' and repairfloorcount == 0:
        ec = ec + costs.repairfloor
        repairfloorcount = repairfloorcount + 1
    if item[44] == '1':
        ec = ec + costs.repairfloortoiletsink
    if item[44] == '2':
        ec = ec + costs.repairfloortoiletsink
    if item[45] == '1':
        ec = ec + costs.replaceinstalllocks
    if item[45] == '2':
        ec = ec + costs.replaceinstalllocks

    # Question Block 7
    if item[46] == '1':
        ec = ec + costs.treeremovalpruning
    if item[46] == '2':
        ec = ec + costs.treeremovalpruning
    if item[47] == '1':
        ec = ec + costs.repairexternalwalkway
    if item[47] == '2':
        ec = ec + costs.repairexternalwalkway
    if item[48] == '1':
        ec = ec + costs.repairpatio
    if item[48] == '2':
        ec = ec + costs.repairpatio
    if item[49] == '1':
        ec = ec + costs.repairporchdeck
    if item[49] == '2':
        ec = ec + costs.repairporchdeck
    if item[50] == '1':
        ec = ec + costs.repairexternalstairway
    if item[50] == '2':
        ec = ec + costs.repairexternalstairway
    if item[51] == '1':
        ec = ec + costs.repairadditionalstructure
    if item[51] == '2':
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

xlsxfile = 'CEOutput.xls'
excel_writer = pd.ExcelWriter(xlsxfile, engine='xlsxwriter')
outputdf.columns=['SurveyID','SquareFootage','NumStories','SurveyResponses','Est. Cost']
outputdf.to_excel(excel_writer, sheet_name="SurveyIDandCE",startrow=0, startcol=0, header = True, index=True)
excel_writer.save()