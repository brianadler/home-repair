from cmath import sqrt
from tkinter import Y
from tokenize import String
import pandas as pd
import os
import math
import numpy

#Read in second file with square footage and num of stories tied to survey ID
file_name = r"C:\Users\brian\Downloads\EXCEL_Survey 1 Returns Dataset_ALL.xlsx"
sheet = "S1_583_RETURNS"
answerkey = pd.read_excel(io = file_name, sheet_name=sheet)

df=pd.DataFrame(answerkey, columns=['Survey_ID','building_square_footage', 'building_stories', 'Combined', 'EstCost'])


#Convert dataframes to a single dimensional array
arr = df.to_numpy()

#Convert arrays to type string to process by character later in code
stringarr2 = arr.astype(str)

count = 0
ec = 0

#introduce cost figures
averagesize = 1390
btu = 1.249
serviceheatingequipment = 284
replaceheatingequipment = 1737
serviceacequipment = 267
replaceinstallacequipment = 3189
replaceacequipment = 2100
wireforelectric = 284
installelectricplugs = 284
concealwiring = 284
#for electrical, beware of multiplier
minorelectrical = 284
upgradelectricservice = 1200
replacebreaker = 1500
moldremediation = 220
repairtoilet = 246
replacetoilet = 459.5
replacepiping = 260
replacewaterheater = 2044
repairsewerline = 2904
snakeraugerline = 405
snakeaugerdrane = 273
exterminatetermite = 577
exterminaterodent = 354
exterminateseal = 596
sealbasement = 1500
sealwindow = 230
sealroof = 454
repairwallsreplacepiping = 515
replacepipingforfaucettoilet = 157
repairfoundationmultiplier = 11.49
repairroof = 696
replaceroofmaterials = 3231.75
replaceroofmaterialsmultiplier = 3.72
replaceroof = 8699
replaceroofmultiplier = 5.01
replacegutters = 1502.95
repairexteriorwall = 785
replaceexteriorwallsagging = 2100
tuckpointing = 1000
repainting = 445.44
repaintingwindowsashes = 2900
replacewindow = 549.4
repairfloor = 411.25
repairfloortoiletsink = 760.15
repairinteriorwall = 1024.88
replaceinstalllocks = 187
treeremovalpruning = 750
repairexternalwalkway = 385.8
repairpatio = 385.8
repairporchdeck = 242
repairexternalstairway = 695
weatherization1 = 1542
weatherization2 = 275
repairadditionalstructure = 2921
tempsqfootage = 0
sealroofmultiplier = 5.226
replaceexteriorwallmultiplier = 1.51
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

#Below is the loop and all of the if conditions that will procure a cost estimate for each survey respondent
for x in stringarr2:
    #The line below looks at the second string entry in x. if we did x[0], we would look at the identifier instead
    item = x[3]
    squarefootage = float(x[1])
    numfloors = float (x[2])
    if (item[0:1]) == '1':
        ec = ec + serviceheatingequipment
        serviceheatcount = serviceheatcount + 1
    if (item[0:1]) == '2':
        ec = ec + (btu*squarefootage)
        replaceheatcount = replaceheatcount + 1
    #Doing a multiple check
    if (item[1:2]) == '1' and (item[2:3]) == 3 and serviceheatcount == 0:
        ec = ec + serviceheatingequipment
        serviceheatcount = serviceheatcount + 1
    if (item[1:2]) == '1' and (item[2:3]) == 4 and serviceheatcount == 0:
        ec = ec + serviceheatingequipment
        serviceheatcount = serviceheatcount + 1
    if (item[1:2]) == '2' and (item[2:3]) == 3 and replaceheatcount == 0:
        ec = ec + (btu*squarefootage)
        replaceheatcount = replaceheatcount + 1
    if (item[1:2]) == '2' and (item[2:3]) == 4 and replaceheatcount == 0:
        ec = ec + (btu*squarefootage)
        replaceheatcount = replaceheatcount + 1
    if (item[1:2]) == '1' and (item[2:3]) == 1 and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[1:2]) == '1' and (item[2:3]) == 2 and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[1:2]) == '2' and (item[2:3]) == 1 and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[1:2]) == '2' and (item[2:3]) == 2 and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if(item[2:3]) == '1' and weatherizationcount == 0:
        ec = ec + weatherization1
    if(item[2:3]) == '2' and weatherizationcount == 0:
        ec = ec + weatherization1

    if (item[3:4]) == '1' and serviceacequipment == 0:
        #print ("Service AC Equipment")
        ec = ec + serviceacequipment
        serviceacequipment = serviceacequipment + 1
    if (item[3:4]) == '2' and replaceacequipment == 0:
       # print ("Replace/Install AC Equipment")
        ec = ec + replaceinstallacequipment
        replaceacequipment = replaceacequipment + 1
    if (item[3:4]) == '4' and replaceacequipment == 0:
       # print ("Replace/Install AC Equipment")
        ec = ec + replaceinstallacequipment
        replaceacequipment = replaceacequipment + 1
    if (item[4:5]) == '1' and (item[2:3]) == '1' and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[4:5]) == '1' and (item[2:3]) == '2' and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[4:5]) == '1' and (item[2:3]) == '3' and serviceheatcount == 0:
        ec = ec + serviceacequipment
        serviceheatcount = serviceheatcount + 1
    if (item[4:5]) == '1' and (item[2:3]) == '4' and serviceheatcount == 0:
        ec = ec + serviceacequipment
        serviceheatcount = serviceheatcount + 1
    if (item[4:5]) == '2' and (item[2:3]) == '1' and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[4:5]) == '2' and (item[2:3]) == '2' and weatherizationcount == 0:
        ec = ec + weatherization1
        weatherizationcount = weatherizationcount + 1
    if (item[4:5]) == '2' and (item[2:3]) == '3':
        ec = ec + replaceacequipment
    if (item[4:5]) == '2' and (item[2:3]) == '4' and serviceheatcount == 0:
        ec = ec + replaceacequipment

    # Moving on to question bunch #2
    if (item[5:6]) == '1':
      #  print ("Wire Rooms for Electricity")
        ec = ec + wireforelectric
    if (item[5:6]) == '2':
       # print ("Wire Rooms for Electricity")
        ec = ec + wireforelectric
    if (item[6:7]) == '1':
        ec = ec + installelectricplugs
        electricplugcount = electricplugcount + 1
    if (item[6:7]) == '2':
        ec = ec + installelectricplugs
        electricplugcount = electricplugcount + 1
    if (item[7:8]) == '1':
        ec = ec + (concealwiring * 1.5)
    if (item[7:8]) == '2':
        ec = ec + (concealwiring * 3)
    if (item[8:9]) == '1':
        ec = ec + (minorelectrical * 1.5)
    if (item[8:9]) == '2':
        ec = ec + (minorelectrical * 3)
    if (item[9:10]) == '1' and electricplugcount == 0:
        ec = ec + installelectricplugs
        electricplugcount = electricplugcount + 1
    if (item[9:10]) == '2' and electricplugcount == 0:
        ec = ec + installelectricplugs
        electricplugcount = electricplugcount + 1
    if (item[10:11]) == '1' and (item[11:12]) == '1' and upgradeelectriccount == 0:
        ec = ec + upgradelectricservice
        upgradeelectriccount = upgradeelectriccount + 1
    if (item[10:11]) == '1' and (item[11:12]) == '2' and replacebreakercount == 0:
        ec = ec + replacebreaker
        replacebreakercount = replacebreakercount + 1
    if (item[10:11]) == '2' and (item[11:12]) == '1' and upgradeelectriccount == 0:
        ec = ec + upgradelectricservice
        upgradeelectriccount = upgradeelectriccount + 1
    if (item[10:11]) == '2' and (item[11:12]) == '2' and replacebreakercount == 0:
        ec = ec + replacebreaker
        replacebreakercount = replacebreakercount + 1
 
    # Moving on to question bunch #3
    if (item[12:13]) == '1':
        ec = ec + moldremediation
    if (item[12:13]) == '2':
        ec = ec + moldremediation
    if (item[13:14]) == '1':
        ec = ec + repairtoilet
    if (item[13:14]) == '2':
        ec = ec + repairtoilet
    if (item[14:15]) == '1':
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[14:15]) == '2':
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[15:16]) == '1':
        ec = ec + replacewaterheater
    if (item[15:16]) == '2':
        ec = ec + replacewaterheater
    if (item[16:17]) == '1':
        ec = ec + snakeraugerline
        snakeaugercount = snakeaugercount + 1
    if (item[16:17]) == '2':
        ec = ec + repairsewerline
    if (item[17:18]) == '1':
        ec = ec + snakeaugerdrane
    if (item[17:18]) == '2' and replacepipingcount == 0:
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[18:19]) == '1' and snakeaugercount == 0:
        ec = ec + snakeaugerdrane
        snakeaugercount = snakeaugercount + 1
    if (item[18:19]) == '2' and replacepipingcount == 0:
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[19:20]) == '1':
        ec = ec + exterminatetermite
        exterminatorcount = exterminatorcount + 1
    if (item[19:20]) == '2':
        ec = ec + exterminatetermite
        exterminatorcount = exterminatorcount + 1
    if (item[20:21]) == '1':
        ec = ec + exterminaterodent
    if (item[20:21]) == '2':
        ec = ec + exterminateseal

    # Question Block 4
    if (item[21:22]) == '1':
        ec = ec + sealbasement
    if (item[21:22]) == '2':
        ec = ec + sealbasement
    if (item[22:23]) == '1':
        ec = ec + sealwindow
    if (item[22:23]) == '2':
        ec = ec + sealwindow
    if (item[23:24]) == '1':
        roofsquarefootage = .05*(1.25*squarefootage)
        ec = ec + (roofsquarefootage * sealroofmultiplier)
    if (item[23:24]) == '2':
        roofsquarefootage = .05*(1.25*squarefootage)
        ec = ec + (roofsquarefootage * sealroofmultiplier)
    if (item[24:25]) == '1' and replacepipingcount == 0:
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[24:25]) == '2' and replacepipingcount == 0:
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[25:26]) == '1':
        ec = ec + repairwallsreplacepiping
    if (item[25:26]) == '2':
        ec = ec + repairwallsreplacepiping
    if (item[26:27]) == '1':
        ec = ec + replacewaterheater
    if (item[26:27]) == '2':
        ec = ec + replacewaterheater
    if (item[27:28]) == '1' and replacepipingcount == 0:
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[27:28]) == '2' and replacepipingcount == 0:
        ec = ec + replacepiping
        replacepipingcount = replacepipingcount + 1
    if (item[28:29]) == '1' and snakeaugercount == 0:
        ec = ec + snakeaugerdrane
        snakeaugercount = snakeaugercount + 1
    if (item[28:29]) == '2' and snakeaugercount == 0:
        ec = ec + snakeaugerdrane
        snakeaugercount = snakeaugercount + 1

    # Question Block 5
    if (item[29:30]) == '1':
        basementwallarea = (sqrt(squarefootage/numfloors)) * 4
        ec = ec + basementwallarea * repairfoundationmultiplier
    if (item[29:30]) == '2':
        basementwallarea = (sqrt(squarefootage/numfloors)) * 4
        ec = ec + basementwallarea * repairfoundationmultiplier
    if (item[30:31]) == '1':
        ec = ec + repairroof
    if (item[30:31]) == '2':
        ec = ec + repairroof
    if (item[31:32]) == '1':
       roofarea = 1.25*(squarefootage/numfloors)
       ec = ec + (roofarea * replaceroofmaterialsmultiplier)
    if (item[31:32]) == '2':
       roofarea = 1.25*(squarefootage/numfloors)
       ec = ec + (roofarea * replaceroofmaterialsmultiplier)
    if (item[32:33]) == '1':
       roofarea = 1.25*(squarefootage/numfloors)
       ec = ec + (roofarea * replaceroofmultiplier)
    if (item[32:33]) == '2':
        roofarea = 1.25*(squarefootage/numfloors)
        ec = ec + (roofarea * replaceroofmultiplier)
    if (item[33:34]) == '1':
        ec = ec + replacegutters
    if (item[33:34]) == '2':
        ec = ec + replacegutters
    if (item[34:35]) == '1':
        ec = ec + tuckpointing
    if (item[34:35]) == '2':
        ec = ec + tuckpointing
    if (item[35:36]) == '1':
        ec = ec + repainting
        repaintingcount = repaintingcount+1
    if (item[35:36]) == '2':
        ec = ec + repainting
        repaintingcount = repaintingcount+1
    if (item[36:37]) == '1':
        ec = ec + (squarefootage * replaceexteriorwallmultiplier)
    if (item[36:37]) == '2':
        ec = ec + (squarefootage * replaceexteriorwallmultiplier)
    if (item[37:38]) == '1':
        ec = ec + replacewindow
    if (item[37:38]) == '2':
        ec = ec + replacewindow

    # Question Block 6
    if (item[38:39]) == '1':
        ec = ec + repairfloor
        repairfloorcount = repairfloorcount + 1
    if (item[38:39]) == '2':
        ec = ec + repairfloor
        repairfloorcount = repairfloorcount + 1
    if (item[39:40]) == '1':
        ec = ec + repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if (item[39:40]) == '2':
        ec = ec + repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if (item[40:41]) == '1' and repaintingcount == 0:
        ec = ec + repainting
        repaintingcount = repaintingcount+1
    if (item[40:41]) == '2' and repaintingcount == 0:
        ec = ec + repainting
        repaintingcount=repaintingcount+1
    if (item[41:42]) == '1' and repairinteriorwallcount == 0:
        ec = ec + repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if (item[41:42]) == '2' and repairinteriorwallcount == 0:
        ec = ec + repairinteriorwall
        repairinteriorwallcount = repairinteriorwallcount + 1
    if (item[42:43]) == '1':
        ec = ec + repaintingwindowsashes
    if (item[42:43]) == '2':
        ec = ec + repaintingwindowsashes
    if (item[43:44]) == '1' and repairfloorcount == 0:
        ec = ec + repairfloor
        repairfloorcount = repairfloorcount + 1
    if (item[43:44]) == '2' and repairfloorcount == 0:
        ec = ec + repairfloor
        repairfloorcount = repairfloorcount + 1
    if (item[44:45]) == '1':
        ec = ec + repairfloortoiletsink
    if (item[44:45]) == '2':
        ec = ec + repairfloortoiletsink
    if (item[45:46]) == '1':
        ec = ec + replaceinstalllocks
    if (item[45:46]) == '2':
        ec = ec + replaceinstalllocks

    # Question Block 7
    if (item[46:47]) == '1':
        ec = ec + treeremovalpruning
    if (item[46:47]) == '2':
        ec = ec + treeremovalpruning
    if (item[47:48]) == '1':
        ec = ec + repairexternalwalkway
    if (item[47:48]) == '2':
        ec = ec + repairexternalwalkway
    if (item[48:49]) == '1':
        ec = ec + repairpatio
    if (item[48:49]) == '2':
        ec = ec + repairpatio
    if (item[49:50]) == '1':
        ec = ec + repairporchdeck
    if (item[49:50]) == '2':
        ec = ec + repairporchdeck
    if (item[50:51]) == '1':
        ec = ec + repairexternalstairway
    if (item[50:51]) == '2':
        ec = ec + repairexternalstairway
    if (item[51:52]) == '1':
        ec = ec + repairadditionalstructure
    if (item[51:52]) == '2':
        ec = ec + repairadditionalstructure

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
    if numpy.isnan(ec):
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