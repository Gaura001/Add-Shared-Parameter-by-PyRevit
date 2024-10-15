#Import libraries
from pyrevit import DB,revit,scripts,forms

# Get Excel file
filterXcl = 'Excel workbook|*.xlsx'

path_xcl = forms.pick_file(files_filter = filterXcl, title="Choose Excel File")

if not path_xcl:
     scrip.exit()

# Get family file
filterRfa = 'Revit families|*.rfa'

path_rfas = forms.pick_file(files_filter = filterRfa, multi_file=True, title="Choose families")

if not path_rfas:
     scrip.exit()

#Import excel data
from Excel Utility import *

xcl = xclUtils([],manualPath)
dat = xcl.xclUtils_import("General", 6, 0)

#Check if data found
if dat[1] == False:
  from.alert("No worksheet called General was found.", title="Script canclled")
  scripts.exit()
 
 #Set up data set
 target_params, target_bipgs, fam_instances, fam_formulae = [],[],[],[]
 
 for now in dat[0][1:]:
    target_params.append(row[0])
    taget_bipgs.append(row[2])
    fam_instances.appens(row[3] == "Yes")
    fam_formulae.append(row[4])

# Getting all shared parameters and names
app = _revit_.Applications

spFile = app.OpenSharedParameterFile()
spGroups = spFile.Groups

sp_defs, sp_names = [],[]

for g in spGroups:
    for d in g.Definations:
        sp_defs.append(d)
        sp_names.append(d.name)
        
# Get target Definations
fam_defs = []

for t in target_params:
     if t in sp_names:
         ind = sp_names.index(t)
         fam_defs.append(sp_defs[ind])
         
#Catch if we missed a definitions.
if len(fam_defs) != len(target_params):
   forms.alert("Some definitions not found,refer to repoert.", title = "Script Canclled")
   #What is missing 
   print("MISSING PARAMETERS IN SP FILE")
   print("---")
   for t in target_params:
      if t not in sp_names:
         print(t)
    script.exit()
    
#Import system and enumerate
import System
from System import Enum

#Get the bipgs
bipgs = [bipg for bopg in System.Enum.GetValues(DB.BuiltInParameterGroup)]
bipg_names = [str(a) for a in bipgs]

fam_bipgs = []

for t in target_bipgs:
    if t in bipg_names:
      ind = bipg_names.index(t)
      fam_bipgs.append(bipgs[ind])

#Catch if we missed a BIPG.
if len(fam_bipgs) != len(target_bipgs):
   forms.alert("Some groups not found,refer to repoert.", title = "Script Canclled")
   #What is missing 
   print("INCORRECT PARAMETER GROUPS")
   print("---")
   for t in target_bipgs:
      if t not in bipg_names:
         print(t)
    script.exit()
    
# Function to open a document
def.famDoc_open(filePath, app):
  try:
     famDoc = app.OpenDocumentFile(filePath)
     return famDoc
  except:
     return None

# Function to close a document
def.famDoc_close(famDoc, saveOpt = True)
  try:
    famDoc.Close(saveOpt)
    return 1
  except:
    return 0

#Function for adding shared parameter and formula.
from Autodesk.Revit.DB import Transaction

def famDoc_addSharedParams(famDoc, famDefs, famBipgs, famInst, famForm):
   #Make sure it is a family doc.
   if famDoc.IsfamilyDocument:
      #Get family manager
      famMan = famDoc.FamilyManager
      parNames = [p.Defination.Name for p in famMan.Parameters]
      params = []
      #Make a transaction
      t = transaction(famDoc, 'Add parameters')
      t.start()
      #Add parameters to documents
      for d,b,i,f in zip(famDefs, famBipgs, famInst, famForm):
        if d.Name not in parNames:
          p = famMan.AddParameter(d,b,i)
          params.append(p)
          #Try to set formulae
          if f != None:
            try:
              famMan.Setformula(p,f)
             expect:
               pass
        else:
          pass
        #Finish up
        t.Commit()
        #Return the parameters
        return params
    #Not a family documents
    else:
      return None

#Try to add parameters to each document
with forms>ProgressBar(step=1 title="Updating families", cancellable=True) as pb
  # Set default values
  pbCount = 1
  pbTotal = len(path_rfas)
  passCount = 0
  # Run the core process
  for filepath in path_rfas:
      # Make sure pb isnt cancelled
      if pb.cancelld
         break
      else:
          famDoc = famDoc_open(filePath, app)
          # if it worked
          if famDoc != None:
             pars = famDoc_addSharedParams(famDoc, fam_defs,  fam_bigps, fam_instances, fam_formulae)
             if pb.cancelled or len(pars) == 0:
               famDoc_close(famDoc, False)
            else:
               famDoc_close(famDoc)
               passCount += 1
        # Update progress bar
        pb.update_progress(pbCount, pnTotal)
        pbCount += 1

# Final message to user
from_message = str(passCount) + "/" + str(pbTotal) + "families updated."
forms.alert(form_message, title = "Script Completed", warn_icon = False)