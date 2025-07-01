# #######################################################
#
# Script de correction du fichier HTML exporté depuis InDesign (cible : Word Métopes) avant ouverture dans Word
# Le script effectue les deux actions suivantes :
# – renommage automatique des ancres de type "<a id='...'>" en "<a name='...'>"
# – correction des appels de liens pour qu'ils ne soient plus relatifs au fichier HTML, ex. :
#       <a href="geo-47-5-propre.html#_idTextAnchor007">
#       becomes
#       <a href="#_idTextAnchor007">
#
# #######################################################
# Auteur : Emmanuel Côtez (emmanuel.cotez@mnhn.fr)
# version : 1.4-os_free
#######################################################

import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox as msgbox
import time
import os
import sys

# alert window functions
# https://stackoverflow.com/questions/2963263/how-can-i-create-a-simple-message-box-in-python
def Mbox(title, text, style=0):
    if style == 0:
        msgbox.showinfo(title, text)
    elif style == 1:
        return msgbox.askokcancel(title, text)
    else:
        msgbox.showinfo(title, text)
    ##  Styles:
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | Cancel 
    ##  6 : Cancel | Try Again | Continue

print ("###########################################")
print ("#### Correct InDesign HTML file (v1.4) ####")
print ("###########################################")

# Select and Read the file
val=True;
print ("---------------------------")
print ('Select an HTML file (won\'t be modified)')

# Testing if the opened file is an HTML file
while val:
    Tk().withdraw()
    selectedfile = askopenfilename(title='Select an HTML file (won\'t be modified)')
    
    if selectedfile == '' or selectedfile is None:
        print("No file selected, exiting...")
        exit(0)
        
    file_name = os.path.basename(selectedfile)
    file_folder = os.path.dirname(os.path.realpath(selectedfile))
    
    if file_name.lower().endswith('.html'):
        val = False
    else:
        print(f"\nThe file {file_name} is not an HTML file.\nPlease select an HTML file")


file = open(selectedfile, "r+", encoding="utf-8")
file_contents = file.read()
print("\nSelected file: " + selectedfile)
print ("---------------------------") ;

# ####################################################################
# Step 1 #############################################################
# Rename the anchors #################################################
# ####################################################################

print ("Step 1: Renaming anchors in the HTML file")
time.sleep(0.5)
text = r'<a id="(_idTextAnchor\d+)">'
text_replace = r'<a name="\1">'
result = re.search(text, file_contents)
num = re.subn(text, text_replace, file_contents)
result = re.compile(text)

i=0
for match in result.finditer(file_contents):
    file_contents = re.sub(text, text_replace, file_contents)
    i=i+1
    print(str(i)+"/"+str(num[1])+" anchor(s) treated ("+str(round(i/num[1]*100))+"%)", end='\r', flush=True)
    time.sleep(0.01)
    
result_step1 = int(num[1])  
    

print ("\n" + str(result_step1) + " anchor(s) renamed")
print ("---------------------------") ;

time.sleep(0.5)

# ####################################################################
# Step 2 #############################################################
# Correct the links to anchors #######################################
# ####################################################################

print ("Step 2: Correcting the links to anchors in the HTML file")
time.sleep(0.5)
text = r'<a\s?href="[\w\s0-9_\.-]*\.html(#_idTextAnchor\d+)">'
# text = r'<a\s?href="geo-47-5-propre.html(#_idTextAnchor\d+)">'
text_replace = r'<a href="\1">'
result = re.search(text, file_contents)
num = re.subn(text, text_replace, file_contents)
result = re.compile(text)

i=0
for match in result.finditer(file_contents):
    file_contents = re.sub(text, text_replace, file_contents)
    i=i+1
    print(str(i)+"/"+str(num[1])+" links treated ("+str(round(i/num[1]*100))+"%)", end='\r', flush=True)
    time.sleep(0.01)
    
result_step2 = int(num[1])

print ("\n" + str(result_step2) + " links to anchors corrected")
print ("---------------------------") ;

time.sleep(0.5)

# Writing the file under a new name (only if some modifications have been applied)
if ((result_step1==0) & (result_step2==0)):
    Mbox('File not saved', 'No tag to correct found in the HTML file!\n\nThe file ' + file_name + ' has not been modified', 0)
else:
    target_file_name = selectedfile.replace(".html","_corrected_"+str(time.time())+".html");
    target_file = open(target_file_name, "x", encoding="utf-8")
    target_file.seek(0)
    target_file.truncate()
    target_file.write(file_contents)
    # information window
    Mbox('File saved', 'The file \n— ' + file_name + '\nhas been treated (HTML corrected) and saved as\n— '+target_file_name+'\nin the folder\n— '+file_folder+'', 0) 
    
time.sleep(0.5)

# To prevent the window from closing
# input('Press ENTER to exit')

