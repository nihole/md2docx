import win32com.client as win32
import re
import sys
import os
import yaml

wdGoToSection = 0
wdCaptionPositionAbove = 0
wdCaptionPositionBelow = 1
wdGoToFirst = 1
wdSection = 8
wdSectionBreakContinuous = 3
wdGoToLine = 3
wdLine = 5
wdGoToHeading = 11
wdParagraph = 4


def insert_data (sec_, md_filepath_): 
    
    doc.Sections(sec_).Range.Delete()
    word.Selection.InsertBreak(Type=wdSectionBreakContinuous)
    word.Selection.GoTo(What=wdGoToSection, Which=wdGoToFirst, Count=sec_)
    tmp_doc_path = dirpath + "/tmp/tmp_word.docx"
    md_full_path = dirpath + "/" + md_filepath_


    os.system("pandoc -s -f markdown -o %s %s --data-dir ./" % (tmp_doc_path, md_full_path))
    word.Selection.InsertFile(FileName= tmp_doc_path)
    os.system("rm %s" % tmp_doc_path)
   

def table_style(TableStyle_, sec_):
     
    for tbl in doc.Sections(sec_).Range.Tables:
        tbl.Style = TableStyle_

def table_caption(sec_):

    for tbl in doc.Sections(sec_).Range.Tables:
        tbl.Select()
        if (not tbl.Title==''):
            word.Selection.GoToPrevious (wdGoToLine)
            word.Selection.Expand(wdLine)
    
            word.Selection.Delete()

            tbl.Range.InsertCaption(Label="Table", TitleAutoText="", Title=". " \
            + tbl.Title, Position=wdCaptionPositionAbove, ExcludeLabel=0)

def figure_caption(sec_):

    for fgr in doc.Sections(sec_).Range.InlineShapes:
        fgr.Select()
        if (not fgr.AlternativeText=='' and not fgr.AlternativeText=='---' and not fgr.AlternativeText=='***' and not fgr.AlternativeText=='___'):
            fgr.Range.InsertCaption(Label="Figure", TitleAutoText="", Title=". " \
            + fgr.AlternativeText, Position=wdCaptionPositionBelow, ExcludeLabel=0)
            word.Selection.GoToNext (wdGoToLine)
            word.Selection.GoToNext (wdGoToLine)
            word.Selection.Expand(wdParagraph)

            word.Selection.Delete()

def structure_verification(sec_, name_):

    doc.Sections(sec_).Range.Select()
    
    word.Selection.Range.GoTo(What=wdGoToHeading, Which=wdGoToFirst)
    word.Selection.HomeKey(wdLine)
    word.Selection.Expand(wdParagraph)
#    word.Selection.Range.Bookmarks("\line").Select
    title = word.Selection.Text
    title = title[:-1]
    if (title == ""):
        print ("  Please remove all blank lines in the beginning of section %s!" % sec_)
        quit()
    if (not title == name_):
        print ("\n  ###############################################")
        print ("\n  Section: %s, title_yaml: %s, title_docx: %s" % (sec_, name_, title))
        print ("  Correct dotx or YAML file!")
        quit()

def update_fields():
    doc.Fields.Update()

############### Main body ######################

## get file's names from the command line
yes_no = "no"
if (len(sys.argv)==2):
    structure_yaml_file = sys.argv[1]
else:
    print ("\n  ###############################################")
    print ("\n  By default YAML file describing the structure of your word document is\n  ./structure.yml. If you want to change it use tne next syntax:")
    print ("\n         python render.py <structure_yaml_file>\n")
    yes_no = input('  Continue with default path (y/n): ')
    if yes_no == 'y':
        structure_yaml_file = "structure.yml"
    else:
        quit()


## current folder:

dirpath_ = os.getcwd()
dirpath = dirpath_.replace("\\", "/")

## take data from YAML file 

my_config=''
f = open( "./%s" % structure_yaml_file )
data1 = f.read()
f.close()

yaml_version = yaml.__version__

if (float(yaml_version) < 5.1):
    yaml_data = yaml.load(data1)
else:
    yaml_data = yaml.load(data1,Loader=yaml.FullLoader)

yes_no = "no"
message = '\n  The next sections will be changed: \n'
for j in yaml_data["sections"]:
    if not j["action"]== "ignore":
        message = message + '\n  Section %s: %s' % (j["number"], j["name"])
print ("\n  ###############################################")
print (message)
yes_no = input('\n  Continue (y/n): ')
if yes_no == 'n':
    quit()
 
    

TableStyle = yaml_data["general"]["table_style"]
dest_word_file = yaml_data["general"]["dest_word_file"]

word = win32.Dispatch('Word.Application')
word.Visible = yaml_data["script"]["word"]["visible"]
doc = word.Documents.Open(dirpath + '/' + dest_word_file)


for j in yaml_data["sections"]:
    if not j["action"]== "ignore":
        sec = j["number"]
        name = j["name"]
        filepath = j["md_file"]
        if (yaml_data["script"]["actions"]["structure_verification"] == "yes"):
            structure_verification(sec, name)
        if (yaml_data["script"]["actions"]["change_data"] == "yes"):
            insert_data(sec, filepath)
        if (yaml_data["script"]["actions"]["table_style"] == "yes"):
            table_style(TableStyle, sec)
        if (yaml_data["script"]["actions"]["table_caption"] == "yes"):
            table_caption(sec)
        if (yaml_data["script"]["actions"]["figure_caption"] == "yes"):
            
            figure_caption(sec)

if (yaml_data["script"]["actions"]["update_fields"] == "yes"):
    update_fields()

if (yaml_data["script"]["word"]["save"] == "yes"):
    doc.Save()
if (yaml_data["script"]["word"]["close"] == "yes"):
    doc.Close()
