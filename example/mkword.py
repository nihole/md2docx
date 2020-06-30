import os
import sys

render_path = '../scripts/render.py'

######### Main Body ######################


######### get file's names from the command line ####################
if (len(sys.argv)==2):
    structure_path = sys.argv[1]
else:
    print ("   ######################################################\n")
    print ("   Syntax is:\n")
    print ("   python3 mkword.py structure/structure.yml\n")
    print ("   Change paths to render.py in mkword.py if needed.\n")
    print ("   ######################################################\n")
    quit()

cmd = 'python %s %s' % (render_path, structure_path)

returned_value = os.system(cmd)

