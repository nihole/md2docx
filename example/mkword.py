import os
import sys

render_path = '../scripts/render.py'

structure_path  = './structure.yml'

######### Main Body ######################


######### get file's names from the command line ####################
if not (len(sys.argv)==1):
    print ("   ######################################################\n")
    print ("   Syntax is:\n")
    print ("   python3 mkword.py\n")
    print ("   Change paths in mkword.py if needed.\n")
    print ("   ######################################################\n")
    quit()

cmd = 'python %s %s' % (render_path, structure_path)

returned_value = os.system(cmd)

