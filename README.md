# md2docx
markdown -> docx word 

## Installation
- clone this project into your local folder
- install Python3 with YAML and pywin32 packages (win32com.client)
- install Pandok
- install any MarkDown editor. I use Typora. Any simple text editor can be used, but this is not very convenient

## Files
- <a href="https://github.com/nihole/md2docx/blob/master/render.py">render.py</a> - Python script. Takes data from the markdown file and inserts it into a Word document
- <a href="https://github.com/nihole/md2docx/blob/master/structure.yml">structure.yml</a> - YAML file with all the necessary parameters for render.py
- template.docx is a Word document that we are going to fill out with the context of markdown documents. This Word document should have all the styles, templates ... you are going to use. This document is divided manually into sections. Each section corresponds to one chapter
- <a href="https://github.com/nihole/md2docx/blob/master/example_chapter.md">example_chapter.md</a> - markdown document with F5 information
- <a href="https://github.com/nihole/md2docx/tree/master/initial">initial folder</a> - initial files

## How to start
- python3 render.py
- yes
- yes

## render.py

From <a href="https://github.com/nihole/md2docx/blob/master/structure.yml">structure.yml</a> you can see how this script works.


actions:  
&nbsp;&nbsp; structure_verification: "yes"  
&nbsp;&nbsp; change_data: "yes" 
&nbsp;&nbsp; table_style: "yes"  
&nbsp;&nbsp; table_caption: "yes"  
&nbsp;&nbsp; figure_caption: "yes"  
&nbsp;&nbsp; update_fields: "yes"  
    
- structure_verification - verifies if the section corresponds the title
- change_data - inserts data in this section
- table_style - changes the table style
- table_caption - transforms markdown caption to the word caption for tables
- figure_caption - - transforms markdown caption to the word caption for figures
- update_fields - updates all fields


 

