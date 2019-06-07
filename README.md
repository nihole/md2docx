# md2docx
markdown -> docx word with cisco templates

## Installation
- clone this project to your local folder
- install Python3 with YAML and pywin32 packages (win32com.client)
- install pandok
install some markdown editor. I use Typora. Any simple text editor can be used, but this is not very convenient.

## Files
- render.py - Python script. Takes data from the markdown file and inserts it into a Word document.
- structure.yml - YAML file with all the necessary parameters for render.py.
- template.docx is a Word document that we are going to fill with the context of markdown documents. This Word document should have all the styles, patterns ... used by Cisco for LLD. This document is divided manually into sections. Each section corresponds to one section.
- example_chapter.md - markdown document with F5 information. We are going to insert this data in the right place with the correct styles in template.docx
