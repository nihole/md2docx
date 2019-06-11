# md2docx
markdown -> docx word 

The Word format is the de facto standard for documentation in many companies. It is a powerful tool with many features and settings. And everything looks perfect if you don’t need to make many changes by many employees. 

A wiki is usually used for this (for example, Confluence). But in this case, you will have problems with exporting your wiki pages to a Word document with the expected styles, Word objects ... I think it is possible, but for me using GIT (instead of a wiki) looks more attractive.

## Idea

We have to use something else instead of Word. I think __MarkDown__ is a good choice. This gives us enough feature set and text format. Therefore, we can use __GIT__ (for example, GitHub) to manage all changes. 

Now, if we have all the documentation created in MarkDown, we can __convert them to Word__ with corporate styles and requirements

## High Level Workflow

- Only MarkDown files are modified manually. Instead of managing one large file, it’s easy to manage several smaller ones
- One of the first operations performed by render.py is to change the MarkDown files to docx format. This Word files have no styles, fields.. They are temporary and render.py removes them at the end
- Using the structure.yml data and temporary docx files (created earlier) as input, render.py adds all the necessary Word objects (for example, fields), implement the styles and paste the converted data into the correct places in the resulting Word document

![render.py workflow:](https://github.com/nihole/md2docx/blob/master/media/md2word_work_flow.png)

## Installation
- clone this project into your local folder
- install Python3 with YAML and pywin32 packages (win32com.client)
- install Pandoc
- install any MarkDown editor. I use Typora. Any simple text editor can be used, but this is not very convenient

## Files
- <a href="https://github.com/nihole/md2docx/blob/master/render.py">render.py</a> - Python script. Takes data from the markdown file, converts and inserts it into a Word document
- <a href="https://github.com/nihole/md2docx/blob/master/structure.yml">structure.yml</a> - YAML file with all the necessary parameters for render.py
- template.docx is a Word document that we are going to fill out with the context of markdown documents. This Word document should have all the styles, templates ... you are going to use. This document is divided manually into sections. Each section corresponds to one chapter
- <a href="https://github.com/nihole/md2docx/blob/master/example_chapter.md">example_chapter.md</a> - markdown document with example information
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
- figure_caption - transforms markdown caption to the word caption for figures
- update_fields - updates all fields


