# md2docx
markdown -> docx word 

The Word format is the de facto standard for documentation in many companies. It is a powerful tool with many features and settings. And everything looks perfect if you don’t need to make a lot of changes by many employees and keep track of all these changes, discussions and approvals.

A wiki is usually used for internal documentation (for example, Confluence).But if you need to provide your documentation to the client and you need you need to solve some problems of exporting wiki pages to a Word document corporate styles. I think it is possible, but for me using GIT (instead of a wiki) looks more attractive.

## Idea

With GIT we have to use something else instead of Word. I think __MarkDown__ is a good choice. It is powerful enough and allows us to use GIT (for example, GitHub) to manage all changes.

Now, if we have all the documentation created in MarkDown, we can __convert them to Word__ with corporate styles and requirements.

## Possible or not?

We have Pandoc to convert MarkDown to Word. But, of course, we can not transform everithing, and we have a lot of things in Word which are missing in MarkDown. And I think that it is impossible or very difficult to solve this task in general for all cases.

But the fact is that we don’t need to solve this problem all cases. We only need to solve this for our particular one.

For example, in my case I only need

- add styles for all tables (in accordance with corporate requirements)
- insert automatic table captions
- insert automatic image captions

and I will have exactly the same word file that I had when I created it manually (in the old style).

## Instruments

- Markdown with some editor (I use Typora) - documents creation
- GIT (GitHub for example) - version control system
- Pandoc - MarkDown -> Wod (docx) conversion
- YAML - describes the strucuture of the final Word document and for other input parameters for python script (render.py)
- Python with pywin32 - pywin32 permits to change Word file
- MS Word (docx) - resulting document

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


