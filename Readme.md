# CL-DOCX

# Introduction

CL-DOCX is a simple reader and writer for docx files written in Common Lisp. It currently works on operating on paragraphs but it's made in Object Oriented style in hope of future extensions. 

The example way of using the library is as such:

``` common-lisp
(with-open-docx (doc "./test.docx")
  (setf paras (get-all-texts doc))
    (write-value (aref paras 0) "hi"))
```

The primary way to engage with this library is the macro **with-open-docx** , The **doc** variable will contain the xml tree node of the document.xml file in the lisp structure format. 

To extract all paragraphs from a document.xml file, the function **get-all-paragraphs** is used. This function returns a VECTOR of PARAGRAPH objects. On these objects, you can use the Generic Functions **read-value** to read the value of each paragraph (string) and **write-value** which gets an PARAGRAPH object, a replacement string and modifies it. The function **get-all-texts** gets every TEXT item inside the word file similarly. 

After the end of **with-open-docx** , all changes to the contents are saved to the docx file. 

Here's another example of reading all paragraph from a docx file: 

`(map 'vector #'read-value (get-all-texts doc))`

You can also get all the tables from a docx file in the same manner. Tables are organzied in this manner:

    1: A table object contains list of row object
    
    2: A row object contains a list of cell objects
    
    
The function **get-all-tables** returns a list of available tables inside the docx files. You can extract rows from a table and cells from a row object with functions **get-rows** and **get-cells**. You can use the generic function **read-values** to get all entries of a table as a list of lists like (This is a 3X3 table containing only two entries):

`(("table2" "table2" NIL) (NIL NIL NIL) (NIL NIL NIL)) `

To edit the value of an existing cell, we can use the same **write-value** method. Here's an example of changing the value of one cell to a new value:


``` common-lisp
(with-open-docx (doc #P"./test.docx")
  (let* ((tables (get-all-tables doc)) ;;LIST of tables in docx
         (first-row (car (get-rows (car tables)))) ;; A LIST of CELL Objects
         (first-cell (car (get-cells first-row))))
    (write-value first-cell "NEWCELL")))
```


# Features Implemented 

  * [x] Reading texts from a docx file
 
  * [x] Editing the texts from a docx file by writing on the text objects

  * [ ] Writing Lisp data inside a docx file
  
  * [ ] Tables operations:
    * [x] Reading Tables/Rows/Cells
    * [x] Editing Tables/Rows/Cells
    * [ ] Add/Remove Tables/Rows/Cells
 
  
# How to Install

1: Clone the repository to your quicklisp's local-project folder 

2: Use this to load the contents:

`(ql:quickload "cl-docx")`
