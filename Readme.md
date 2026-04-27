# CL-DOCX

# Introduction

CL-DOCX is a simple reader and writer for docx files written in Common Lisp. It currently works on operating on paragraphs but it's made in Object Oriented style in hope of future extensions. 

The example way of using the library is as such:

``` common-lisp
(with-open-docx (doc "./test.docx")
  (setf paras (get-all-paragraphs doc))
    (write-value (aref paras 0) "hi"))
```

The primary way to engage with this library is the macro **with-open-docx** , The **doc** variable will contain the xml tree node of the document.xml file in the lisp structure format. 

To extract all paragraphs from a document.xml file, the function **get-all-paragraphs** is used. This function returns a VECTOR of PARAGRAPH objects. On these objects, you can use the Generic Functions **read-value** to read the value of each paragraph (string) and **write-value** which gets an PARAGRAPH object, a replacement string and modifies it.

After the end of **with-open-docx** , all changes to the contents are saved to the docx file. 

Here's another example of reading all paragraph from a docx file: 

`(map 'vector #'read-value (get-all-paragraph doc)')`

# How to Install

1: Clone the repository to your quicklisp's local-project folder 

2: Use this to load the contents:

`(ql:quickload "cl-docx")`
