(asdf:defsystem #:cl-docx
  :name "CL-DOCX"
  :description "Simple Reader and Writer library for Docx files in Common Lisp"
  :author "kagetora66 <roninofthepersia@gmail.com>"
  :license "MIT"
  :serial t
  :components ((:file "package")
               (:file "docx"))
  :depends-on (:zip :flexi-streams :cxml))

