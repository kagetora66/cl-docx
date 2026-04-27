(in-package #:cl-docx)

(defun repack-directory-to-docx (source-directory output-docx)
  "Compresses the SOURCE-DIRECTORY into a new DOCX file at OUTPUT-DOCX."
  (zip:zip output-docx source-directory :if-exists :supersede))
  
(defclass paragraph ()
  ((node :initarg :node
         :accessor node-reader
         :type dom:element
         :documentation "Abstraction layer for xml nodes representing paragraphs")))

(declaim (inline wrap-paragraph))

(defun wrap-paragraph (node-instance)
  (make-instance 'paragraph :node node-instance))

(defun wrap-paragraphs (node-vector)
  (map 'vector #'wrap-paragraph node-vector))


(defun get-all-paragraphs (treenode)
  "Returns a VECTOR of PARAGRAPH objects"
  (wrap-paragraphs (map 'vector #'dom:last-child (dom:get-elements-by-tag-name treenode "w:t"))))


(defmethod read-value ((text paragraph))
  "Read the value of paragraph"
  (dom:node-value (node-reader text)))

(defmethod write-value ((text paragraph) new)
  "Replaces the paragraph with NEW string"
  (setf (dom:node-value (node-reader text)) new))

(defun get-xml-tree (doc-path)
  "Retruns TREENODE object from temporary DOC_PATH"
  (cxml:parse-file (merge-pathnames "word/document.xml" doc-path) (cxml-dom:make-dom-builder)))

(defun unzip (pathname)
  "Uses operating system unzio funtions to extract docx file to temporary folder. Retunrs the extracted path"
  (let ((output-dir (ensure-directories-exist
                     (uiop:ensure-directory-pathname
                      (merge-pathnames
                       (uiop:temporary-directory) (format nil "temp_~a" (pathname-name pathname)))))))
    #+linux (uiop:launch-program
             (list "unzip" "-o" (namestring pathname) "-d" (namestring output-dir)))
    output-dir))

(defun repackage (doc-path treenode original-doc)
  "Repackage the contents of directory hosting DOCPATH of temporary after converting TREENODE to document.xml and replacing it"
  (with-open-file (stream (merge-pathnames "word/document.xml" doc-path) :direction :output
                                                                         :if-exists :supersede
                                                                         :element-type '(unsigned-byte 8))
    (dom:map-document (cxml:make-octet-stream-sink stream) treenode)
    :include-doctype t :include-xmlns-attributes t
    :include-namespace-uri t :include-default-values t)
  (repack-directory-to-docx doc-path original-doc))


(defmacro with-open-docx ((docvar docpath) &body body)
  "DOCVAR is an xml treenode representing the content of document.xml. Changes are saved to the DOCPATH file at the end of macro"
  (let ((temp-path (gensym)))
    `(let* ((,temp-path (unzip ,docpath))
            (,docvar (get-xml-tree ,temp-path)))
       (unwind-protect
            (progn ,@body)
         (repackage ,temp-path ,docvar ,docpath)))))


#+test
(time (with-open-docx (doc #P"/home/safari/Documents/lisp/cl-docx/gf.docx")
  (setf paras (get-all-paragraphs doc))
        (map 'vector #'read-value paras)))
