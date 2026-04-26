(in-package #:cl-docx)

(defun repack-directory-to-docx (source-directory output-docx)
  "Compresses the SOURCE-DIRECTORY into a new DOCX file at OUTPUT-DOCX."
  (zip:zip output-docx source-directory :if-exists :supersede))
  
(defclass paragraph ()
  ((node :initarg :node
         :accessor node-reader
         :type xmls:node
         :documentation "Abstraction layer for xml nodes representing paragraphs")))

(declaim (inline wrap-paragraph))

(defun wrap-paragraph (node-instance)
  (make-instance 'paragraph :node node-instance))

(defun wrap-paragraphs (node-list)
  (mapcar #'wrap-paragraph node-list))

(defun get-all-paragraphs (treenode)
  "Returns a list of PARAGRAPH objects"
  (wrap-paragraphs (remove-if #'null (mapcar (lambda (node)
                             (xmls:xmlrep-find-child-tag :t node nil))
                           (mapcar (lambda (node)
                                     (xmls:xmlrep-find-child-tag :r node nil))
                                   (xmls:xmlrep-find-child-tags :p
                                                                (xmls:xmlrep-find-child-tag :body treenode)))))))
(defmethod read-value ((text paragraph))
  "Read the value of paragraph"
  (xmls:xmlrep-string-child (node-reader text) nil))

(defmethod write-value ((text paragraph) new)
  "Replaces the paragraph with NEW string"
  (setf (xmls:node-children (node-reader text)) (list new)))

(defun get-xml-tree (doc-path)
  "Retruns TREENODE object from temporary DOC_PATH"
  (let ((xml-content (uiop:read-file-string (merge-pathnames "word/document.xml" doc-path))))
    (xmls:parse xml-content)))

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
  (let ((output-xml (xmls:toxml treenode)))
    (with-open-file (stream (merge-pathnames "word/document.xml" doc-path) :if-exists :overwrite :direction :output)
      (write-line output-xml stream)))
  (repack-directory-to-docx doc-path original-doc))

;;EXPERIMENTAL MAY CORRUPT
;;FIXME Probably an error in xml library
(defmacro with-open-docx ((docvar docpath) &body body)
  "DOCVAR is an xml treenode representing the content of document.xml. Changes are saved to the DOCPATH file at the end of macro"
  (let ((temp-path (gensym)))
    `(let* ((,temp-path (unzip ,docpath))
            (,docvar (get-xml-tree ,temp-path)))
       (unwind-protect
            (progn ,@body)
         (repackage ,temp-path ,docvar ,docpath)))))

(defun get-paragraphs (pathname)
  "Get all paragraph object from docx PATHNAME"
  (let ((temp-path (unzip pathname)))
    (get-all-paragraphs (get-xml-tree temp-path))))


#+test
(time (with-open-docx (doc "./test.docx")
  (setf paras (get-all-paragraphs doc))
    (write-value (first paras) "hi")))
