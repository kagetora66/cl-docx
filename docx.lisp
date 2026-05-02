(in-package #:cl-docx)

(defun repack-directory-to-docx (source-directory output-docx)
  "Compresses the SOURCE-DIRECTORY into a new DOCX file at OUTPUT-DOCX."
  (zip:zip output-docx source-directory :if-exists :supersede))
  
(defclass paragraph ()
  ((node :initarg :node
         :accessor node-reader
         :type dom:element
         :documentation "Abstraction layer for xml nodes representing paragraphs")
  (text :initarg :text
        :accessor text-acc
        :documentation "TEXT node of paragraph node")))

(defclass text ()
  ((text-value :initarg :text-value
              :accessor text-get
              :documentation "Representation of text object ")))

(defun wrap-paragraph-constructor (node-instance)
  (make-instance 'paragraph :node node-instance
                            :text (if (array-in-bounds-p (dom:get-elements-by-tag-name node-instance "w:t") 0)
                                      (aref (dom:get-elements-by-tag-name node-instance "w:t") 0)
                                      nil)))

(defun wrap-text-constructor (node-instance)
  (make-instance 'text :text-value node-instance))

(defun wrap-texts (node-vector)
  (map 'vector #'wrap-text-constructor node-vector))

(defun wrap-paragraphs (node-vector)
  (map 'vector #'wrap-paragraph-constructor node-vector))


(defun get-all-paragraphs (treenode)
  "Returns a VECTOR of PARAGRAPH objects"
  (wrap-paragraphs (dom:get-elements-by-tag-name treenode "w:p")))

(defun get-all-texts (treenode)
  "Returns a VECTOR of TEXT objects"
  (wrap-texts (dom:get-elements-by-tag-name treenode "w:t")))

(defmethod read-value ((text paragraph))
  "Read the TEXT value of a PARAGRAPH object"
  (unless (null (text-acc text))
    (dom:node-value (dom:last-child (text-acc text)))))

(defmethod read-value ((text text))
  "Read the TEXT value of a TEXT object"
  (unless (null (text-get text))
    (dom:node-value (dom:last-child (text-get text)))))

(defmethod write-value ((text paragraph) new)
  "Replaces the TEXT paragraph with NEW string"
  (setf (dom:node-value (text-acc text)) new))

(defmethod write-value ((text text) new)
  "Replaces the TEXT Object with NEW string"
  (setf (dom:node-value (text-get text)) new))

(defun get-xml-tree (doc-path)
  "Retruns TREENODE object from temporary DOC_PATH"
  (cxml:parse-file (merge-pathnames "word/document.xml" doc-path) (cxml-dom:make-dom-builder)))

;;We're doing this due to a bug in the zip library
(defun unzip (pathname)
  "Uses operating system unzip funtions to extract docx file to temporary folder. Retunrs the extracted path"
  (let ((pathname (if (pathnamep pathname) pathname (pathname pathname)))
        (output-dir (ensure-directories-exist
                     (uiop:ensure-directory-pathname
                      (merge-pathnames
                       (uiop:temporary-directory) (format nil "temp_~a" (pathname-name pathname)))))))
    #+linux (uiop:run-program
             (list "unzip" "-o" (namestring pathname) "-d" (namestring output-dir)))
    #+windows (uiop:run-program
               (list "powershell" "-Command"
                  (format nil "Expand-Archive -Path '~a' -DestinationPath '~a' -Force"
                    (namestring pathname) (namestring output-dir))))
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

(defmethod remove-paragraph ((para paragraph))
  "Removes PARAGRAPH object from docx tree (experimental)"
  ;;NOTE we're assuming the immediate parent of w:p is always w:body. Needs more testing
  (dom:remove-child (dom:parent-node (node-reader para)) (node-reader para)))

#+test
(time (with-open-docx (doc #P"./TEST2.docx")
        (setf paras (get-all-texts doc))
        (setf doctree doc)
        (map 'vector #'read-value paras)))
        
