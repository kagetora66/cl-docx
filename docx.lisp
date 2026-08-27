(in-package #:cl-docx)

(defun repack-directory-to-docx (source-directory output-docx)
  "Compresses the SOURCE-DIRECTORY into a new DOCX file at OUTPUT-DOCX."
  (zip:zip output-docx source-directory :if-exists :supersede)
  ;;validate for the purpose of preventing wildcard tomfoolery
  (uiop:delete-directory-tree source-directory :validate #'uiop:directory-pathname-p))

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

(defclass table ()
  ((row-list :initarg :row-list :accessor row-list)))

(defclass table-row ()
  ((cell-list :initarg :cell-list :accessor cell-list)
   (row-dom :initarg :row-dom
            :accessor row-dom
            :documentation "DOM object representing a row node in xml"))
  (:documentation "Abstraction for a row in a table. Includes a list of CELL object and a tree node for the row"))

(defclass cell ()
  ((cell :initarg :cell :accessor cell)))


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

(defmethod read-value ((cl cell))
  "READS the text value inside a CELL object"
  (if (array-in-bounds-p (dom:get-elements-by-tag-name (cell cl) "w:t") 0)
      (dom:node-value (dom:last-child
                       (aref (dom:get-elements-by-tag-name (cell cl) "w:t") 0)))
                                      nil))

(defmethod read-value ((row table-row))
  "Reads and returns a list of cell values for a TABLE-ROW object"
  (mapcar #'read-value (cell-list row)))

(defmethod read-value ((table table))
  "Reads and returns a list of rows (each row being a list of cell values) for a TABLE object"
  (mapcar #'read-value (row-list table)))

(defmethod write-value ((text paragraph) new)
  "Replaces the TEXT paragraph with NEW string"
  (setf (dom:node-value (dom:last-child (text-acc text))) new))

(defmethod write-value ((text text) new)
  "Replaces the TEXT Object with NEW string"
  (setf (dom:node-value (dom:last-child (text-get text))) new))

(defmethod write-value ((text cell) new)
  "Replaces the text of CELL Object with NEW string"
  (if (array-in-bounds-p (dom:get-elements-by-tag-name (cell text) "w:t") 0)
      (setf (dom:node-value (dom:last-child
                             (aref (dom:get-elements-by-tag-name (cell text) "w:t") 0)))
            new)
      ;;IF CELL HAS NO TEXT, ADD CHILDREN ACCORDINGLY
      (let* ((cell-node (cell text))
             (document (dom:owner-document cell-node))
             (text-node (dom:create-text-node document new))
             (w-t-element (dom:create-element document "w:t")))
        (dom:append-child w-t-element text-node)
        (dom:append-child (aref (dom:get-elements-by-tag-name cell-node "w:r") 0)
                          w-t-element))))

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
    (dom:map-document (cxml:make-octet-stream-sink stream) treenode))
  (repack-directory-to-docx doc-path original-doc))


(defmacro with-open-docx ((docvar docpath) &body body)
  "DOCVAR is an xml treenode representing the content of document.xml. Changes are saved to the DOCPATH file at the end of macro"
  (let ((temp-path (gensym)))
    `(let* ((,temp-path (unzip ,docpath))
            (,docvar (get-xml-tree ,temp-path)))
       (unwind-protect
            (progn ,@body)
         (repackage ,temp-path ,docvar ,docpath)))))

(defmethod remove-item ((para paragraph))
  "Removes PARAGRAPH object from docx tree (experimental)"
  ;;NOTE we're assuming the immediate parent of w:p is always w:body. Needs more testing
  (dom:remove-child (dom:parent-node (node-reader para)) (node-reader para)))

(defmethod remove-item ((row-node table-row))
  "Removes ROW object from docx tree"
  (dom:remove-child (dom:parent-node (row-dom row-node)) (row-dom row-node)))


(defun wrap-cell-constructor (row-node)
  "Gets a ROW xml node and generates a list of CELL objects from it"
  (map 'list (lambda (cl) (make-instance 'cell :cell cl)) (dom:get-elements-by-tag-name row-node "w:tc")))

(defun row-constrcutor (table-node)
  "Constructs TABLE-ROWS containing CELL objects from TABLE NODE"
  (map 'list
       (lambda (row) (make-instance 'table-row :cell-list (wrap-cell-constructor row)
                                    :row-dom row))
            (dom:get-elements-by-tag-name table-node "w:tr")))

(defun get-all-tables (node-instance)
  "Constuct a list of TABLE objects from document node, NODE-INSTANCE"
  (map 'list (lambda (table) (make-instance 'table :row-list (row-constrcutor table)))
       (dom:get-elements-by-tag-name node-instance "w:tbl")))

(defun get-rows (table-instance)
  "Returns a list of all row objects"
  (row-list table-instance))

(defun get-cells (row-instance)
  "Get all CELL objects inside a ROW"
  (cell-list row-instance))

(defmethod add-row-copy ((table table))
  "Adds a row at to the table by copying the first row"
  (let ((first-row (car (get-rows table))))
    (dom:append-child
     (dom:parent-node
      (row-dom first-row))
     (dom:clone-node (row-dom first-row) t))))

;;for reading paragraphs
#+test
(time (with-open-docx (doc #P"./test.docx")
        (setf paras (get-all-texts doc))
        (setf treenode doc)
        (map 'vector #'read-value paras)
        ))

;;Changing the value of a cell
#+test
(with-open-docx (doc #P"./test.docx")
  (let* ((tables (get-all-tables doc))
         (first-row (car (get-rows (car tables))))
         (first-cell (cadr (get-cells first-row))))
    (write-value first-cell "TESTCELL")

    (dotimes (i 10) ;;ADD TEN ROWS
      (add-row-copy (car tables)))
    (read-value (car tables))
    ))
