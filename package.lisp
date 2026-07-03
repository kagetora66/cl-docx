(defpackage #:cl-docx
  (:use #:cl)
  (:export #:get-all-paragraphs
           #:with-open-docx
           #:read-value
           #:write-value
           #:remove-item
           #:get-all-texts
           #:get-all-tables
           #:get-rows
           #:get-cells))
