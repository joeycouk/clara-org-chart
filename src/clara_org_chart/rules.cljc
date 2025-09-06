(ns clara-org-chart.rules
  (:require
   [clara-org-chart.util :as util]
   [clara-org-chart.xlsx :as xlsx]
   [clara-org-chart.eav :as eav]
   [clara-org-chart.position :as pos]
   #?@(:clj [[clara.rules :as rules]
             [clara-eav.rules :as er]
             [clara.rules.accumulators :as accum]
             ;;  [portal.api :as p] ;; must comment out before building cljs
             ]
      )))
;; 

;; identify column of values, always located in the first row of the sheet
(defrecord ColumnValue [
                         sheet-name
                         column
                         value
])


(er/defrule extract-column-headers
  "Extract column headers from the first row of each sheet" 
  [[?e1 :cell/row 0]] 
  [[?e1 :cell/sheet ?sheet]]
  [[?e1 :cell/value ?value]]
  [[?e1 :cell/col ?col]]

  =>
  (rules/insert! (->ColumnValue ?sheet ?col ?value)))


(rules/defquery get-column-values
  "Query to get all column values"
  []
  [?column-values <- (accum/all) :from [ColumnValue]])


;; "Position"
;; "Correct"
;; "Location"
;; "Region"
;; "TB_Adjustment"
;; "Time_Base"
;; "Unique_Flag"
;; "SerialNumber"
;; "ClassCode"
;; "UnitCode"
;; "AgencyCode"
;; "Dotted_Reports"
;; "Comments/Notes"
;; "City"
;; "Position Reports To Position (Dotted Line)"
;; "Position Reports To Position"
;; "Current Employee"
;; "Title"


(comment
  
  (def eavs (eav/xlsx->eav (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx") :version :v1))
  ;; Clara-EAV expects raw EAV records, not transformed vectors

  ;; Session is now defined above, outside the comment block  ;; INCORRECT: Explicit rule vectors don't work properly in Clara-EAV
  (er/defsession test-session 'clara-org-chart.rules)


  (def results (-> test-session
                   (er/upsert eavs)
                   (rules/fire-rules)))
  
  (tap> (rules/query results get-column-values))

  (tap> (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx"))

  (tap> (pos/extract-positions (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx")))

  (tap> results)
  (tap> (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx"))
  :rcf)