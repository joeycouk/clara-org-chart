(ns clara-org-chart.position-rules
  (:require
   [clara-org-chart.xlsx :as xlsx]
   [clara.rules :refer [defrule defquery insert! insert-all! insert-unconditional! defsession insert-all fire-rules query]]
   [clara.rules.accumulators :as accum]
   [clara-org-chart.data-dictionary-extractor])
  (:import (data_types Position)))


(defrecord CalculatedPosition
           [row-num
            position
            title
            current-employee
            reports-to-position
            dotted-line-reports-to-position
            city
            comments-notes
            dotted-reports
            agencycode
            unitcode
            classcode
            serialnumber
            unique-flag
            time-base
            tenure
            region 
            position-weight 
            total-subordinates
            ])

(defrule calculate-weight-for-ft-total-suborinate-count
  "Apply rules to calculate the position weight based on time-base and total subordinates" 
  [?pos <- Position (= "FT" time-base)]
  =>
   (insert! (map->CalculatedPosition
            (assoc ?pos :position-weight 1.0))))


(defrule calculate-weight-for-pt-total-suborinate-count
  "Apply rules to calculate the position weight based on time-base and total subordinates" 
  [?pos <- Position (= "PT" time-base)]
  =>
  (insert! (map->CalculatedPosition
            (assoc ?pos :position-weight 0.5))
  
  ))

;; can be a string or a double already, return true if can be parsed as double
(defn test-time-base-can-be-double [time-base-str]
  (try
    ;; first test to see if it is already a double
    (if (double? time-base-str)
      true
      ;; else try to parse as double
      (if (string? time-base-str)
        (not (Double/isNaN (Double/parseDouble time-base-str)))
        false))
    (catch Exception e
      false)))

(defrule calculate-weight-for-hard-coded-numerical-value
  "If there's already a numerical value in the Time-Base field, use that as the position weight" 
  [?pos <- Position (or (double? time-base) (not (nil? time-base)))]
  ;; test the time-base is not an empty string or nil
  ;; any number can match, not just doubles

  ;; only match if it's a character sequence that represents a number
  [:test (test-time-base-can-be-double (:time-base ?pos))] 
  =>
  (insert! (map->CalculatedPosition
            (assoc ?pos :position-weight (:time-base ?pos)))
           )
  
  )

(defrule calculate-weight-for-int-total-suborinate-count
  "Apply rules to calculate the position weight based on time-base and total subordinates" 
  [?pos <- Position (= "INT" time-base)]
  =>
  (insert! (map->CalculatedPosition
            (assoc ?pos :position-weight 0))
  
  ))


(defrule calculate-weight-for-blank-time-base-total-suborinate-count
  "Apply rules to calculate the position weight when the time-base is blank for total subordinates" 
  [?pos <- Position (or (= "" time-base) (nil? time-base) )]
  =>
  ;; if class code like 9XX then 0 else 1
  (let [weight (if (and (string? (:classcode ?pos))
                        (re-matches #"9\d{2}" (:classcode ?pos)))
                 0.0
                 1.0)]
  (insert! (map->CalculatedPosition
            (assoc ?pos :position-weight weight))
  
  )) 
  )



(defquery get-all-extracted-positions
  "Query to get all the extracted positions with calculated weights"
  []
  [?calculatedPositions <- (accum/all) :from [CalculatedPosition]])



;; Define the session at the top level (compile time)
(defsession pos-rules-sess 'clara-org-chart.position-rules)

;; a function that will fire the rules and return a sequence of ExtractedPosition records
(defn process-positions
  "Process a sequence of Position records and return ExtractedPosition records with calculated weights"
  [positions]
  (let [session (-> pos-rules-sess
                    (insert-all positions)
                    (fire-rules))
        results (query session get-all-extracted-positions)]
    (:?calculatedPositions (first results))))




(comment

  
  (tap> 
   (process-positions (pos/extract-positions (xlsx/extract-data "resources/OrgChart_HQ03.xlsx"))))




  ;; Use the session defined above
  (def results-streaming (-> pos-rules-sess
                             (insert-all
                              (concat
                               (pos/extract-positions (xlsx/extract-data "resources/OrgChart_HQ03.xlsx" :streaming true))))
                             (fire-rules)))


  (tap> (pos/extract-positions (xlsx/extract-data "resources/OrgChart_HQ03.xlsx" :streaming true)))



  (tap> (:?calculatedPositions (first (query results-streaming get-all-extracted-positions))))

  ;; filter out all results where position-weight is less than 1
  (def rcf (filter #(< (:position-weight %) 1.0)
                   (:?calculatedPositions (first (query results-streaming get-all-extracted-positions)))))

  (tap> rcf)




  :rcf)