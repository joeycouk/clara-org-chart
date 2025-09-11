(ns clara-org-chart.rules
  (:require
   [clara-org-chart.xlsx :as xlsx]
   [org-tangle :as tangle]
   [clara-org-chart.position :as pos]
   [pdfBoxing :as pdf]
   [clara.rules :as rules]
   [clara.rules.accumulators :as accum]
   [clara-org-chart.org-chart-extractor :as extractor]
   )
  (:import (clara_org_chart.position Position)
           (clara_org_chart.org_chart_extractor 
                   OrgChartPageResult
                )))



;; Org chart page results have a list of postiions.  This isn't ideal for writing easy to understand rules.  Rather than a list i prefer a single object with a position value
(defrecord OrgChartPosition
           [
             position
             file-name
             page
             ])
(defrecord MatchingPosition
           [ 
             row-num
             position 
             reports-to-position
             file-name
             page
             ]
  )






(defrecord OrgChartError 
           [
             page
             missing-position-code
             extra-position-code
             duplicate-position-code
             path
             file-name
             description
           ])

(defrecord PositionWarning
           [position 
             description])

(defrecord ExtractedPosition
           [
            row-num 
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
            tb-adjustment
            region
            direct-subordinates 
            total-subordinates])



(defn position-code?
  "Return true if the value looks like a position code (###-###-####-###), false otherwise.
  Returns false for names, empty strings, or nil values."
  [value]
  (and (string? value)
       (not (empty? value))
       (re-matches #"^\d{3}-\d{3}-\d{4}-\d{3}$" value)))




(rules/defrule calculate-total-subordinates
  "Recursively calculate total subordinates for each position"
  [?pos <- Position (= ?posNum position)]
  [?subs <- (accum/count) :from [Position (= ?posNum reports-to-position)]]
  ;; [?subSubs <- (accum/sum :total-subordinates) :from [ExtractedPosition (= ?posNum reports-to-position)]]
  =>

  ;;  (println "Calculating total subordinates for position:" ?pos (get ?pos :position))
   (rules/insert! (->ExtractedPosition
                   (get ?pos :row-num)
                   (get ?pos :position)
                   (get ?pos :title)
                   (get ?pos :current-employee)
                   (get ?pos :reports-to-position)
                   (get ?pos :dotted-line-reports-to-position)
                   (get ?pos :city)
                   (get ?pos :comments-notes)
                   (get ?pos :dotted-reports)
                   (get ?pos :agencycode)
                   (get ?pos :unitcode)
                   (get ?pos :classcode)
                   (get ?pos :serialnumber)
                   (get ?pos :unique-flag)
                   (get ?pos :time-base)
                   (get ?pos :tb-adjustment)
                   (get ?pos :region)
                   (or ?subs 0)
                   0
                  ;;  (+ ?subs (or ?subSubs 0))
                   ))
   )


(rules/defrule break-org-chart-pages-into-positions
  "Break org chart pages into individual positions"
  [OrgChartPageResult (= ?page page) (= ?positions positions) (= ?fileName file-name)]
  =>
  (doseq [pos ?positions]
    (rules/insert! (->OrgChartPosition pos ?fileName ?page))
    ) 
  )

;; TODO rename this thing
(rules/defrule create-matching-rules
  "Where the position numbers match perfectly, generate an instance of MatchingPosition"
  [OrgChartPosition (= ?page page) (= ?position position) (= ?fileName file-name)]
  [ExtractedPosition (= position ?position) (= ?reportsToPos reports-to-position) (= ?rowNum row-num) ]
  =>
  (rules/insert! (->MatchingPosition
                  ?rowNum
                  ?position
                  ?reportsToPos
                  ?fileName
                  ?page
                  ))
  
  )


;; TODO rename this rule
(rules/defrule detect-org-chart-position-mismatches
  "Detect positions where the org chart pdf specifies a position number which does not exist within the xlsx document"
  [OrgChartPosition (= ?page page) (= ?position position) (= ?name file-name)]
  [:not [MatchingPosition (= ?page page) (= ?name file-name) (= ?position position)]]
  [OrgChartPageResult (= ?page page) (= ?path file-path) (= ?name file-name) (= ?description description)]
  =>
  (rules/insert! (->OrgChartError
                  ?page
                  ?position
                  nil ;; missing position
                  nil ;; duplicate position
                  ?path
                  ?name
                  "Orgchart specified in the PDF that does not exist in the xlsx document"
                  ))
  )


(rules/defrule detect-org-chart-position-duplicates
  "Detect when multiple positions are mapped to an org chart with the same position number"
  [?matchingRows <- (accum/distinct :row-num) :from [MatchingPosition (= ?page page) (= ?name file-name) (= ?position position)]]
  [:test (> (count ?matchingRows) 1)]
  [OrgChartPageResult (= ?page page) (= ?path file-path) (= ?name file-name) (= ?description description)]
  =>
  (rules/insert! (->OrgChartError
                  ?page
                  nil ;; missing-position-code
                  nil ;;   extra-position-code
                  ?position ;;duplicate position
                  ?path
                  ?name
                  (str "Orgchart has multiple matched positions with the same number " ?matchingRows))))


(rules/defrule detect-org-chart-missing-position
  "Detect position contained within the xlsx file but not within the org chart pdf extraction"
  [OrgChartPageResult 
   (= ?page page) 
   (= ?path file-path) 
   (= ?name file-name) 
   (= ?description description) 
   (= ?positions positions) ]
   [MatchingPosition (= ?page page) (= ?name file-name) (= ?position position) (= ?reportsToPosition reports-to-position) ]
  ;; where the reports to position doesn't exist in the pdf
  [:not [OrgChartPosition (= ?page page) (= ?name file-name) (= ?reportsToPosition position)]]
  ;; now we need to check to see if this manager position is just meant to be represented on another org chart - this may just be a separation point
  [?numberOfParticipatingOrgCharts <- (accum/count) :from [OrgChartPosition (not= ?page page) (= ?name file-name) (= ?reportsToPosition position) () ] ]
  [:test (= ?numberOfParticipatingOrgCharts 0)]
  =>
  (tap> {
         :page ?page
         :name ?name
         :description (str "XLSX specified a position not captured in this ORG chart:  " ?position " reports to " ?reportsToPosition)
  })
  (rules/insert! (->OrgChartError
                  ?page
                  ?reportsToPosition
                  nil ;; missing position
                  nil ;; duplicate position
                  ?path
                  ?name
                  (str "XLSX specified a position not captured in this ORG chart:" ?position " reports to " ?reportsToPosition))))




;; (rules/defrule detect-duplicate-positions
;;   "Detect duplicate position codes"
;;   [ Position (= ?posNum position) (= ?rowNum row-num)]
;;   [ Position (not= ?rowNum row-num) (= ?posNum position) (= ?otherRowNum row-num)]
;;   [:not [PositionWarning (= ?posNum position)]]
;;   => 
;;   (rules/insert! (->PositionWarning
;;                    ?posNum
;;                    (str "Duplicate position code detected: " ?posNum) "and" ?otherRowNum))
;;   )



(rules/defquery get-position-values
  "Query to get all position values"
  []
  [?position-values <- (accum/all) :from [ExtractedPosition]])


(rules/defquery get-all-org-chart-page-values
  "Query to get back all the org chart page values"
  []
  [?orgChartPageResults <- (accum/all) :from [OrgChartPageResult (= ?page page) (= ?fileName file-name) ]]
  [?orgChartPositionMatches <- (accum/all) :from [MatchingPosition (= ?page page) (= ?fileName file-name) ]]
   [?orgChartErrors <- (accum/all) :from [OrgChartError (= ?page page) (= ?fileName file-name)]]
  )


(rules/defquery get-all-org-chart-positions
  "Query to get back all the org chart positions"
  []
  [?orgChartPPositions <- (accum/all) :from [OrgChartPosition]])

(rules/defquery get-all-org-chart-errors
  "Query to get back all the org chart errors"
  []
  [?orgChartErrors <- (accum/all) :from [OrgChartError]])

;; (rules/defquery get-simple-position-values
;;   "Query to get all simple report values"
;;   []
;;   [?simpleReports <- (accum/all) :from [SimpleReport]])

(comment

  ;; (def eavs (eav/xlsx->eav (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx") :version :v1))
  ;; Clara-EAV expects raw EAV records, not transformed vectors

  ;; Session is now defined above, outside the comment block  ;; INCORRECT: Explicit rule vectors don't work properly in Clara-EAV
  (rules/defsession test-session 'clara-org-chart.rules)

  ;; 2. Streaming for very large files (memory efficient)
  (def results-streaming (-> test-session
                             (rules/insert-all 
                               (concat
                                 (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true))
                                 (extractor/load-org-chart-pages-as-records "extracted-org-chart-positions.edn")))
                             (rules/fire-rules)))

  (tap> (extractor/load-org-chart-pages-as-records "extracted-org-chart-positions.edn"))
  (tap> (:?position-values (first (rules/query results-streaming get-position-values))))
  (tap> (:?orgChartErrors (first (rules/query results-streaming get-all-org-chart-errors))))
  (tap> (:?orgChartPPositions (first (rules/query results-streaming get-all-org-chart-positions))))
  (tap> (:?orgChartPageResults (first (rules/query results-streaming get-all-org-chart-page-values))))
  (tap> (rules/query results-streaming get-all-org-chart-page-values))
  (tap> (:?position-values (first (rules/query results-streaming get-position-values))))
  (tap> (rules/query results-streaming get-simple-position-values))

  (def test-extraction (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true)))

  (tap> (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true))

  (tap> test-extraction)

  (pos/diagnose-hierarchy-issues test-extraction)


  ;; Get positions with subordinate counts calculated
(def positions-with-counts 
  (pos/extract-positions-with-counts 
    (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true)))

  (pos/verify-hierarchy-consistency positions-with-counts)

  (pos/debug-subordinate-calculation
   (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true))
   ["541-031-7500-001"])

(tap> positions-with-counts)
  

  
    ;; Generate SVG for specific codes
  (tangle/save-org-chart-for-codes (:?position-values (first (rules/query results-streaming get-position-values)))
                                   (pdf/positions-on-page "resources/Sac HQ Org Charts 01.01.25.pdf" 41)
                                   "San Bernardino Unit.svg"
                                   :title (str "Sac HQ Org Charts 01.01.25- Generated on " (java.time.LocalDate/now))
                                   :format "svg" 
                                   :strict-filter true
                                   :errors (:?orgChartErrors (first (rules/query results-streaming get-all-org-chart-errors)))
                                   )


  (tangle/save-org-chart-for-codes (:?position-values (first (rules/query results-streaming get-position-values)))
                                   ["541-028-4802-001"
                                     "541-028-4800-004"
                                     "541-028-4800-009"
                                     "541-028-4801-003" 
                                     "541-028-4800-015"
                                     "541-028-4800-016"
                                     "541-028-4800-904"
                                     "541-020-7500-008"
                                     "541-028-4800-022"]
                                   "Contracts & Grants.svg"
                                   :format "svg"
                                   :strict-filter true
                                   :errors (:?orgChartErrors (first (rules/query results-streaming get-all-org-chart-errors))))


  (tangle/save-org-chart-for-codes test-extraction
                                   (pdf/positions-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3))
  ;; Generate SVG for specific codes
  (tangle/save-org-chart-for-codes test-extraction
                                   (pdf/positions-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3)
                                   "subset-org-chart.svg"
                                   :format "svg")


  (tangle/save-org-chart-for-codes test-extraction
                                   (pdf/positions-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3)
                                   "subset-org-chart.dot"
                                   :format "dot")
  :rcf)