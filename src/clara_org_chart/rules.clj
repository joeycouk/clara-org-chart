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
                   OrgChartFileSummary
                   OrgChartFileExtraction
                   ExtractionMetadata
                   OrgChartPositionExtraction)))


(defrecord PositionWarning
           [position 
             description])

(defrecord ExtractedPosition
           [position
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
  [?pos <- Position (= position ?posNum)]
  [?subs <- (accum/count) :from [Position (= ?posNum reports-to-position)]]
  ;; [?subSubs <- (accum/sum :total-subordinates) :from [ExtractedPosition (= ?posNum reports-to-position)]]
  =>

  ;;  (println "Calculating total subordinates for position:" ?pos (get ?pos :position))
   (rules/insert! (->ExtractedPosition
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
  [?orgChartPageResults <- (accum/all) :from [OrgChartPageResult]])

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

  (tap> (rules/query results-streaming get-all-org-chart-page-values))
  (tap> (rules/query results-streaming get-position-values))
  (tap> (rules/query results-streaming get-simple-position-values))

  (def test-extraction (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true)))

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