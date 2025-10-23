
(ns clara-org-chart.rules
  (:require
   [clara-org-chart.xlsx :as xlsx]
   [org-tangle :as tangle]
   [clara-org-chart.position :as pos]
   [pdfBoxing :as pdf]
   [clara.rules :refer [defrule defquery insert! insert-all! insert-unconditional! defsession insert-all fire-rules query]]
   [clara.tools.inspect :as inspect]
   [clara.rules.accumulators :as accum]
   [clara-org-chart.org-chart-extractor :as extractor]
   [clojure.string :as str]
   [clara-org-chart.title-hierarchy :as th]
   [clara-org-chart.data-dictionary-extractor])
  (:import (clara_org_chart.position Position)
           (clara_org_chart.org_chart_extractor
            OrgChartPageResult)
           (clara_org_chart.data_dictionary_extractor
            ClassificationCode)))



;; Org chart page results have a list of positions. This isn't ideal for writing easy to understand rules. Rather than a list i prefer a single object with a position value
(defrecord OrgChartPosition
           [position      ; The position code text (string)
            file-name     ; File name where the position was found
            page          ; Page number where the position was found
            x             ; X coordinate of the position in the PDF  
            y             ; Y Coordinate of the position in the PDF
            width         ; Width of the position text bounding box
            height        ; Height of the position text bounding box
            ])


(defrecord TempDuplicateMatchingPosition
           [row-num
            position
            title
            class-title
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
            ;; unique-flag
            time-base
            tb-adjustment
            region
            direct-subordinates
            total-subordinates
            part-time
            temporary
            file-name
            page
            info])


(defrecord TemporaryPositionMatchingPosition
           [row-num
            position
            title
            class-title
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
            ;; unique-flag
            time-base
            tb-adjustment
            region
            direct-subordinates
            total-subordinates
            part-time
            temporary
            file-name
            page
            info])



(defrecord MatchingPosition
           [row-num
            position
            title
            class-title
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
            ;; unique-flag
            time-base
            tb-adjustment
            region
            direct-subordinates
            total-subordinates
            part-time
            temporary
            file-name
            page
            info
            
            ])






(defrecord OrgChartError
           [page
            positionCodes
            path
            file-name
            description])

(defrecord PositionWarning
           [position
            description])

(defrecord ExtractedPosition
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
            ;; unique-flag
            time-base
            tb-adjustment
            region
            direct-subordinates
            total-subordinates
            part-time
            temporary
            ])





(defn parse-position-with-time-base
  "Parse a position string that may contain (.5) indicating part-time.
  Returns a map with :position (cleaned) and :part-time (boolean).
  
  Examples:
    '542-062-1039-904 (.5)' -> {:position '542-062-1039-904', :part-time true}
    '542-062-1039-904'      -> {:position '542-062-1039-904', :part-time false}"
  [position-str]
  (when (and position-str (string? position-str))
    (let [trimmed (str/trim position-str)
          part-time? (str/includes? trimmed "(.5)")
          cleaned-position (-> trimmed
                               (str/replace #"\s*\(\.5\)\s*" "")
                               str/trim)]
      {:position cleaned-position
       :part-time part-time?})))


(defn ends-with-9xx?
  "Return true if the last or second-to-last 3-digit section of a string starts with '9'.
  Handles cases like '542-720-1060-910' => true, '542-720-1041-906-002' => true, '542-720-1060-001' => false."
  [s]
  (when (string? s)
    (let [matches (re-seq #"\d{3,4}" s)
          serial-num (nth matches 3 "")
          ]
      (cond
        (nil? serial-num) false
        (str/starts-with? serial-num "9") true
  :else false))))


(defn extract-four-digit-section  
  "Extract the first 4-digit section from a string like '542-064-1095-904-019'."  
  [s]  
  (let [pattern #"\b\d{4}\b"]    
  (first (re-seq pattern s))))

(defn convert-to-matching-positions
  "Convert a set of records into a sequence of MatchingPosition instances."
  [records]
  (map (fn [record]
         (map->MatchingPosition
          {:row-num (:row-num record)
           :position (:position record)
           :title (:title record)
           :class-title (:class-title record)
           :current-employee (:current-employee record)
           :reports-to-position (:reports-to-position record)
           :dotted-line-reports-to-position (:dotted-line-reports-to-position record)
           :city (:city record)
           :comments-notes (:comments-notes record)
           :dotted-reports (:dotted-reports record)
           :agencycode (:agencycode record)
           :unitcode (:unitcode record)
           :classcode (:classcode record)
           :serialnumber (:serialnumber record)
          ;;  :unique-flag (:unique-flag record)
           :time-base (:time-base record)
           :tb-adjustment (:tb-adjustment record)
           :region (:region record)
           :direct-subordinates (:direct-subordinates record)
           :total-subordinates (:total-subordinates record)
           :part-time (:part-time record)
           :temporary (:temporary record)
           :file-name (:file-name record)
           :page (:page record)
           :info (:info record)}))
       records))



  (defrule calculate-total-subordinates
    "Recursively calculate total subordinates for each position"
    [?pos <- Position (= ?posNum position)]
    [?subs <- (accum/count) :from [Position (= ?posNum reports-to-position)]]
    ;; [?subSubs <- (accum/sum :total-subordinates) :from [ExtractedPosition (= ?posNum reports-to-position)]]
    =>

    (let [position-info (parse-position-with-time-base (get ?pos :position))
          cleaned-position (:position position-info)
          part-time? (:part-time position-info)
          temporary? (ends-with-9xx? ?posNum)
          ]
      (insert! (->ExtractedPosition
                (get ?pos :row-num)
                cleaned-position
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
                ;; (get ?pos :unique-flag)
                (get ?pos :time-base)
                (get ?pos :tb-adjustment)
                (get ?pos :region)
                (or ?subs 0)
                (get ?pos :total-subordinates)
                part-time?
                temporary?
                ;;  (+ ?subs (or ?subSubs 0))
                ))))


  (defrule break-org-chart-pages-into-positions
    "Break org chart pages into individual positions"
    [OrgChartPageResult (= ?page page) (= ?positions positions) (= ?fileName file-name)]
    =>
    (doseq [pos-record ?positions]
      (insert! (->OrgChartPosition
                (:text pos-record)    ; Extract position text from PositionWithCoordinates
                ?fileName
                ?page
                (:x pos-record)       ; X coordinate
                (:y pos-record)       ; Y coordinate  
                (:width pos-record)   ; Width
                (:height pos-record)  ; Height
                ))))

  

 (defrule find-matching-temporary-position-numbers-and-xlsx-positions
   "For temporary positions only - add the Matching Position number only when they report to somebody directly within the org chart"
    [OrgChartPosition (= ?page page) (= ?position position) (= ?fileName file-name)]
    [?extractedPosition <- ExtractedPosition (= ?position position) (= ?reportsToPosition reports-to-position)]
    [:test (= (get ?extractedPosition :temporary) true)]
    [:test (not-empty ?reportsToPosition)]
   [MatchingPosition (= ?page page) (= ?reportsToPosition position) (= ?fileName file-name)] 
   [ClassificationCode (= (extract-four-digit-section ?position) class-code) (= ?classTitle class-title)]
   =>
    (insert-unconditional! (->TemporaryPositionMatchingPosition
                           (get ?extractedPosition :row-num)
                           (get ?extractedPosition :position)
                           (get ?extractedPosition :title)
                           ?classTitle
                           (get ?extractedPosition :current-employee)
                           (get ?extractedPosition :reports-to-position)
                           (get ?extractedPosition :dotted-line-reports-to-position)
                           (get ?extractedPosition :city)
                           (get ?extractedPosition :comments-notes)
                           (get ?extractedPosition :dotted-reports)
                           (get ?extractedPosition :agencycode)
                           (get ?extractedPosition :unitcode)
                           (get ?extractedPosition :classcode)
                           (get ?extractedPosition :serialnumber)
                          ;;  (get ?extractedPosition :unique-flag)
                           (get ?extractedPosition :time-base)
                           (get ?extractedPosition :tb-adjustment)
                           (get ?extractedPosition :region)
                           (get ?extractedPosition :direct-subordinates)
                           (get ?extractedPosition :total-subordinates)
                           (get ?extractedPosition :part-time)
                           (get ?extractedPosition :temporary)
                           ?fileName
                           ?page
                           ""))
   
   )

  

 (defrule find-matching-temporary-position-numbers-and-xlsx-positions-no-class-code
   "For temporary positions only - add the Matching Position number only when they report to somebody directly within the org chart"
   [OrgChartPosition (= ?page page) (= ?position position) (= ?fileName file-name)]
   [?extractedPosition <- ExtractedPosition (= ?position position) (= ?reportsToPosition reports-to-position)]
   [:test (= (get ?extractedPosition :temporary) true)]
   [:test (not-empty ?reportsToPosition)]
   [MatchingPosition (= ?page page) (= ?reportsToPosition position) (= ?fileName file-name)]
   [:not [ClassificationCode (= (extract-four-digit-section ?position) class-code)]]
   =>
   (insert-unconditional! (->TemporaryPositionMatchingPosition
                           (get ?extractedPosition :row-num)
                           (get ?extractedPosition :position)
                           (get ?extractedPosition :title)
                           "Unknown"
                           (get ?extractedPosition :current-employee)
                           (get ?extractedPosition :reports-to-position)
                           (get ?extractedPosition :dotted-line-reports-to-position)
                           (get ?extractedPosition :city)
                           (get ?extractedPosition :comments-notes)
                           (get ?extractedPosition :dotted-reports)
                           (get ?extractedPosition :agencycode)
                           (get ?extractedPosition :unitcode)
                           (get ?extractedPosition :classcode)
                           (get ?extractedPosition :serialnumber)
                          ;;  (get ?extractedPosition :unique-flag)
                           (get ?extractedPosition :time-base)
                           (get ?extractedPosition :tb-adjustment)
                           (get ?extractedPosition :region)
                           (get ?extractedPosition :direct-subordinates)
                           (get ?extractedPosition :total-subordinates)
                           (get ?extractedPosition :part-time)
                           (get ?extractedPosition :temporary)
                           ?fileName
                           ?page
                           "")))




  (defrule find-matching-org-chart-position-numbers-and-xlsx-positions
    "Where the position numbers match perfectly, generate an instance of MatchingPosition"
    [OrgChartPosition (= ?page page) (= ?position position) (= ?fileName file-name)]
    [ClassificationCode (= (extract-four-digit-section ?position) class-code) (= ?classTitle class-title)]
    [?extractedPosition <- ExtractedPosition (= ?position position)]
    ;; Not mapping in temporary positions in this function because there could be dupolicate position codes
    [:test (not= (get ?extractedPosition :temporary) true)]
    =>
    (insert! (->MatchingPosition
              (get ?extractedPosition :row-num)
              (get ?extractedPosition :position)
              (get ?extractedPosition :title)
              ?classTitle
              (get ?extractedPosition :current-employee)
              (get ?extractedPosition :reports-to-position)
              (get ?extractedPosition :dotted-line-reports-to-position)
              (get ?extractedPosition :city)
              (get ?extractedPosition :comments-notes)
              (get ?extractedPosition :dotted-reports)
              (get ?extractedPosition :agencycode)
              (get ?extractedPosition :unitcode)
              (get ?extractedPosition :classcode)
              (get ?extractedPosition :serialnumber)
              ;; (get ?extractedPosition :unique-flag)
              (get ?extractedPosition :time-base)
              (get ?extractedPosition :tb-adjustment)
              (get ?extractedPosition :region)
              (get ?extractedPosition :direct-subordinates)
              (get ?extractedPosition :total-subordinates)
              (get ?extractedPosition :part-time)
              (get ?extractedPosition :temporary)
              ?fileName
              ?page
              "")))

  

  (defrule find-matching-org-chart-position-numbers-and-xlsx-positions-no-matching-class-code
  "Where the position numbers match perfectly, generate an instance of MatchingPosition with no matching classification code known"
  [OrgChartPosition (= ?page page) (= ?position position) (= ?fileName file-name)]
  [:not [ClassificationCode (= (extract-four-digit-section ?position) class-code)]]
  [?extractedPosition <- ExtractedPosition (= ?position position)] 
    ;; Not mapping in temporary positions in this function because there could be dupolicate position codes
  [:test (not= (get ?extractedPosition :temporary) true)]
  =>
  (insert! (->MatchingPosition
            (get ?extractedPosition :row-num)
            (get ?extractedPosition :position)
            (get ?extractedPosition :title)
            "Unknown"
            (get ?extractedPosition :current-employee)
            (get ?extractedPosition :reports-to-position)
            (get ?extractedPosition :dotted-line-reports-to-position)
            (get ?extractedPosition :city)
            (get ?extractedPosition :comments-notes)
            (get ?extractedPosition :dotted-reports)
            (get ?extractedPosition :agencycode)
            (get ?extractedPosition :unitcode)
            (get ?extractedPosition :classcode)
            (get ?extractedPosition :serialnumber)
            ;; (get ?extractedPosition :unique-flag)
            (get ?extractedPosition :time-base)
            (get ?extractedPosition :tb-adjustment)
            (get ?extractedPosition :region)
            (get ?extractedPosition :direct-subordinates)
            (get ?extractedPosition :total-subordinates)
            (get ?extractedPosition :part-time)
            (get ?extractedPosition :temporary)
            ?fileName
            ?page
            "")))


  
  ;; deduplicate the temporary possitions
  (defrule flatten-out-temporary-position-duplicates
    "deduplicate the temporary possitions"
    [?distinctDuplicates <- (accum/distinct) :from [TemporaryPositionMatchingPosition]]
    =>
    ;; (tap> (convert-to-matching-positions ?distinctDuplicates))
    (insert-all! (convert-to-matching-positions ?distinctDuplicates)))


  ;; deduplicate the consistently missing manager positions
  (defrule flatten-out-missing-manager-duplicates
    "deduplicate the missing managers identified"
    [?distinctDuplicates <- (accum/distinct) :from [TempDuplicateMatchingPosition]]
    =>
    ;; (tap> (convert-to-matching-positions ?distinctDuplicates))
    (insert-all! (convert-to-matching-positions ?distinctDuplicates)))



  ;; New rule to identify when multiple extract positions within the same page report to a person but that person is not included yet
  (defrule find-consistently-missing-manager
    "identify when multiple extract positions within the same page report to a person but that person is not included yet"
    [OrgChartPosition (= ?page page) (= ?position position) (= ?fileName file-name)] 
    [ExtractedPosition (= ?position position) (= ?reportsToPosition reports-to-position)]
    [:test (not-empty ?reportsToPosition)]
    [ClassificationCode (= (extract-four-digit-section ?reportsToPosition) class-code) (= ?classTitle class-title)]
    [:not [OrgChartPosition (= ?page page) (= ?reportsToPosition position) (= ?fileName file-name)]]
    [?reportsToExtractedPosition <- ExtractedPosition (= ?reportsToPosition position)]
    [?numberOfPeopleReportingToThisGuy <- (accum/count) :from [MatchingPosition (= ?page page) (= ?reportsToPosition reports-to-position) (= ?fileName file-name)]]
    [:test (> ?numberOfPeopleReportingToThisGuy 1)]

    =>
    (insert-unconditional! (->TempDuplicateMatchingPosition
                            (get ?reportsToExtractedPosition :row-num)
                            (get ?reportsToExtractedPosition :position)
                            (get ?reportsToExtractedPosition :title)
                            ?classTitle
                            (get ?reportsToExtractedPosition :current-employee)
                            (get ?reportsToExtractedPosition :reports-to-position)
                            (get ?reportsToExtractedPosition :dotted-line-reports-to-position)
                            (get ?reportsToExtractedPosition :city)
                            (get ?reportsToExtractedPosition :comments-notes)
                            (get ?reportsToExtractedPosition :dotted-reports)
                            (get ?reportsToExtractedPosition :agencycode)
                            (get ?reportsToExtractedPosition :unitcode)
                            (get ?reportsToExtractedPosition :classcode)
                            (get ?reportsToExtractedPosition :serialnumber)
                            ;; (get ?reportsToExtractedPosition :unique-flag)
                            (get ?reportsToExtractedPosition :time-base)
                            (get ?reportsToExtractedPosition :tb-adjustment)
                            (get ?reportsToExtractedPosition :region)
                            (get ?reportsToExtractedPosition :direct-subordinates)
                            (get ?reportsToExtractedPosition :total-subordinates)
                            (get ?reportsToExtractedPosition :part-time)
                            (get ?reportsToExtractedPosition :temporary)
                            ?fileName
                            ?page
                            "Supervisor Not on PDF")))




  ;; TODO rename this rule
  (defrule detect-org-chart-position-mismatches
    "Detect positions where the org chart pdf specifies a position number which does not exist within the xlsx document"
    [OrgChartPosition (= ?page page) (= ?position position) (= ?name file-name)]
    [:not [MatchingPosition (= ?page page) (= ?name file-name) (= ?position position)]]
    [OrgChartPageResult (= ?page page) (= ?path file-path) (= ?name file-name) (= ?description description)]
    =>
    (insert! (->OrgChartError
              ?page
              (vector ?position)
              ?path
              ?name
              "Orgchart position specified in the PDF that does not exist in the xlsx document")))


  ;; (defrule identify-invalid-subordinate
  ;;   "Detect when a person is a subordinate of an invalid person based on title"
  ;;   [MatchingPosition (= ?position position) (= ?reportsToPosition reports-to-position) (= ?title class-title) (= ?fileName file-name) (= ?page page)]
  ;;   [:test (not-empty ?reportsToPosition)]
  ;;   [MatchingPosition (= ?fileName file-name) (= ?page page) (= ?reportsToPosition position) (= ?titleSup class-title) ]
  ;;   [:test (nil? (th/can-report-to? ?title ?titleSup))]
  ;;   =>
  ;;   ;; (tap> {:title "identify-invalid-subordinate" :page ?page :positions  (vector ?position ?reportsToPosition) :name ?fileName})
  ;;   (insert! (->OrgChartError
  ;;             ?page
  ;;             (vector ?position ?reportsToPosition)
  ;;             ""
  ;;             ?fileName
  ;;             (str "Invalid Supervisor detected for " ?position " " ?title " "  "and "  ?reportsToPosition " " ?titleSup))))



  (defrule detect-org-chart-position-duplicates
    "Detect when multiple positions are mapped to an org chart with the same position number"
    [?matchingRows <- (accum/distinct :position) :from [MatchingPosition (= ?page page) (= ?name file-name) (= ?position position)]]
    [:test (> (count ?matchingRows) 1)]
    [OrgChartPageResult (= ?page page) (= ?path file-path) (= ?name file-name) (= ?description description)]
    =>
    (insert! (->OrgChartError
              ?page
              (vector ?position) ;;duplicate position
              ?path
              ?name
              (str "Orgchart has multiple matched positions with the same number " ?matchingRows))))


  (defrule detect-org-chart-missing-position
    "Detect position contained within the xlsx file but not within the org chart pdf extraction"
    [OrgChartPageResult
     (= ?page page)
     (= ?path file-path)
     (= ?name file-name)
     (= ?description description)
     (= ?positions positions)]
    [MatchingPosition (= ?page page) (= ?name file-name) (= ?position position) (= ?reportsToPosition reports-to-position)]
    ;; where the reports to position doesn't exist in the pdf
    [:not [OrgChartPosition (= ?page page) (= ?name file-name) (= ?reportsToPosition position)]]
    ;; now we need to check to see if this manager position is just meant to be represented on another org chart - this may just be a separation point
    [?numberOfParticipatingOrgCharts <- (accum/count) :from [OrgChartPosition (= file-name ?name ) (= position ?reportsToPosition)]]
    [:test (= ?numberOfParticipatingOrgCharts 0)]
    =>
    ;; (tap> {:page ?page :positions  (vector ?position ?reportsToPosition) :path ?path :name ?name})
    (insert! (->OrgChartError
              ?page
              (vector ?position ?reportsToPosition)
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



  (defquery get-position-values
    "Query to get all position values"
    []
    [?position-values <- (accum/all) :from [ExtractedPosition]])


  (defquery get-all-unique-position-titles
    "For the purposes of creating a unique list of position titles"
    []
    [?position-values <- (accum/distinct :title) :from [ExtractedPosition]])


    (defquery get-all-classification-codes
    "For the purposes of creating a unique list of position titles"
    []
    [?classCodes <- (accum/all) :from [ClassificationCode]])

    (defquery get-class-title-from-classification-code
    "For the purposes of creating a unique list of position titles"
    [:?classCode]
    [?classCodes <- (accum/all) :from [ClassificationCode (= ?classCode class-code)]])

  


  (defquery get-all-matched-positions
  "interested in finding match positions where :info !== blank"
  []
  [?matchedPositions <- (accum/all) :from [MatchingPosition]]
    )


  (defquery get-wacky-matched-positions
    "interested in finding match positions where :info !== blank"
    []
    [?matchedPositions <- (accum/all) :from [MatchingPosition (not= "" info)]]
    )



  (defquery get-all-org-chart-page-values
    "Query to get back all the org chart page values"
    []
    [?orgChartPageResults <- (accum/all) :from [OrgChartPageResult (= ?page page) (= ?fileName file-name)]]
    [?orgChartPositionMatches <- (accum/all) :from [MatchingPosition (= ?page page) (= ?fileName file-name)]]
    [?orgChartErrors <- (accum/all) :from [OrgChartError (= ?page page) (= ?fileName file-name)]])


  (defquery get-specific-org-chart-page-values
    "Query to get back a specified the org chart page value"
    [:?fileName :?page]
    [?orgChartPageResults <- (accum/all) :from [OrgChartPageResult (= ?page page) (= ?fileName file-name)]]
    [?orgChartPositionMatches <- (accum/all) :from [MatchingPosition (= ?page page) (= ?fileName file-name)]]
    [?orgChartErrors <- (accum/all) :from [OrgChartError (= ?page page) (= ?fileName file-name)]])

  
  (defquery get-sampling-of-supervisor-supverisee
    "Query all supervisors and their supervisors position titles"
    []
    [MatchingPosition(= ?position position) (= ?reportsToPosition reports-to-position)]
    [ClassificationCode (= (extract-four-digit-section ?position) class-code) (= ?subordinateTitle class-title)]
    [ClassificationCode (= (extract-four-digit-section ?reportsToPosition) class-code) (= ?managerTitle class-title)]
    )
  

    (defquery get-all-extracted-positions-without-class-codes 
      ""
    [] 
      [ExtractedPosition (= ?position position) ]
      [:test (not-empty ?position)]
      [not [ClassificationCode (= (extract-four-digit-section ?position) class-code)]]
    )


  (defquery get-all-org-chart-positions
    "Query to get back all the org chart positions"
    []
    [?orgChartPPositions <- (accum/all) :from [OrgChartPosition]])

  (defquery get-all-org-chart-errors
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
    ;; (load-file "src/clara_org_chart/rules.clj")

  (th/can-report-to? "CUSTODIAN SUPERVISOR I" "CHIEF OF PLANT OPERATION I")

    ;; Session is now defined above, outside the comment block  ;; INCORRECT: Explicit rule vectors don't work properly in Clara-EAV
    (defsession test-session 'clara-org-chart.rules)

     (pos/extract-positions (xlsx/extract-data "resources/CNR_MOC01.xlsx" :streaming true))
    
(tap> (pos/extract-positions-with-counts (xlsx/extract-data "resources/CNR_MOC01.xlsx" :streaming true)))


    ;; 2. Streaming for very large files (memory efficient)
    (def results-streaming (-> test-session
                               (insert-all
                                (concat
                                ;;  (pos/extract-positions (xlsx/extract-data "resources/CNR_MOC01.xlsx" :streaming true))
                                 (pos/extract-positions-with-counts (xlsx/extract-data "resources/CNR_MOC01.xlsx" :streaming true))
                                 (extractor/load-org-chart-pages-as-records "extracted-org-chart-positions.edn")
                                 (clara-org-chart.data-dictionary-extractor/extract-classification-codes (xlsx/extract-data "resources/CAL FIRE Data Dictionary.xlsx"))))
                               (fire-rules)))
    
    
    
    (tap> (get-in (inspect/inspect results-streaming) [:rule-matches detect-org-chart-missing-position]))

    (tap> test-inspect)
    (def test-inspect (inspect/inspect results-streaming))


    
(tap>  (inspect/explain-activations(-> test-session
                                  (insert-all
                                   (concat
                                    (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true))
                                    (extractor/load-org-chart-pages-as-records "extracted-org-chart-positions.edn")
                                    (clara-org-chart.data-dictionary-extractor/extract-classification-codes (xlsx/extract-data "resources/CAL FIRE Data Dictionary.xlsx"))))
                                  (fire-rules))))

    (tap> (extractor/load-org-chart-pages-as-records "extracted-org-chart-positions.edn"))
    (tap> (:?position-values (first (query results-streaming get-position-values))))
    (tap> (:?orgChartErrors (first (query results-streaming get-all-org-chart-errors))))
    (tap> (:?orgChartPPositions (first (query results-streaming get-all-org-chart-positions))))
    
    (tap> (:?orgChartPageResults (first (query results-streaming get-all-org-chart-page-values))))

    (tap> (query results-streaming get-class-title-from-classification-code :?classCode "1060"))

 (tap> (filter #(not (string? (:reports-to-position %))) (:?matchedPositions (first (query results-streaming get-all-matched-positions)))))


    (tap> (query results-streaming get-wacky-matched-positions))
    (tap> (query results-streaming get-sampling-of-supervisor-supverisee))

    (tap> (group-by :?subordinateTitle (query results-streaming get-sampling-of-supervisor-supverisee)))

        (tap> (into {}
                (map (fn [[sub-title items]]
                       [sub-title (set (map :?managerTitle items))]))
                (group-by :?subordinateTitle (query results-streaming get-sampling-of-supervisor-supverisee))))

     (tap> (into {}
      (map (fn [[sub-title items]]
        [sub-title (set (map :?managerTitle items))]))
      (group-by :?subordinateTitle (query results-streaming get-sampling-of-supervisor-supverisee))))

    (tap> (query results-streaming get-all-extracted-positions-without-class-codes))

    (tap> (query results-streaming get-all-classification-codes))
    (tap> (query results-streaming get-all-org-chart-page-values))
    (tap> (query results-streaming get-specific-org-chart-page-values :?fileName "Sac HQ Org Charts 01.01.25" :?page 43))
    (tap> (:?position-values (first (query results-streaming get-position-values))))
    (tap> (query results-streaming get-simple-position-values))
    (tap> (query results-streaming get-all-unique-position-titles))

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


    (tap> (:?orgChartPositionMatches (first (query results-streaming get-specific-org-chart-page-values :?fileName "Sac HQ Org Charts 01.01.25" :?page 47))))


    (let [queryResult (query results-streaming get-specific-org-chart-page-values :?fileName "Sac HQ Org Charts 01.01.25" :?page 72)
          positions (:?orgChartPositionMatches (first queryResult))
          errors (:?orgChartErrors (first queryResult))
          title (:description (first (:?orgChartPageResults (first queryResult))))
          format "svg"
          pdfDocuumentName (:file-name (first (:?orgChartPageResults (first queryResult))))
          page (:page (first (:?orgChartPageResults (first queryResult))))
          filename (str pdfDocuumentName "-" title " page " page "." format)]
      (tap> {:positions positions
             :errors errors
             :title title
             :format format
             :pdfDocuumentName pdfDocuumentName
             :page page
             :filename filename})

      (tangle/save-org-chart-for-codes positions
                                       []
                                       filename
                                       :title title
                                       :format format
                                       :errors errors
                                       :report-missing-positions false))
    



    (let [queryResult (query results-streaming get-all-org-chart-page-values)
      results-with-errors (filter #(> (count (:?orgChartErrors %)) 0) queryResult)]
      (doseq [orgChart results-with-errors]
        (let [positions (:?orgChartPositionMatches orgChart)
             errors (:?orgChartErrors orgChart)
             title (:description (first (:?orgChartPageResults orgChart)))
             format "svg"
             pdfDocuumentName (:file-name (first (:?orgChartPageResults  orgChart)))
             page (:page (first (:?orgChartPageResults orgChart)))
             filename (str pdfDocuumentName "-" title " page " page "." format)]
         (tap> {:positions positions
               :errors errors
               :title title
               :format format
               :pdfDocuumentName pdfDocuumentName
               :page page
               :filename filename})
        
        (tangle/save-org-chart-for-codes positions
                                         []
                                         filename
                                         :title title
                                         :format format
                                         :errors errors
                                         :report-missing-positions false))
        ))

          


    ;; (tap> (:?position-values (first (query results-streaming get-position-values))))

    ;; Generate SVG for specific codes
    (tangle/save-org-chart-for-codes (:?position-values (first (query results-streaming get-position-values)))
                                     (pdf/positions-on-page "resources/Sac HQ Org Charts 01.01.25.pdf" 41)
                                     "San Bernardino Unit.svg"
                                     :title (str "Sac HQ Org Charts 01.01.25- Generated on " (java.time.LocalDate/now))
                                     :format "svg"
                                     :errors (:?orgChartErrors (first (query results-streaming get-all-org-chart-errors))))


    (tangle/save-org-chart-for-codes (:?position-values (first (query results-streaming get-position-values)))
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
                                     :errors (:?orgChartErrors (first (query results-streaming get-all-org-chart-errors))))


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

