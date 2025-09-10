(ns clara-org-chart.org-chart-extractor
  (:require [clojure.edn :as edn]
            [clojure.java.io :as io]
            [pdfBoxing :as pdf]
            [clojure.pprint :as pprint]))

;; Records for structured org chart position extraction data

(defrecord OrgChartPageResult
  [page                ; Page number (integer)
   description         ; Description of the org chart (string)
   position-count      ; Number of positions found (integer)
   positions           ; Vector of position code strings
   status              ; :success or :error (keyword)
   error])             ; Error message if status is :error (optional string)

(defrecord OrgChartFileSummary
  [total-pages         ; Total number of pages processed (integer)
   successful-pages    ; Number of successfully processed pages (integer)  
   total-positions])   ; Total positions extracted across all pages (integer)

(defrecord OrgChartFileExtraction
  [file-path           ; Path to the PDF file (string)
   name                ; Display name of the file (string)
   summary             ; OrgChartFileSummary record
   org-charts])        ; Vector of OrgChartPageResult records

(defrecord ExtractionMetadata
  [extracted-at        ; Timestamp of extraction (java.time.Instant)
   total-files])       ; Number of files processed (integer)

(defrecord OrgChartPositionExtraction
  [metadata            ; ExtractionMetadata record
   extractions])       ; Vector of OrgChartFileExtraction records

(defn load-org-chart-map
  "Load the org chart mapping from the edn file."
  [map-file-path]
  (-> map-file-path
      slurp
      edn/read-string))

(defn load-extracted-positions
  "Load the extracted positions from EDN file. Returns plain maps, not records."
  [extraction-file-path]
  (-> extraction-file-path
      slurp
      edn/read-string))

(defn load-extracted-positions-as-records
  "Load extracted positions and convert back to record types."
  [extraction-file-path]
  (let [data (load-extracted-positions extraction-file-path)]
    ;; Convert back to records if needed - for now return as maps since they're easier to work with
    data))

(defn page-map->record
  "Convert a map to an OrgChartPageResult record."
  [page-map]
  (->OrgChartPageResult 
    (:page page-map)
    (:description page-map)
    (:position-count page-map)
    (:positions page-map)
    (:status page-map)
    (:error page-map)))

(defn load-org-chart-pages-as-records
  "Load org chart pages from extracted positions file and convert to OrgChartPageResult records.
  This is the main function to use when you need OrgChartPageResult records for Clara Rules."
  [extraction-file-path]
  (->> (load-extracted-positions extraction-file-path)
       :extractions
       (mapcat :org-charts)
       (map page-map->record)
       vec))

(defn extract-positions-from-chart-page
  "Extract positions from a specific org chart page. Returns an OrgChartPageResult record."
  [file-path page description]
  (try
    (let [positions (pdf/positions-on-page file-path page :unique? true)]
      (->OrgChartPageResult
       page
       description
       (count positions)
       positions
       :success
       nil))
    (catch Exception e
      (->OrgChartPageResult
       page
       description
       0
       []
       :error
       (.getMessage e)))))

(defn extract-positions-from-org-charts
  "Extract all positions from all org charts defined in the mapping.
  Returns an OrgChartPositionExtraction record."
  [org-chart-map]
  (let [extraction-timestamp (java.time.Instant/now)
        extractions
        (mapv (fn [{:keys [file-path Name org-charts]}]
                (println "Processing:" Name "(" (count org-charts) "pages)")
                (let [chart-results
                      (mapv (fn [{:keys [page description]}]
                              (print "  Page" page "...")
                              (let [result (extract-positions-from-chart-page file-path page description)]
                                (println (if (= (:status result) :success)
                                           (str (:position-count result) " positions")
                                           "ERROR"))
                                result))
                            org-charts)
                      
                      total-positions (reduce + 0 (map :position-count chart-results))
                      successful-pages (count (filter #(= (:status %) :success) chart-results))
                      
                      summary (->OrgChartFileSummary
                               (count org-charts)
                               successful-pages
                               total-positions)]
                  
                  (->OrgChartFileExtraction
                   file-path
                   Name
                   summary
                   chart-results)))
              org-chart-map)
        
        metadata (->ExtractionMetadata extraction-timestamp (count org-chart-map))]
    
    (->OrgChartPositionExtraction metadata extractions)))

(defn convert-for-edn
  "Convert records to EDN-friendly maps, handling Java objects like Instant."
  [data]
  (cond
    (instance? java.time.Instant data)
    (str data)  ; Convert Instant to string
    
    (record? data)
    (into {} (map (fn [[k v]] [k (convert-for-edn v)]) data))
    
    (sequential? data)
    (mapv convert-for-edn data)
    
    (map? data)
    (into {} (map (fn [[k v]] [k (convert-for-edn v)]) data))
    
    :else
    data))

(defn save-position-extraction
  "Save the position extraction results to an EDN file, converting records to maps."
  [extraction-results output-path]
  (let [edn-friendly-data (convert-for-edn extraction-results)]
    (with-open [w (io/writer output-path)]
      (pprint/pprint edn-friendly-data w))
    (println "Saved position extraction to:" output-path)))

(defn extract-and-save-org-chart-positions
  "Main function: Load org chart map, extract positions, and save results."
  [map-file-path output-path]
  (println "Loading org chart mapping from:" map-file-path)
  (let [org-chart-map (load-org-chart-map map-file-path)
        _ (println "Found" (count org-chart-map) "file(s) to process")
        extraction-results (extract-positions-from-org-charts org-chart-map)]
    
    (save-position-extraction extraction-results output-path)
    
    ;; Print summary
    (let [total-charts (reduce + 0 (map #(get-in % [:summary :total-pages]) (:extractions extraction-results)))
          total-positions (reduce + 0 (map #(get-in % [:summary :total-positions]) (:extractions extraction-results)))]
      (println "\n=== EXTRACTION SUMMARY ===")
      (println "Total org chart pages processed:" total-charts)
      (println "Total position codes extracted:" total-positions))
    
    extraction-results))

(defn get-positions-for-description
  "Helper function to find all positions for org charts matching a description pattern.
  Takes an OrgChartPositionExtraction record and returns a vector of position codes."
  [extraction-results description-pattern]
  (let [pattern (re-pattern (str "(?i)" description-pattern))] ; case-insensitive
    (->> (:extractions extraction-results)
         (mapcat :org-charts)
         (filter #(re-find pattern (:description %)))
         (mapcat :positions)
         distinct
         vec)))

(defn get-positions-by-file
  "Helper function to get all positions organized by file.
  Takes an OrgChartPositionExtraction record and returns a vector of file summaries."
  [extraction-results]
  (->> (:extractions extraction-results)
       (map (fn [{:keys [name file-path org-charts]}]
              {:file-name name
               :file-path file-path
               :all-positions (->> org-charts
                                   (mapcat :positions)
                                   distinct
                                   vec)}))
       vec))

;; Main execution function
(defn run-extraction!
  "Execute the org chart position extraction process."
  []
  (println "=== ORG CHART POSITION EXTRACTION ===")
  (extract-and-save-org-chart-positions 
    "src/org_chart_map.edn" 
    "extracted-org-chart-positions.edn"))

(comment
  ;; Run the extraction
  (run-extraction!)
  
  ;; Load and query results
  (def results (-> "extracted-org-chart-positions.edn" slurp clojure.edn/read-string))
  
  ;; Find positions in Leadership charts
  (get-positions-for-description results "Leadership")
  
  ;; Find positions in Accounting charts  
  (get-positions-for-description results "Accounting")
  
  ;; Get all positions by file
  (get-positions-by-file results)
  
  :rcf)