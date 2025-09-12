(ns clara-org-chart.org-chart-extractor
  (:require [clojure.edn :as edn]
            [clojure.java.io :as io]
            [pdfBoxing :as pdf]
            [clojure.pprint :as pprint]))

;; Records for structured org chart position extraction data

(defrecord PositionWithCoordinates
  [text                ; The position code string (e.g., "542-434-1083-901")
   x                   ; X coordinate of the position code in the PDF
   y                   ; Y coordinate of the position code in the PDF
   width               ; Width of the position code text bounding box
   height])            ; Height of the position code text bounding box

(defrecord OrgChartPageResult
  [page                ; Page number (integer)
   description         ; Description of the org chart (string)
   position-count      ; Number of positions found (integer)
   positions           ; Vector of PositionWithCoordinates records
   status              ; :success or :error (keyword)
   error               ; Error message if status is :error (optional string)
   file-path           ; Path to the PDF file (string)
   file-name])         ; Display name of the file (string)

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
    (:error page-map)
    (:file-path page-map)
    (:file-name page-map)))

(defn load-org-chart-pages-as-records
  "Load org chart pages from extracted positions file and convert to OrgChartPageResult records.
  This is the main function to use when you need OrgChartPageResult records for Clara Rules."
  [extraction-file-path]
  (let [data (load-extracted-positions extraction-file-path)]
    (->> (:extractions data)
         (mapcat (fn [file-extraction]
                   (let [file-path (:file-path file-extraction)
                         file-name (:name file-extraction)]
                     (map (fn [page-map]
                            (page-map->record 
                              (assoc page-map 
                                     :file-path file-path
                                     :file-name file-name)))
                          (:org-charts file-extraction)))))
         vec)))

(defn extract-positions-from-chart-page
  "Extract positions from a specific org chart page. Returns an OrgChartPageResult record."
  [file-path page description file-name]
  (try
    (let [position-coords (pdf/positions-with-coordinates-on-page file-path page)
          position-records (mapv (fn [coord-map]
                                   (->PositionWithCoordinates
                                    (:text coord-map)
                                    (:x coord-map)
                                    (:y coord-map)
                                    (:width coord-map)
                                    (:height coord-map)))
                                 position-coords)]
      (->OrgChartPageResult
       page
       description
       (count position-records)
       position-records
       :success
       nil
       file-path
       file-name))
    (catch Exception e
      (->OrgChartPageResult
       page
       description
       0
       []
       :error
       (.getMessage e)
       file-path
       file-name))))

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
                              (let [result (extract-positions-from-chart-page file-path page description Name)]
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
    
    ;; extraction-results
    ))

;; Helper functions for working with PositionWithCoordinates records

(defn position-text
  "Extract the position code text from a PositionWithCoordinates record."
  [position-record]
  (:text position-record))

(defn positions->texts
  "Convert a collection of PositionWithCoordinates records to position code strings."
  [position-records]
  (map position-text position-records))

(defn get-positions-for-description
  "Helper function to find all positions for org charts matching a description pattern.
  Returns a vector of PositionWithCoordinates records."
  [extraction-results description-pattern]
  (let [pattern (re-pattern (str "(?i)" description-pattern))] ; case-insensitive
    (->> (:extractions extraction-results)
         (mapcat :org-charts)
         (filter #(re-find pattern (:description %)))
         (mapcat :positions)
         ;; Deduplicate by position text while preserving first occurrence with coordinates
         (group-by :text)
         (map (fn [[_text records]] (first records)))
         vec)))

(defn get-position-texts-for-description
  "Helper function to find all position code texts for org charts matching a description pattern.
  Returns a vector of position code strings (legacy compatibility)."
  [extraction-results description-pattern]
  (->> (get-positions-for-description extraction-results description-pattern)
       (map :text)
       vec))

(defn get-positions-by-file
  "Helper function to get all positions organized by file.
  Takes an OrgChartPositionExtraction record and returns a vector of file summaries with PositionWithCoordinates."
  [extraction-results]
  (->> (:extractions extraction-results)
       (map (fn [{:keys [name file-path org-charts]}]
              {:file-name name
               :file-path file-path
               :all-positions (->> org-charts
                                   (mapcat :positions)
                                   ;; Deduplicate by position text, keeping first occurrence
                                   (group-by :text)
                                   (map (fn [[_text records]] (first records)))
                                   vec)
               :all-position-texts (->> org-charts
                                        (mapcat :positions)
                                        (map :text)
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