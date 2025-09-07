(ns clara-org-chart.rules
  (:require
             [clara-org-chart.xlsx :as xlsx]
             [org-tangle :as tangle]
             [clara-org-chart.position :as pos]
             [clara.rules :as rules]
             [clara.rules.accumulators :as accum]
             [clojure.java.io :as io]
             [clojure.java.shell :as shell])
  (:import (clara_org_chart.position Position))
  )


(rules/defquery get-position-values
  "Query to get all position values"
  []
  [?position-values <- (accum/all) :from [Position]])


(comment

  ;; (def eavs (eav/xlsx->eav (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx") :version :v1))
  ;; Clara-EAV expects raw EAV records, not transformed vectors

  ;; Session is now defined above, outside the comment block  ;; INCORRECT: Explicit rule vectors don't work properly in Clara-EAV
  (rules/defsession test-session 'clara-org-chart.rules)


  (def test-extraction (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true)))
  
  ;; Debug info
  (println "Current working directory:" (System/getProperty "user.dir"))
  (println "Number of positions extracted:" (count test-extraction))
  
  ;; Create an improved hierarchical org chart
  (try
    (println "Creating improved hierarchical org chart...")
    
  
;; Generate chart with the full dataset
(let [full-dataset rules/test-extraction]
  (println "Generating chart with full dataset...")
  (println "Total positions in full dataset:" (count full-dataset))
  
  ;; Generate with a reasonable subset but use full dataset for supervisor lookup
  (tangle/save-simplified-org-chart full-dataset "full-org-chart.dot"
                                    :format "dot"
                                    :max-positions 1000  ; Show 100 positions but lookup supervisors from full dataset
                                    :show-vacant true)
  
  ;; Check if John is now included
  (let [basic-selection (take 100 full-dataset)
        with-supervisors (tangle/ensure-supervisors-included basic-selection full-dataset)
        john-included? (some #(= (:position %) "542-411-9723-610") with-supervisors)]
    
    (println "Basic selection:" (count basic-selection))
    (println "With supervisors:" (count with-supervisors))
    (println "John Dominguez now included?" john-included?)
    
    ;; Generate PNG
    (let [result (clojure.java.shell/sh "dot" "-Tpng" "-Gdpi=150" "full-org-chart.dot" "-o" "full-org-chart.png")]
      (if (= 0 (:exit result))
        (let [png-file (clojure.java.io/file "full-org-chart.png")]
          (println "✓ Full org chart PNG generated successfully! Size:" (.length png-file) "bytes"))
        (println "✗ PNG generation failed:" (:err result))))))
    
    (catch Exception e
      (println "Error:" (.getMessage e))
      (.printStackTrace e)))

  ;; Try different chart generation strategies for large datasets:

  ;; 1. Executive summary (top-level only)
  (tangle/save-executive-summary-chart test-extraction "executive-summary.png")

  ;; 2. Simplified chart (limited positions, no vacant roles)
  (tangle/save-simplified-org-chart test-extraction "simplified-chart.png"
                                    :max-positions 50
                                    :show-vacant false)

  ;; 3. Departmental charts (separate chart for each department)
  (tangle/save-hierarchical-charts test-extraction "dept-chart"
                                   :max-per-chart 25
                                   :group-by :unitcode)

  ;; 4. Filter by specific department if needed
  ;; (def sr-positions (tangle/filter-positions-by-department test-extraction ["SR"]))
  ;; (tangle/save-org-chart sr-positions "sr-department.png" :show-details false)

  (tap> test-extraction)

  ;; Performance options for large files:
  ;; 1. Minimal extraction (fastest)
  (def results-minimal (-> test-session
                           (rules/insert-all (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :minimal true)))
                           (rules/fire-rules)))

  ;; 2. Streaming for very large files (memory efficient)
  (def results-streaming (-> test-session
                             (rules/insert-all (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :streaming true)))
                             (rules/fire-rules)))

  ;; 3. Batch processing for memory-constrained environments
  (def results-batch (-> test-session
                         (rules/insert-all (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :batch-size 1000)))
                         (rules/fire-rules)))

  ;; Standard extraction with transducer optimizations
  (def results (-> test-session
                   (rules/insert-all (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx")))
                   (rules/fire-rules)))

  (tap> (rules/query results-streaming get-position-values))

  (tap> (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :minimal true))

  (tap> (pos/extract-positions (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx")))

  (tap> results)
  (tap> (xlsx/extract-data "resources/smaller Org Chart Data Analysis.xlsx"))
  :rcf)