(ns clara-org-chart.position
  (:require [clara.rules :refer :all]
            [clara-org-chart.xlsx :as xlsx]
            [data-types :as dt]
              [clara-org-chart.position-rules :as pos-rules])
    (:import (data_types Position))
  )

(defn debug-headers 
  "Debug function to see what headers are actually in the Excel file"
  [xlsx-data]
  (let [sheet (->> (:sheets xlsx-data)
                   (filter #(= (:sheet-name %) "ALL REGIONS"))
                   first)
        cells (:cells sheet)
        headers (->> cells
                     (filter #(= (:row %) 0))
                     (sort-by :col)
                     (map :value))]
    (println "=== EXCEL HEADERS DEBUG ===")
    (println "Found" (count headers) "headers:")
    (doseq [[idx header] (map-indexed vector headers)]
      (println (format "%2d: %s" idx (pr-str header))))
    headers))


(defn extract-positions [xlsx-data]
  (let [sheet (->> (:sheets xlsx-data)
                   (filter #(= (:sheet-name %) "OrgChart_HQ03"))
                   first)
        cells (:cells sheet)
        ;; Get headers with their column positions
        header-cells (->> cells
                          (filter #(= (:row %) 0))
                          (sort-by :col))
        max-col (apply max (map :col header-cells))
        ;; Create a complete header mapping with all columns 0 to max-col
        complete-headers (into {} (map (fn [cell] [(:col cell) (:value cell)]) header-cells))
        headers (mapv #(get complete-headers % nil) (range 0 (inc max-col)))
        
        rows (->> cells
                  (remove #(= (:row %) 0))
                  (group-by :row))]
    (->> rows
         (sort-by first) ; deterministic ordering by row number
         (map (fn [[row-num row-cells]]
                ;; Create a complete row mapping with all columns 0 to max-col
                (let [row-cell-map (into {} (map (fn [cell] [(:col cell) (:value cell)]) row-cells))
                      complete-row-data (mapv #(get row-cell-map % nil) (range 0 (inc max-col)))
                      row-map (zipmap headers complete-row-data)]
                  (dt/map->Position
                   {;; original sheet row index (header row is 0)
                    :row-num row-num
                    :position (get row-map "Position")
                    :title (get row-map "Title")
                    :current-employee (get row-map "Current Employee")
                    :reports-to-position (get row-map "Position Reports To Position")
                    :dotted-line-reports-to-position (get row-map "Position Reports To Position (Dotted Line)")
                    :city (get row-map "City")
                    :comments-notes (get row-map "Comments/Notes")
                    :dotted-reports (get row-map "Dotted_Reports")
                    :agencycode (get row-map "AGENCY")
                    :unitcode (get row-map "UNIT")
                    :classcode (get row-map "CLASS")
                    :serialnumber (get row-map "SERIAL")
                    ;; :unique-flag (get row-map "Unique_Flag")
                    :time-base (get row-map "Time_Base")
                    :tenure (get row-map "Tenure")
                    :region (get row-map "Region")
                    :position-weight 0
                    :total-subordinates nil})))))))

(defn get-position-weight
  "Get the position weight, defaulting to 1.0 if not set or invalid"
  [position]
  (:position-weight position 1.0))

(defn calculate-subordinate-counts-simple
  "Simple recursive version with cycle detection - safer for debugging.
  Now accounts for tb-adjustment to weight partial positions."
  [positions]
  (let [reports-map (group-by :reports-to-position positions)
        valid-positions (filter #(and (:position %) 
                                     (seq (str (:position %))))
                               positions)]
    
    (letfn [(count-subordinates [position-id visited]
              ;; Prevent cycles by tracking visited positions
              (if (contains? visited position-id)
                0.0
                (let [direct-reports (get reports-map position-id [])
                      ;; Calculate weighted direct count based on tb-adjustment
                      direct-count (reduce + 0.0 (map get-position-weight direct-reports))
                      new-visited (conj visited position-id)
                      indirect-count (reduce + 0.0 
                                            (map #(count-subordinates (:position %) new-visited) 
                                                 direct-reports))
                      ;; If there are direct reports, include self in count, otherwise count stays 0
                      self-count (if (> direct-count 0) 1.0 0.0)]
                  (+ direct-count indirect-count self-count))))]
      
      (into {} 
            (map (fn [pos] 
                   [(:position pos) (count-subordinates (:position pos) #{})]) 
                 valid-positions)))))

(defn calculate-subordinate-counts-optimized
  "Calculate subordinate counts using bottom-up approach with safety guards.
  Returns a map of position-id -> total subordinate count (including indirect reports)."
  [positions]
  (let [valid-positions (filter #(and (:position %) 
                                     (seq (str (:position %))))
                               positions)
        position-map (into {} (map (juxt :position identity) valid-positions))
        reports-map (group-by :reports-to-position valid-positions)
        all-position-ids (set (map :position valid-positions))
        
        ;; Calculate data quality metrics upfront
        self-reporters (filter #(= (:position %) (:reports-to-position %)) valid-positions)
        broken-chains (filter #(and (:reports-to-position %) 
                                   (seq (str (:reports-to-position %)))
                                   (not (contains? all-position-ids (:reports-to-position %))))
                             valid-positions)
        
        ;; Find leaf nodes (positions with no direct reports)
        leaf-positions (filter #(empty? (get reports-map (:position %) [])) valid-positions)]
    
    (if (empty? leaf-positions)
      ;; No leaf nodes found - possible circular references, use simple version
      (do
        (println "Warning: No leaf nodes found, using simple algorithm")
        (calculate-subordinate-counts-simple positions))
      
      ;; Process bottom-up with safety guards
      (loop [processed #{}
             counts {}
             to-process (set (map :position leaf-positions))
             iterations 0
             max-iterations (* 3 (count valid-positions))] ; Safety limit
        
        (cond
          ;; Safety: prevent infinite loops
          (> iterations max-iterations)
          (let [skipped (- (count valid-positions) (count processed))
                self-reporters (count (filter #(= (:position %) (:reports-to-position %)) valid-positions))
                broken-chains (count (filter #(and (:reports-to-position %) 
                                                  (not (contains? all-position-ids (:reports-to-position %))))
                                            valid-positions))]
            (println "Subordinate count calculation completed:")
            (println " - Processed:" (count processed) "positions")
            (println " - Skipped:" skipped "positions due to data issues")
            (when (> (+ self-reporters broken-chains) 0)
              (println " - Data issues found: " self-reporters "self-reporters," broken-chains "broken chains"))
            counts)
          
          ;; Success: all done
          (empty? to-process)
          (do
            ;; Always show summary when completed
            (let [total-issues (+ (count self-reporters) (count broken-chains))]
              (when (> total-issues 0)
                (println "Subordinate count calculation completed:")
                (println " - Processed:" (count processed) "of" (count valid-positions) "positions")
                (println " - Data issues handled:" total-issues "(" (count self-reporters) "self-reporters," (count broken-chains) "broken chains)")))
            counts)
          
          ;; Continue processing
          :else
          ;; Process ALL ready positions in this iteration instead of just one
          (let [ready-positions (filter (fn [pos-id]
                                          (let [direct-reports (get reports-map pos-id [])
                                                subordinate-ids (map :position direct-reports)]
                                            (every? #(contains? processed %) subordinate-ids)))
                                        to-process)]
            
            (if (empty? ready-positions)
              ;; No positions are ready - this shouldn't happen in a valid hierarchy
              ;; Process the first one anyway with partial data to prevent infinite loops
              (let [current (first to-process)
                    remaining (disj to-process current)
                    direct-reports (get reports-map current [])
                    subordinate-ids (map :position direct-reports)
                    ;; Use weighted count based on tb-adjustment
                    direct-count (reduce + 0.0 (map get-position-weight direct-reports))
                    partial-indirect (reduce + 0.0 (keep #(get counts %) subordinate-ids))
                    ;; If there are direct reports, include self in count, otherwise count stays 0
                    self-count (if (> direct-count 0) 1.0 0.0)
                    total-count (+ direct-count partial-indirect self-count)
                    new-counts (assoc counts current total-count)
                    new-processed (conj processed current)]
                (recur new-processed new-counts remaining (inc iterations) max-iterations))
              
              ;; Process all ready positions
              (let [new-counts (reduce (fn [counts-acc pos-id]
                                         (let [direct-reports (get reports-map pos-id [])
                                               subordinate-ids (map :position direct-reports)
                                               ;; Use weighted count based on tb-adjustment
                                               direct-count (reduce + 0.0 (map get-position-weight direct-reports))
                                               indirect-count (reduce + 0.0 (map #(get counts-acc % 0.0) subordinate-ids))
                                               ;; If there are direct reports, include self in count, otherwise count stays 0
                                               self-count (if (> direct-count 0) 1.0 0.0)
                                               total-count (+ direct-count indirect-count self-count)]
                                           (assoc counts-acc pos-id total-count)))
                                       counts ready-positions)
                    new-processed (into processed ready-positions)
                    
                    ;; Add supervisors to queue
                    supervisors-to-add (reduce (fn [acc pos-id]
                                                 (let [supervisor-pos (get position-map pos-id)
                                                       supervisor-id (:reports-to-position supervisor-pos)]
                                                   (if (and supervisor-id 
                                                           (contains? all-position-ids supervisor-id)
                                                           (not (contains? new-processed supervisor-id))
                                                           (not (contains? acc supervisor-id))
                                                           (not= supervisor-id pos-id))
                                                     (conj acc supervisor-id)
                                                     acc)))
                                               #{} ready-positions)
                    remaining-positions (apply disj to-process ready-positions)
                    next-to-process (into remaining-positions supervisors-to-add)]
                
                (recur new-processed new-counts next-to-process (inc iterations) max-iterations)))))))))

(defn add-subordinate-counts
  "Add subordinate counts to position records using optimized bottom-up calculation.
  Returns positions with :total-subordinates field populated."
  [positions]
  (let [counts (calculate-subordinate-counts-optimized positions)]
    (map (fn [pos] 
           (assoc pos :total-subordinates (get counts (:position pos) 0)))
         positions)))

(defn add-subordinate-counts-quiet
  "Same as add-subordinate-counts but uses simple algorithm to avoid warnings."
  [positions]
  (let [counts (calculate-subordinate-counts-simple positions)]
    (map (fn [pos] 
           (assoc pos :total-subordinates (get counts (:position pos) 0)))
         positions)))

(defn debug-subordinate-calculation
  "Debug function to trace the subordinate calculation for specific positions."
  [positions target-positions]
  (let [valid-positions (filter #(and (:position %) 
                                     (seq (str (:position %))))
                               positions)
        reports-map (group-by :reports-to-position valid-positions)
        position-map (into {} (map (juxt :position identity) valid-positions))]
    
    (println "\n=== DEBUGGING SUBORDINATE CALCULATION ===")
    (doseq [target target-positions]
      (let [pos (get position-map target)
            direct-reports (get reports-map target [])
            direct-count (count direct-reports)]
        
        (println "\nPosition:" target)
        (println " Title:" (:title pos))
        (println " Direct reports count:" direct-count)
        (println " Weighted direct reports count:" (reduce + 0.0 (map get-position-weight direct-reports)))
        (println " Direct report positions:")
        (doseq [report direct-reports]
          (let [weight (get-position-weight report)]
            (println "  -" (:position report) "(" (:title report) ")" 
                     "weight:" weight 
                     (when (:tenure report) (str "tb-adj:" (:tenure report))))))
        
        ;; Calculate using simple algorithm for comparison  
        (let [simple-count (get (calculate-subordinate-counts-simple [pos]) target 0.0)
              optimized-count (get (calculate-subordinate-counts-optimized [pos]) target 0.0)]
          (println " Simple algorithm result:" simple-count)
          (println " Optimized algorithm result:" optimized-count))))))

(defn verify-hierarchy-consistency
  "Check for specific inconsistencies in supervisor-subordinate counts."
  [positions-with-counts]
  (let [reports-map (group-by :reports-to-position positions-with-counts)]
    
    (println "\n=== HIERARCHY CONSISTENCY CHECK ===")
    (let [inconsistencies 
          (for [pos positions-with-counts
                :let [direct-reports (get reports-map (:position pos) [])
                      max-subordinate-count (if (empty? direct-reports) 
                                              0 
                                              (apply max (map :total-subordinates direct-reports)))
                      pos-count (:total-subordinates pos)]
                :when (and (> max-subordinate-count 0)
                          (< pos-count max-subordinate-count))]
            {:supervisor pos
             :max-subordinate-count max-subordinate-count
             :supervisor-count pos-count
             :direct-reports direct-reports})]
      
      (if (empty? inconsistencies)
        (println "No hierarchy inconsistencies found!")
        (do
          (println "Found" (count inconsistencies) "hierarchy inconsistencies:")
          (doseq [{:keys [supervisor max-subordinate-count supervisor-count direct-reports]} (take 5 inconsistencies)]
            (println " Supervisor:" (:position supervisor) "(" (:title supervisor) ")")
            (println "  Has count:" supervisor-count "but subordinate has:" max-subordinate-count)
            (println "  Direct reports with high counts:")
            (doseq [report (filter #(= (:total-subordinates %) max-subordinate-count) direct-reports)]
              (println "   -" (:position report) "has" (:total-subordinates report) "subordinates"))
            (println)))))))

(defn diagnose-hierarchy-issues
  "Analyze the organizational hierarchy to identify data quality issues."
  [positions]
  (let [valid-positions (filter #(and (:position %) 
                                     (seq (str (:position %))))
                               positions)
        position-ids (set (map :position valid-positions))
        reports-map (group-by :reports-to-position valid-positions)]
    
    (println "\n=== HIERARCHY DIAGNOSTIC REPORT ===")
    (println "Total positions:" (count valid-positions))
    
    ;; Find positions that report to non-existent positions
    (let [broken-chains (filter #(and (:reports-to-position %) 
                                     (not (contains? position-ids (:reports-to-position %))))
                               valid-positions)]
      (println "Positions with missing supervisors:" (count broken-chains))
      (when (> (count broken-chains) 0)
        (println "Examples of broken reporting chains:")
        (doseq [pos (take 5 broken-chains)]
          (println " " (:position pos) "reports to non-existent" (:reports-to-position pos)))))
    
    ;; Find potential circular references
    (let [self-reporters (filter #(= (:position %) (:reports-to-position %)) valid-positions)]
      (println "Self-reporting positions:" (count self-reporters))
      (doseq [pos self-reporters]
        (println " " (:position pos) "reports to itself")))
    
    ;; Find leaf nodes (no direct reports)
    (let [leaf-positions (filter #(empty? (get reports-map (:position %) [])) valid-positions)]
      (println "Leaf positions (no direct reports):" (count leaf-positions)))
    
    ;; Find root positions (no supervisor)
    (let [root-positions (filter #(or (nil? (:reports-to-position %))
                                     (empty? (str (:reports-to-position %))))
                                valid-positions)]
      (println "Root positions (no supervisor):" (count root-positions))
      (when (> (count root-positions) 0)
        (println "Root positions:")
        (doseq [pos (take 10 root-positions)]
          (println " " (:position pos) "-" (:title pos)))))
    
    ;; Return summary for programmatic use
    {:total-positions (count valid-positions)
     :broken-chains (count (filter #(and (:reports-to-position %) 
                                        (not (contains? position-ids (:reports-to-position %))))
                                  valid-positions))
     :self-reporters (count (filter #(= (:position %) (:reports-to-position %)) valid-positions))
     :leaf-positions (count (filter #(empty? (get reports-map (:position %) [])) valid-positions))
     :root-positions (count (filter #(or (nil? (:reports-to-position %))
                                        (empty? (str (:reports-to-position %))))
                                   valid-positions))}))

(defn calculated-position->position
  "Convert CalculatedPosition records back to Position maps for compatibility"
  [calculated-positions]
  (map (fn [ep]
         (dt/map->Position
          (select-keys ep [:row-num :position :title :current-employee :reports-to-position
                          :dotted-line-reports-to-position :city :comments-notes :dotted-reports
                          :agencycode :unitcode :classcode :serialnumber :unique-flag 
                          :time-base :tenure :region :position-weight :total-subordinates])))
       calculated-positions))

(defn extract-positions-with-counts
  "Extract positions from xlsx-data and calculate subordinate counts.
  This is the recommended function to use for complete position data."
  [xlsx-data]
  (-> xlsx-data
      extract-positions
      pos-rules/process-positions
      calculated-position->position
      add-subordinate-counts
      ))


(comment
  (tap>
 (pos-rules/process-positions (extract-positions (xlsx/extract-data "resources/OrgChart_HQ03.xlsx"))))

  
  (tap> (extract-positions-with-counts (xlsx/extract-data "resources/OrgChart_HQ03.xlsx")))


  
  :rcf)