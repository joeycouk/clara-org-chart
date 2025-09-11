(ns org-tangle
  (:require [tangle.core :refer [graph->dot dot->image dot->svg]]
            [clojure.java.io :refer [file copy]]
            [clojure.java.shell :as shell]
            [clojure.string :as str]
            [clojure.set :as set]
            [clara-org-chart.position :as pos]
            [clara-org-chart.xlsx :as xlsx]))



(defn filter-positions-by-codes
  "Return all positions where :position or :reports-to-position matches any code in codes."
  [positions codes]
  (let [code-set (set codes)]
    (filter #(or (contains? code-set (:position %))
                 (contains? code-set (:reports-to-position %)))
            positions)))

(defn filter-positions-by-exact-codes
  "Return only positions where :position exactly matches one of the codes (no automatic inclusion of reports)."
  [positions codes]
  (let [code-set (set codes)]
    (filter #(contains? code-set (:position %))
            positions)))

(defn position->node-id
  "Convert a position string to a valid node ID"
  [position]
  (when position
    (let [position-str (str position)  ; Convert to string first
          cleaned (-> position-str
                      (str/replace #"[^a-zA-Z0-9_-]" "_")
                      (str/replace #"_{2,}" "_")
                      (str/trim))]
      (when (not (str/blank? cleaned))
        cleaned))))

(defn format-employee-name
  "Format employee name for display"
  [name]
  (if (or (nil? name) (= name "Vacant") (str/blank? name))
    "VACANT"
    (str/upper-case name)))

(defn create-position-node
  "Create a node representation for a position"
  [position & {:keys [show-codes] :or {show-codes false}}]
  (let [node-id (position->node-id (:position position))
        employee-name (format-employee-name (:current-employee position))
        title (:title position)
        position-code (:position position)
        
        ;; Create improved label with employee, position, and title
        label-parts (cond-> []
                      ;; Always include employee name
                      true (conj employee-name)
                      ;; Always include position code  
                      position-code (conj (str position-code))
                      ;; Always include title
                      title (conj title)
                      ;; Optionally include additional details
                      (and show-codes (:agencycode position)) (conj (str "Agency: " (:agencycode position)))
                      (and show-codes (:unitcode position)) (conj (str "Unit: " (:unitcode position))))
        
        label (str/join "\\n" label-parts)
        
        ;; Check management status for styling
        has-subordinates-in-org (and (:direct-subordinates position) (> (:direct-subordinates position) 0))
        is-ra-position (= (:time-base position) "RA")
        
        ;; Determine node styling based on position type and management status
        base-attrs (cond
                     (= employee-name "VACANT") 
                     {:fillcolor "lightgray" :fontcolor "red"}
                     
                    ;;  (str/includes? (str/upper-case (or title "")) "DIRECTOR")
                    ;;  {:fillcolor "lightblue" :fontweight "bold"}
                     
                    ;;  (str/includes? (str/upper-case (or title "")) "CHIEF")
                    ;;  {:fillcolor "lightgreen" :fontweight "bold"}
                     
                    ;;  (str/includes? (str/upper-case (or title "")) "DEPUTY")
                    ;;  {:fillcolor "lightyellow"}
                     
                    ;;  (str/includes? (str/upper-case (or title "")) "ASSISTANT")
                    ;;  {:fillcolor "lightcyan"}
                     
                    ;;  (str/includes? (str/upper-case (or title "")) "CAPTAIN")
                    ;;  {:fillcolor "wheat"}
                     
                     :else
                     {:fillcolor "white"})
        
        ;; Determine shape and style based on management status and time-base
        shape-attrs (cond
                      ;; Special handling for vacant positions
                      (= employee-name "VACANT")
                      {:shape "box" :style "filled,dashed"}
                      
                      ;; Special styling for RA positions (double dotted squares - inner and outer)
                      (and is-ra-position has-subordinates-in-org)
                      {:shape "box" :style "filled,dotted" :penwidth 2 :peripheries 2}
                      
                      (and is-ra-position (not has-subordinates-in-org))
                      {:shape "box" :style "filled,rounded,dotted" :penwidth 2 :peripheries 2}
                      
                      ;; Regular styling for non-RA positions
                      ;; Square boxes for people with direct-subordinates > 0 
                      has-subordinates-in-org
                      {:shape "box" :style "filled"}
                      
                      ;; Rounded squares for people with no organizational subordinates
                      :else
                      {:shape "box" :style "filled,rounded"})
        
        node-attrs (merge base-attrs shape-attrs)]
    
    (merge {:id node-id
            :label label}
           node-attrs)))

(defn find-root-positions
  "Find positions that don't report to anyone (organization roots)"
  [positions]
  (let [all-positions (set (map :position positions))
        reporting-positions (set (map :reports-to-position positions))
        ;; Find positions that exist but are not in anyone's reports-to field
        roots (filter (fn [pos]
                        (and (:position pos)
                             (not (contains? reporting-positions (:position pos)))))
                      positions)]
    (if (empty? roots)
      ;; If no clear root found, find positions with missing/invalid reports-to
      (filter (fn [pos]
                (or (nil? (:reports-to-position pos))
                    (str/blank? (:reports-to-position pos))
                    (not (contains? all-positions (:reports-to-position pos)))))
              positions)
      roots)))

(defn sort-positions-hierarchically
  "Sort positions with roots first, then by hierarchy depth"
  [positions]
  (let [roots (find-root-positions positions)
        root-ids (set (map :position roots))
        position-map (into {} (map (fn [p] [(:position p) p]) positions))
        
        ;; Function to calculate depth from root
        depth-map (atom {})
        
        calculate-depth (fn calculate-depth [pos-id visited]
                          (if (contains? visited pos-id)
                            1000 ; Circular reference, put at bottom
                            (if (contains? root-ids pos-id)
                              0
                              (if-let [pos (get position-map pos-id)]
                                (let [parent-id (:reports-to-position pos)]
                                  (if (and parent-id (contains? position-map parent-id))
                                    (+ 1 (calculate-depth parent-id (conj visited pos-id)))
                                    500)) ; Missing parent, put towards bottom
                                1000)))) ; Position not found
        
        ;; Calculate depths for all positions
        _ (doseq [pos positions]
            (swap! depth-map assoc (:position pos) 
                   (calculate-depth (:position pos) #{})))
        
        ;; Sort by depth, then by title importance, then by name
        title-priority (fn [title]
                         (let [title-upper (str/upper-case (or title ""))]
                           (cond
                             (str/includes? title-upper "DIRECTOR") 1
                             (str/includes? title-upper "CHIEF") 2
                             (str/includes? title-upper "DEPUTY") 3
                             (str/includes? title-upper "ASSISTANT") 4
                             (str/includes? title-upper "CAPTAIN") 5
                             :else 6)))]
    
    (sort-by (fn [pos]
               [(get @depth-map (:position pos) 1000)
                (title-priority (:title pos))
                (:current-employee pos)])
             positions)))

(defn create-reporting-edges
  "Create edges representing reporting relationships (supervisor -> subordinate)"
  [positions]
  (for [position positions
        :let [subordinate-id (position->node-id (:position position))
              supervisor-id (position->node-id (:reports-to-position position))]
        :when (and subordinate-id supervisor-id 
                   (not (str/blank? subordinate-id))
                   (not (str/blank? supervisor-id)))]
    ;; Note: This creates edges FROM supervisor TO subordinate (downward flow)
    [supervisor-id subordinate-id {:color "black" :arrowhead "normal"}]))

(defn create-dotted-line-edges
  "Create dotted line edges representing matrix/dotted line relationships"
  [positions]
  (for [position positions
        :let [position-id (position->node-id (:position position))
              dotted-manager-id (position->node-id (:dotted-line-reports-to-position position))]
        :when (and position-id dotted-manager-id 
                   (not (str/blank? position-id))
                   (not (str/blank? dotted-manager-id)))]
    [position-id dotted-manager-id {:style "dashed" :color "gray"}]))

(defn create-referenced-supervisor-nodes
  "Create nodes for supervisors that are referenced but not in the main position list"
  [positions]
  (let [existing-position-ids (set (map :position positions))
        referenced-supervisor-ids (set (filter some? (map :reports-to-position positions)))
        missing-supervisor-ids (clojure.set/difference referenced-supervisor-ids existing-position-ids)]
    (for [supervisor-id missing-supervisor-ids
          :when (not (str/blank? supervisor-id))]
      {:id (position->node-id supervisor-id)
       :label (str supervisor-id "\n(Referenced)")
       :fillcolor "yellow"
       :fontcolor "black"
       :style "filled"
       :shape "box"
       :penwidth 2})))

;; Error visualization functions
(defn create-error-node
  "Create a node representation for an org chart error"
  [error]
  (let [error-id (str "error_" (hash error))
        missing-pos (:missing-position-code error)
        extra-pos (:extra-position-code error)
        page (:page error)
        description (:description error)
        
        ;; Create error label with more prominent styling
        label-parts (cond-> []
                      missing-pos (conj (str "Missing: " missing-pos))
                      extra-pos (conj (str "Extra: " extra-pos))
                      description (conj description)
                      page (conj (str "Page: " page)))
        
        label (str/join "\\n" label-parts)]
    
    {:id error-id
     :label label
     :fillcolor "red"
     :fontcolor "white" 
     :style "filled"
     :shape "box"
     :penwidth 2
     :fontsize 10}))

(defn create-error-anchor
  "Create an invisible anchor node to position errors at the bottom"
  [errors]
  (when (seq errors)
    [{:id "error_anchor_top"
      :label ""
      :shape "point"
      :style "invis"
      :width 0
      :height 0}
     {:id "error_anchor_bottom"
      :label ""
      :shape "point"
      :style "invis"
      :width 0
      :height 0}]))  ; Create two anchors for better control

(defn create-error-edges
  "Create edges for errors: invisible ranking edges and visible relation edges"
  [errors positions root-positions]
  (let [position-id-map (into {} (map (juxt :position #(position->node-id (:position %))) positions))
        
        ;; Create invisible ranking edges to force errors to bottom
        ranking-edges (when (seq errors)
                        (let [;; Connect all root positions to top anchor
                              root-to-top (for [root root-positions
                                               :let [root-id (position->node-id (:position root))]]
                                           [root-id "error_anchor_top" {:style "invis" :constraint "true" :weight 100}])
                              ;; Connect top anchor to bottom anchor  
                              anchor-chain [["error_anchor_top" "error_anchor_bottom" {:style "invis" :constraint "true" :weight 100}]]
                              ;; Connect bottom anchor to all errors
                              bottom-to-errors (for [error errors
                                                    :let [error-id (str "error_" (hash error))]]
                                                ["error_anchor_bottom" error-id {:style "invis" :constraint "true" :weight 100}])]
                          (concat root-to-top anchor-chain bottom-to-errors)))
        
        ;; Visible edges from errors to related positions (only for extra positions that exist)
        relation-edges (for [error errors
                            :let [error-id (str "error_" (hash error))
                                  extra-pos (:extra-position-code error)]
                            :when (and extra-pos (get position-id-map extra-pos))
                            :let [target-id (get position-id-map extra-pos)]]
                        [error-id target-id {:color "red" :style "dashed" :arrowhead "diamond" :constraint "false" :weight 1}])]
    (concat ranking-edges relation-edges)))

;; Data filtering and optimization functions for large org charts
(defn filter-positions-by-level
  "Filter positions to show only certain organizational levels"
  [positions & {:keys [max-levels include-titles exclude-vacant]
                :or {max-levels 3 include-titles #{"Director" "Chief" "Deputy" "Assistant" "Captain" "Manager"} exclude-vacant false}}]
  (let [title-filter (fn [pos] 
                       (some #(str/includes? (str/upper-case (or (:title pos) "")) (str/upper-case %)) 
                             include-titles))
        vacant-filter (fn [pos]
                        (if exclude-vacant
                          (not= (str/upper-case (or (:current-employee pos) "")) "VACANT")
                          true))]
    (->> positions
         (filter title-filter)
         (filter vacant-filter)
         (take (* max-levels 50))))) ; Reasonable limit per level

(defn filter-positions-by-department
  "Filter positions by department/unit code"
  [positions department-codes]
  (let [dept-set (set (map str/upper-case department-codes))]
    (filter #(contains? dept-set (str/upper-case (or (:unitcode %) ""))) positions)))

(defn filter-positions-by-region
  "Filter positions by region/location"
  [positions regions]
  (let [region-set (set (map str/upper-case regions))]
    (filter #(or (contains? region-set (str/upper-case (or (:city %) "")))
                 (contains? region-set (str/upper-case (or (:region %) "")))) positions)))

(defn create-hierarchical-subsets
  "Create smaller org chart subsets based on hierarchy"
  [positions & {:keys [max-per-chart group-by-field]
                :or {max-per-chart 50 group-by-field :unitcode}}]
  (let [grouped (group-by group-by-field positions)]
    (map (fn [[group-key group-positions]]
           {:group group-key
            :positions (take max-per-chart group-positions)
            :total-count (count group-positions)})
         grouped)))

(defn simplify-large-chart
  "Simplify a large org chart by reducing detail and connections"
  [positions & {:keys [max-positions show-vacant-positions simplify-titles]
                :or {max-positions 100 show-vacant-positions false simplify-titles true}}]
  (let [filtered-positions (cond->> positions
                             (not show-vacant-positions)
                             (remove #(= (str/upper-case (or (:current-employee %) "")) "VACANT"))
                             
                             true
                             (take max-positions))
        
        simplified-positions (if simplify-titles
                               (map (fn [pos]
                                      (update pos :title #(-> %
                                                               (str/replace #"with Differential|Paramedic|\(A\)|\{B\}" "")
                                                               (str/trim))))
                                    filtered-positions)
                               filtered-positions)]
    simplified-positions))

(defn is-valid-position-id?
  "Check if a string looks like a valid position ID"
  [id]
  (and (string? id)
       (seq id)  ; not empty
       (re-matches #"^\d{3}-.*" id)  ; starts with 3 digits and a hyphen
       (not (re-find #"(?i)(page|unit|no id|missing)" id))))  ; doesn't contain common invalid text

(defn find-missing-supervisor-positions
  "Find positions that are referenced as supervisors but not in the dataset"
  [positions]
  (let [existing-positions (set (map :position positions))
        referenced-supervisors (set (filter some? (map :reports-to-position positions)))
        missing-supervisor-ids (set/difference referenced-supervisors existing-positions)
        ;; Only consider valid-looking position IDs
        valid-missing-ids (filter is-valid-position-id? missing-supervisor-ids)]
    valid-missing-ids))

(defn analyze-missing-supervisors
  "Analyze which supervisor positions are missing from the dataset"
  [positions]
  (let [existing-positions (set (map :position positions))
        referenced-supervisors (set (filter some? (map :reports-to-position positions)))
        missing-supervisor-ids (set/difference referenced-supervisors existing-positions)
        ;; Filter for what look like valid position IDs (start with digits and contain hyphens)
        valid-looking-ids (filter #(and (string? %)
                                        (seq %)
                                        (re-matches #"^\d{3}-.*" %)) missing-supervisor-ids)
        invalid-ids (set/difference missing-supervisor-ids (set valid-looking-ids))]
    {:total-positions (count positions)
     :existing-position-count (count existing-positions)
     :referenced-supervisors-count (count referenced-supervisors)
     :missing-supervisor-count (count missing-supervisor-ids)
     :valid-looking-missing-ids valid-looking-ids
     :invalid-supervisor-refs invalid-ids
     :sample-invalid (take 10 invalid-ids)}))

(defn debug-org-chart-data
  "Debug function to understand the data structure"
  []
  (let [data (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx"))
        analysis (analyze-missing-supervisors data)]
    (println "=== ORG CHART DATA ANALYSIS ===")
    (println "Total positions:" (:total-positions analysis))
    (println "Referenced supervisors:" (:referenced-supervisors-count analysis))
    (println "Missing supervisors:" (:missing-supervisor-count analysis))
    (println "Valid-looking missing IDs:" (count (:valid-looking-missing-ids analysis)))
    (println "Invalid supervisor references:" (count (:invalid-supervisor-refs analysis)))
    (println "Sample invalid references:")
    (doseq [id (:sample-invalid analysis)]
      (println " -" (pr-str id)))
    (when (seq (:valid-looking-missing-ids analysis))
      (println "Valid-looking missing supervisor IDs:")
      (doseq [id (take 5 (:valid-looking-missing-ids analysis))]
        (println " -" id)))
    analysis))

(defn create-placeholder-position
  "Create a placeholder position for missing supervisors"
  [position-id]
  {:position position-id
   :current-employee "External Position"
   :title "Outside This Dataset"
   :reports-to-position nil
   :dotted-line-reports-to-position nil
   :placeholder true})

(defn ensure-complete-hierarchy
  "Ensure all referenced positions are included in the dataset"
  [positions]
  (let [missing-ids (find-missing-supervisor-positions positions)
        placeholder-positions (map create-placeholder-position missing-ids)]
    (concat positions placeholder-positions)))

(defn positions->org-chart
  "Convert a collection of Position records to org chart nodes and edges"
  [positions & {:keys [show-codes include-dotted-lines strict-filter errors] 
                :or {show-codes false include-dotted-lines true strict-filter false errors []}}]
  (let [;; Optionally ensure we have all referenced positions
        complete-positions (if strict-filter 
                             positions  ; Use only the provided positions
                             (ensure-complete-hierarchy positions))
        ;; Sort positions hierarchically with roots first
        sorted-positions (sort-positions-hierarchically complete-positions)
        position-nodes (map #(create-position-node % :show-codes show-codes) sorted-positions)
        ;; When using strict filter, create nodes for referenced supervisors
        referenced-supervisor-nodes (if strict-filter 
                                       (create-referenced-supervisor-nodes positions)
                                       [])
        error-nodes (map create-error-node errors)
        error-anchors (create-error-anchor errors)
        all-nodes (concat position-nodes referenced-supervisor-nodes error-nodes (or error-anchors []))
        solid-edges (create-reporting-edges sorted-positions)
        dotted-edges (if include-dotted-lines (create-dotted-line-edges sorted-positions) [])
        root-positions (find-root-positions sorted-positions)
        error-edges (create-error-edges errors sorted-positions root-positions)
        all-edges (concat solid-edges dotted-edges error-edges)]
    {:nodes all-nodes
     :edges all-edges
     :root-positions root-positions}))

(defn generate-large-org-chart-dot
  "Generate DOT notation optimized for large org charts (500+ positions)"
  [positions & {:keys [show-codes include-dotted-lines rankdir strict-filter errors title]
                :or {show-codes false include-dotted-lines false rankdir "TB" strict-filter false errors [] title nil}}]
  (let [{:keys [nodes edges]} (positions->org-chart positions 
                                                                :show-codes show-codes
                                                                :include-dotted-lines include-dotted-lines
                                                                :strict-filter strict-filter
                                                                :errors errors)
        graph-config (cond-> {:rankdir rankdir          ; Top-down or left-right layout
                              :dpi 150                  ; Fixed resolution
                              :splines "ortho"          ; Orthogonal edges
                              :nodesep 1.0              ; FIXED spacing between nodes
                              :ranksep 1.5              ; FIXED spacing between levels
                              :concentrate true         ; Reduce edge crossings
                              :newrank true             ; Better ranking algorithm
                              :overlap false            ; Prevent node overlap
                              :packmode "node"}         ; Pack by nodes, not area
                       title (assoc :label title
                                   :labelloc "t"        ; Position label at top
                                   :fontsize 16         ; Larger font for title
                                   :fontname "Arial Bold"))] ; Bold title font
    (graph->dot nodes edges 
                {:directed? true
                 :graph graph-config
                 :node {:fontname "Arial Bold"     ; Bold font for readability
                        :fontsize 12               ; FIXED font size
                        :margin 0.2                ; FIXED margin
                        :width 2.5                 ; FIXED box width
                        :height 1.0                ; FIXED box height
                        :style "filled,rounded"    ; Filled with rounded corners
                        :penwidth 1.5}             ; FIXED border thickness
                 :edge {:fontname "Arial"
                        :fontsize 9                ; FIXED edge font size
                        :arrowsize 0.8             ; FIXED arrow size
                        :penwidth 1.5}             ; FIXED edge thickness
                 :node->id :id
                 :node->descriptor (fn [node] (dissoc node :id))})))

(defn generate-org-chart-dot
  "Generate DOT notation for an org chart with proper hierarchical layout"
  [positions & {:keys [show-codes include-dotted-lines rankdir strict-filter errors title]
                :or {show-codes false include-dotted-lines true rankdir "TB" strict-filter false errors [] title nil}}]
  (let [{:keys [nodes edges]} (positions->org-chart positions 
                                                                :show-codes show-codes
                                                                :include-dotted-lines include-dotted-lines
                                                                :strict-filter strict-filter
                                                                :errors errors)
        graph-config (cond-> {:rankdir rankdir          ; Top-down layout
                              :dpi 150                  ; High resolution
                              :splines "ortho"          ; Orthogonal edges
                              :nodesep 0.8              ; Space between nodes at same level
                              :ranksep 1.2              ; Space between hierarchy levels
                              :concentrate true         ; Reduce edge crossings
                              :newrank true             ; Better ranking algorithm
                              :compound true}           ; Allow compound layouts
                       title (assoc :label title
                                   :labelloc "t"        ; Position label at top
                                   :fontsize 16         ; Larger font for title
                                   :fontname "Arial Bold"))] ; Bold title font
    (graph->dot nodes edges 
                {:directed? true
                 :graph graph-config
                 :node {:fontname "Arial"
                        :fontsize 11
                        :margin 0.2
                        :width 2.5                ; Minimum box width
                        :height 1.0}              ; Minimum box height
                 :edge {:fontname "Arial"
                        :fontsize 9
                        :arrowsize 0.8}
                 :node->id :id
                 :node->descriptor (fn [node] (dissoc node :id))})))

;; Optimized chart generation functions for large datasets
(defn select-connected-positions
  "Select positions ensuring supervisors are included when subordinates are selected"
  [positions max-positions]
  (let [position-map (into {} (map (juxt :position identity) positions))
        
        ;; Find positions that don't have their supervisor in the dataset (pseudo-roots)
        existing-position-ids (set (map :position positions))
        pseudo-roots (filter #(let [supervisor-id (:reports-to-position %)]
                               (or (nil? supervisor-id)
                                   (empty? supervisor-id)
                                   (not (contains? existing-position-ids supervisor-id))))
                             positions)
        
        ;; If no pseudo-roots found, just take the first few positions
        starting-positions (if (seq pseudo-roots)
                            pseudo-roots
                            (take 10 positions))
        
        ;; Function to collect a position and its chain of supervisors and subordinates
        collect-connected (fn collect [pos collected depth]
                           (if (or (nil? pos) 
                                   (contains? collected (:position pos))
                                   (> depth 10)) ; prevent infinite loops
                             collected
                             (let [pos-id (:position pos)
                                   new-collected (conj collected pos-id)
                                   
                                   ;; Get supervisor
                                   supervisor-id (:reports-to-position pos)
                                   supervisor (when (and supervisor-id (contains? position-map supervisor-id))
                                               (get position-map supervisor-id))
                                   
                                   ;; Get direct reports
                                   subordinates (filter #(= (:reports-to-position %) pos-id) positions)
                                   
                                   ;; Collect supervisor
                                   with-supervisor (if supervisor
                                                    (collect supervisor new-collected (inc depth))
                                                    new-collected)
                                   
                                   ;; Collect a few subordinates
                                   with-subordinates (reduce (fn [acc sub]
                                                              (if (< (count acc) max-positions)
                                                                (collect sub acc (inc depth))
                                                                acc))
                                                            with-supervisor
                                                            (take 5 subordinates))]
                               with-subordinates)))
        
        ;; Collect connected positions starting from pseudo-roots
        selected-ids (reduce (fn [acc pos]
                              (if (< (count acc) max-positions)
                                (into acc (collect-connected pos #{} 0))
                                acc))
                            #{}
                            starting-positions)
        
        ;; Return the actual position records, limited to max-positions
        selected-positions (take max-positions (keep #(get position-map %) selected-ids))]
    
    selected-positions))

(defn ensure-supervisors-included
  "Given a set of positions, ensure their supervisors are also included from the full dataset"
  [selected-positions full-dataset]
  (let [full-position-map (into {} (map (juxt :position identity) full-dataset))
        selected-ids (set (map :position selected-positions))
        
        ;; Find supervisors referenced by selected positions
        referenced-supervisor-ids (set (filter some? (map :reports-to-position selected-positions)))
        
        ;; Get supervisor positions from full dataset
        missing-supervisors (keep #(when (and (contains? referenced-supervisor-ids %)
                                             (not (contains? selected-ids %)))
                                    (get full-position-map %))
                                 referenced-supervisor-ids)]
    
    ;; Combine selected positions with their supervisors
    (concat selected-positions missing-supervisors)))

(defn save-simplified-org-chart
  "Save a simplified org chart optimized for large datasets"
  [positions filename & {:keys [format max-positions show-vacant include-dotted-lines]
                         :or {format "png" max-positions 75 show-vacant false include-dotted-lines false}}]
  (let [;; First, take a basic selection
        basic-selection (take max-positions positions)
        
        ;; Then ensure supervisors are included from the full dataset
        positions-with-supervisors (ensure-supervisors-included basic-selection positions)
        
        ;; Filter out vacant positions if requested
        filtered-positions (if show-vacant 
                            positions-with-supervisors
                            (remove #(= (str/upper-case (or (:current-employee %) "")) "VACANT") positions-with-supervisors))
        
        ;; For SVG, limit to prevent memory issues - use safe limits based on testing
        final-positions (case format
                          "svg" (take 100 filtered-positions)   ; Safe limit: ~165 actual positions
                          :svg (take 100 filtered-positions)    ; Safe limit: ~165 actual positions  
                          filtered-positions)                   ; PNG can handle more
        
        ;; Use large chart format for 200+ positions, regular for smaller charts
        dot-content (if (>= (count final-positions) 200)
                      (generate-large-org-chart-dot final-positions
                                                    :show-codes false
                                                    :include-dotted-lines include-dotted-lines
                                                    :rankdir "TB")
                      (generate-org-chart-dot final-positions
                                              :show-codes false
                                              :include-dotted-lines include-dotted-lines
                                              :rankdir "TB"))]
    (case (keyword format)
      :png (dot->image dot-content filename)
      :svg (let [dot-filename (str/replace filename #"\.svg$" ".dot")]
             (spit dot-filename dot-content)
             ;; Use simple, reliable SVG conversion with memory limits
             (let [result (shell/sh "dot" "-Tsvg" "-Gmaxiter=100" "-Gnslimit=50" dot-filename "-o" filename)]
               (if (= 0 (:exit result))
                 (println (str "✓ SVG generated with " (count final-positions) " positions (limited for memory)"))
                 (println "Warning: SVG conversion failed:" (:err result)))))
      :dot (spit filename dot-content))
    (println (str "Simplified org chart (" (count final-positions) " positions) saved to: " filename))))


(defn save-org-chart
  "Save an org chart as an image file"
  [positions filename & {:keys [format show-codes include-dotted-lines rankdir strict-filter errors title]
                         :or {format "png" show-codes false
                              include-dotted-lines true rankdir "TB" strict-filter false errors [] title nil}}]
  (let [dot (generate-org-chart-dot positions
                                    :show-codes show-codes
                                    :include-dotted-lines include-dotted-lines
                                    :rankdir rankdir
                                    :strict-filter strict-filter
                                    :errors errors
                                    :title title)]
    (case format
      "svg" (copy (dot->svg dot) (file filename))
      "png" (copy (dot->image dot "png") (file filename))
      "jpg" (copy (dot->image dot "jpg") (file filename))
      "dot" (spit filename dot)  ; <-- This line was added
      (throw (IllegalArgumentException. (str "Unsupported format: " format))))))

(defn save-hierarchical-charts
  "Save multiple smaller org charts by department/unit"
  [positions output-prefix & {:keys [format max-per-chart group-by]
                              :or {format "png" max-per-chart 30 group-by :unitcode}}]
  (let [subsets (create-hierarchical-subsets positions 
                                             :max-per-chart max-per-chart
                                             :group-by-field group-by)]
    (doseq [{:keys [group positions total-count]} subsets]
      (when (> (count positions) 1) ; Only create charts with multiple positions
        (let [filename (str output-prefix "_" (name group) "." (name format))
              safe-filename (str/replace filename #"[^a-zA-Z0-9._-]" "_")]
          (save-org-chart positions safe-filename :format format)
          (println (str "Chart for " group ": " (count positions) "/" total-count " positions -> " safe-filename)))))
    (println (str "Generated " (count subsets) " departmental org charts"))))

(defn save-org-chart-for-codes
  "Save an org chart for only the positions matching the given codes (by :position or :reports-to-position)."
  [positions codes output-file & {:keys [format show-codes include-dotted-lines rankdir strict-filter errors title]
                                  :or {format "png" show-codes false include-dotted-lines true rankdir "TB" strict-filter false errors [] title nil}}]
  (let [filtered (if strict-filter 
                   (filter-positions-by-exact-codes positions codes)
                   (filter-positions-by-codes positions codes))]
    (save-org-chart filtered output-file
                    :format format
                    :show-codes show-codes
                    :include-dotted-lines include-dotted-lines
                    :rankdir rankdir
                    :strict-filter strict-filter
                    :errors errors
                    :title title)))

(defn save-executive-summary-chart
  "Save a high-level executive summary chart with only top positions"
  [positions filename & {:keys [format] :or {format "png"}}]
  (let [executive-positions (filter-positions-by-level positions 
                                                       :max-levels 2
                                                       :include-titles #{"Director" "Chief" "Deputy"}
                                                       :exclude-vacant true)]
    (save-org-chart executive-positions filename 
                    :format format 
                    :show-codes false
                    :include-dotted-lines false)
    (println (str "Executive summary chart (" (count executive-positions) " positions) saved to: " filename))))



(defn create-department-subgraph
  "Create a subgraph for a specific department/agency"
  [positions agency-code & {:keys [show-codes] 
                            :or {show-codes false}}]
  (let [dept-positions (filter #(= (:agencycode %) agency-code) positions)]
    (positions->org-chart dept-positions 
                          :show-codes show-codes
                          :include-dotted-lines false)))

(defn generate-multi-department-chart
  "Generate an org chart with departments as subgraphs"
  [positions & {:keys [show-codes rankdir]
                :or {show-codes false rankdir "TB"}}]
  (let [all-nodes (map #(create-position-node % :show-codes show-codes) positions)
        all-edges (create-reporting-edges positions)]
    
    (graph->dot all-nodes all-edges
                {:directed? true
                 :graph {:rankdir rankdir
                         :dpi 150
                         :compound true}
                 :node {:fontname "Arial"
                        :fontsize 10}
                 :node->id :id
                 :node->descriptor (fn [node] (dissoc node :id))})))

;; Convenience functions for common use cases
(defn quick-org-chart
  "Quickly generate and save an org chart from Excel data"
  [excel-file output-file & {:keys [format minimal] :or {format "png" minimal true}}]
  (let [xlsx-data (xlsx/extract-data excel-file :minimal minimal)
        positions (pos/extract-positions xlsx-data)]
    (save-org-chart positions output-file :format format)
    (println (str "Org chart saved to: " output-file))))

(defn org-chart-stats
  "Get statistics about the org chart"
  [positions]
  (let [total-positions (count positions)
        vacant-positions (count (filter #(or (nil? (:current-employee %))
                                              (= (:current-employee %) "Vacant")) positions))
        filled-positions (- total-positions vacant-positions)
        departments (distinct (map :agencycode positions))
        titles (frequencies (map :title positions))]
    {:total-positions total-positions
     :filled-positions filled-positions
     :vacant-positions vacant-positions
     :fill-rate (double (/ filled-positions total-positions))
     :departments (count departments)
     :department-list departments
     :title-distribution titles}))




(defn save-complete-org-chart
  "Generate and save a complete org chart image for all positions using tangle.core."
  [positions output-file & {:keys [format show-codes include-dotted-lines rankdir]
                            :or {format "png" show-codes false include-dotted-lines true rankdir "TB"}}]
  (let [{:keys [nodes edges]} (positions->org-chart positions
                                                    :show-codes show-codes
                                                    :include-dotted-lines include-dotted-lines)
        dot (graph->dot nodes edges
                        {:directed? true
                         :graph {:rankdir rankdir :dpi 150 :splines "ortho" :concentrate true}
                         :node {:fontname "Arial" :fontsize 11 :margin 0.2 :width 2.5 :height 1.0}
                         :edge {:fontname "Arial" :fontsize 9 :arrowsize 0.8}
                         :node->id :id
                         :node->descriptor (fn [node] (dissoc node :id))})]
    (case format
      "png" (copy (dot->image dot "png") (file output-file))
      "svg" (copy (dot->svg dot) (file output-file))
      "dot" (spit output-file dot)
      (throw (IllegalArgumentException. (str "Unsupported format: " format))))
    (println (str "Complete org chart saved to: " output-file))))




(comment
  ;; Example usage:
  
  ;; Load data and create a basic org chart
  (def positions (pos/extract-positions (xlsx/extract-data "resources/Org Chart Data Analysis.xlsx" :minimal true)))
  
  ;; Save a simple org chart
  (save-org-chart positions "org-chart.png")
  
  ;; Save with more details
  (save-org-chart positions "detailed-org-chart.svg" 
                  :format "svg" 
                  :show-codes true)
  
  ;; Create a horizontal layout
  (save-org-chart positions "horizontal-org-chart.png" 
                  :rankdir "LR")
  
  ;; Quick generation from Excel file
  (quick-org-chart "resources/Org Chart Data Analysis.xlsx" "quick-chart.png")
  
  ;; Get org chart statistics
  (org-chart-stats positions)
  
  ;; Generate DOT for manual editing
  (def dot-output (generate-org-chart-dot positions))
  (spit "org-chart.dot" dot-output)
  
  :rcf)