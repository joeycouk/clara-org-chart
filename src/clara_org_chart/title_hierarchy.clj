(ns clara-org-chart.title-hierarchy
  (:require [clojure.edn :as edn]
            [clojure.string :as str]
            [clojure.set :as set]))

(defn load-title-hierarchy
  "Load the title reporting hierarchy from EDN file"
  []
  (edn/read-string (slurp "title-reporting-hierarchy.edn")))

(defn get-valid-supervisors
  "Get all valid supervisor titles for a given title"
  [title]
  (let [hierarchy (load-title-hierarchy)]
    (get hierarchy title #{})))

(defn get-potential-subordinates
  "Get all titles that can report to the given supervisor title"
  [supervisor-title]
  (let [hierarchy (load-title-hierarchy)]
    (set (for [[subordinate supervisors] hierarchy
               :when (contains? supervisors supervisor-title)]
           subordinate))))

(defn fuzzy-title-match
  "Find titles that approximately match the given string (case-insensitive)"
  [title-fragment]
  (let [hierarchy (load-title-hierarchy)
        titles (keys hierarchy)
        fragment-lower (str/lower-case title-fragment)]
    (filter #(str/includes? (str/lower-case %) fragment-lower) titles)))

(defn normalize-title
  "Normalize a title by trimming whitespace and standardizing case variations"
  [title]
  (when title
    (-> title
        str/trim
        (str/replace #"\s+" " ")
        str/lower-case)))

(defn normalize-title-for-lookup
  "Normalize a title for hierarchy lookup - preserves original case but normalizes spacing"
  [title]
  (when title
    (-> title
        str/trim
        (str/replace #"\s+" " "))))

(defn can-report-to?
  "Check if subordinate-title can report to supervisor-title (case and spacing insensitive)"
  [subordinate-title supervisor-title]
  (let [hierarchy (load-title-hierarchy)
        normalized-sub (normalize-title subordinate-title)
        normalized-sup (normalize-title supervisor-title)]
    ;; First try exact match with original titles
    (if-let [valid-supervisors (get hierarchy subordinate-title #{})]
      (or (contains? valid-supervisors supervisor-title)
          ;; If exact supervisor match fails, try normalized supervisor comparison
          (some #(= normalized-sup (normalize-title %)) valid-supervisors))
      ;; If no exact subordinate match, try with normalized titles
      (let [hierarchy-normalized (into {}
                                       (for [[k v] hierarchy]
                                         [(normalize-title k)
                                          (set (map normalize-title v))]))
            valid-supervisors-norm (get hierarchy-normalized normalized-sub #{})]
        (contains? valid-supervisors-norm normalized-sup)))))

(defn find-exact-or-similar-title
  "Find exact match or similar titles from the hierarchy"
  [title]
  (let [hierarchy (load-title-hierarchy)
        normalized-title (normalize-title-for-lookup title)
        exact-match (get hierarchy normalized-title)]
    (if exact-match
      {:exact-match normalized-title :supervisors exact-match}
      {:similar-matches (fuzzy-title-match title)})))

(defn validate-reporting-relationship
  "Validate if a reporting relationship makes organizational sense"
  [subordinate-title supervisor-title]
  (let [valid? (can-report-to? subordinate-title supervisor-title)
        valid-supervisors (get-valid-supervisors subordinate-title)
        supervisor-subordinates (get-potential-subordinates supervisor-title)]
    {:valid? valid?
     :subordinate subordinate-title
     :supervisor supervisor-title
     :valid-supervisors-for-subordinate valid-supervisors
     :valid-subordinates-for-supervisor supervisor-subordinates}))

(defn get-organizational-level
  "Determine the organizational level of a title based on common patterns"
  [title]
  (let [title-lower (str/lower-case (or title ""))]
    (cond
      (re-find #"director" title-lower) :executive
      (re-find #"deputy director" title-lower) :deputy-director
      (re-find #"assistant deputy director" title-lower) :assistant-deputy-director
      (re-find #"division chief" title-lower) :division-chief
      (re-find #"assistant chief" title-lower) :assistant-chief
      (re-find #"deputy chief" title-lower) :deputy-chief
      (re-find #"battalion chief" title-lower) :battalion-chief
      (re-find #"unit chief" title-lower) :unit-chief
      (re-find #"captain" title-lower) :captain
      (re-find #"fire fighter ii" title-lower) :fire-fighter-ii
      (re-find #"fire fighter i" title-lower) :fire-fighter-i
      (re-find #"manager" title-lower) :manager
      (re-find #"supervisor" title-lower) :supervisor
      (re-find #"analyst" title-lower) :analyst
      (re-find #"specialist" title-lower) :specialist
      (re-find #"technician" title-lower) :technician
      (re-find #"assistant" title-lower) :assistant
      :else :unknown)))

(defn suggest-supervisors-by-level
  "Suggest possible supervisors based on organizational level patterns"
  [title]
  (let [level (get-organizational-level title)]
    (case level
      :fire-fighter-i #{:captain :battalion-chief :division-chief}
      :fire-fighter-ii #{:captain :battalion-chief :division-chief}
      :captain #{:battalion-chief :division-chief :assistant-chief}
      :battalion-chief #{:division-chief :assistant-chief :deputy-chief}
      :unit-chief #{:division-chief :assistant-chief}
      :deputy-chief #{:division-chief :assistant-deputy-director}
      :assistant-chief #{:division-chief :assistant-deputy-director}
      :division-chief #{:assistant-deputy-director :deputy-director}
      :analyst #{:manager :supervisor :division-chief}
      :specialist #{:manager :supervisor :division-chief}
      :technician #{:supervisor :manager :division-chief}
      :manager #{:division-chief :assistant-chief}
      :supervisor #{:manager :division-chief}
      #{:division-chief :assistant-chief})))

(defn analyze-title-coverage
  "Analyze how many titles from position-mapping.edn are covered in the hierarchy"
  [position-mapping-file]
  (let [hierarchy (load-title-hierarchy)
        hierarchy-titles (set (keys hierarchy))
        position-titles (set (edn/read-string (slurp position-mapping-file)))
        covered (set/intersection hierarchy-titles position-titles)
        missing (set/difference position-titles hierarchy-titles)]
    {:total-position-titles (count position-titles)
     :covered-titles (count covered)
     :missing-titles (count missing)
     :coverage-percentage (double (* 100 (/ (count covered) (count position-titles))))
     :missing-title-list (sort missing)}))

(defn get-chain-of-command
  "Get the potential chain of command upward from a given title"
  [title & {:keys [max-levels] :or {max-levels 5}}]
  (let [hierarchy (load-title-hierarchy)]
    (loop [current-title title
           chain [title]
           level 0]
      (if (or (>= level max-levels) (empty? (get hierarchy current-title)))
        chain
        (let [supervisors (get hierarchy current-title #{})]
          (if (empty? supervisors)
            chain
            ;; Take the first supervisor for simplicity, in practice you might want to handle multiple paths
            (let [next-supervisor (first supervisors)]
              (recur next-supervisor 
                     (conj chain next-supervisor)
                     (inc level)))))))))

;; Testing and utility functions
(defn test-hierarchy-functions
  "Test the hierarchy functions with some example data"
  []
  (println "=== Title Hierarchy Testing ===")
  
  ;; Test basic reporting relationships
  (println "\n1. Basic Reporting Tests:")
  (println "Fire Fighter II can report to Battalion Chief:" 
           (can-report-to? "Fire Fighter II" "Relief Battalion Chief"))
  (println "Division Chief can report to Fire Fighter I:" 
           (can-report-to? "DIVISION CHIEF" "Fire Fighter II"))
  
  ;; Test getting supervisors
  (println "\n2. Valid Supervisors:")
  (println "Fire Fighter II supervisors:" (get-valid-supervisors "Fire Fighter II"))
  (println "Battalion Chief subordinates:" (get-potential-subordinates "Relief Battalion Chief"))
  
  ;; Test fuzzy matching
  (println "\n3. Fuzzy Title Matching:")
  (println "Titles containing 'fire':" (take 5 (fuzzy-title-match "fire")))
  (println "Titles containing 'chief':" (take 5 (fuzzy-title-match "chief")))
  
  ;; Test organizational levels
  (println "\n4. Organizational Levels:")
  (println "Division Chief level:" (get-organizational-level "DIVISION CHIEF"))
  (println "Fire Fighter II level:" (get-organizational-level "Fire Fighter II"))
  
  ;; Test chain of command
  (println "\n5. Chain of Command:")
  (println "Fire Fighter II chain:" (get-chain-of-command "Fire Fighter II"))
  
  ;; Test validation
  (println "\n6. Relationship Validation:")
  (println "Valid relationship test:" 
           (validate-reporting-relationship "Fire Fighter II" "Relief Battalion Chief")))

(comment
  ;; Example usage:
  
  ;; Load and test the hierarchy
  (test-hierarchy-functions)
  
  ;; Check specific relationships
  (clara-org-chart.title-hierarchy/can-report-to? "Staff Services Analyst" " Forester III") ; => true
  (can-report-to? "DIVISION CHIEF" "Fire Fighter II") ; => false
  
  ;; Get all valid supervisors for a title
  (get-valid-supervisors "Staff Services Analyst")
  
  ;; Find who can report to a specific supervisor
  (get-potential-subordinates "Division Chief Administration")
  
  ;; Find similar titles
  (fuzzy-title-match "captain")
  (fuzzy-title-match "analyst")
  
  ;; Validate a relationship
  (validate-reporting-relationship "Office Assistant" "STAFF SERVICES MANAGER I PERF MANAGEMENT")
  
  ;; Get organizational level
  (get-organizational-level "Battalion Chief Law Enforcement")
  
  ;; Get chain of command
  (get-chain-of-command "Fire Fighter II")
  
  ;; Analyze coverage (requires position-mapping.edn file)
  (analyze-title-coverage "position-mapping.edn")
  
  :rcf)