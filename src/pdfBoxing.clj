(ns pdfBoxing
  (:import [org.apache.pdfbox.pdmodel PDDocument]
           [org.apache.pdfbox.text PDFTextStripper]
           [org.apache.pdfbox.text TextPosition])
  (:require [clojure.java.io :as io]
            [clojure.string :as str]))

;; Regex for position codes - matches regular, vacant, and X-pattern position formats
;; Regular: 542-061-1031-003 or 542-061-1031-003-XXX (requires 4 segments minimum)
;; Vacant with slash: 541-064-5157/5393-709 (indicates position could be multiple codes but is vacant)
;; Vacant with X: 541-022-4179-7XX (X represents variable digits in vacant positions)
;; The part after / appears to be a partial code: ####-### format
;; Note: Position codes MUST have 4 segments (###-###-####-###) - 3 segments like 541-031-5157 are not valid
(def ^:private position-code-regex #"\d{3}-\d{3}-\d{4}-[0-9X]+(?:-[0-9X]+)*(?:/\d{4}-\d{3}(?:-[0-9X]+)*)*")

(defn- clean-position-code
  "Extract just the position code part from a match that may include surrounding text.
  Handles regular cases like '542-061-1031-003Staff' -> '542-061-1031-003'
  vacant position cases like '541-064-5157/5393-709Text' -> '541-064-5157/5393-709'
  and X-pattern cases like '541-022-4179-7XXStaff' -> '541-022-4179-7XX'
  Note: Only matches valid 4-segment position codes (###-###-####-###)"
  [match-text]
  (when match-text
    ;; Try to match vacant position format first (with slash) - requires 4 segments before slash
    (or (re-find #"\d{3}-\d{3}-\d{4}-[0-9X]+(?:-[0-9X]+)*/\d{4}-\d{3}(?:-[0-9X]+)*" match-text)
        ;; Fall back to regular position format (including X patterns) - requires 4 segments
        (re-find #"\d{3}-\d{3}-\d{4}-[0-9X]+(?:-[0-9X]+)*" match-text))))


(defn extract-pages-pdfbox
  "Return a vector of page texts using PDFBox directly."
  [pdf-path]
  (with-open [doc (PDDocument/load (io/file pdf-path))]
    (let [stripper (PDFTextStripper.)]
      (mapv
       (fn [page]
         (.setStartPage stripper page)
         (.setEndPage stripper page)
         (.getText stripper doc))
       (range 1 (inc (.getNumberOfPages doc)))))))

(defn positions-with-coordinates-on-page
  "Extract position codes with their x,y coordinates from a specific page.
  Returns a vector of maps with :text (position code), :x, :y, :width, :height.
  Returns ALL distinct instances - duplicates with identical coordinates are removed."
  [pdf-path page-number & {:keys [debug?] :or {debug? false}}]
  (when debug? (println "DEBUG: Starting positions-with-coordinates-on-page for page" page-number))
  (with-open [doc (PDDocument/load (io/file pdf-path))]
    (when debug? (println "DEBUG: PDF loaded successfully, total pages:" (.getNumberOfPages doc)))
    (let [position-codes (atom [])
          text-positions (atom [])
          full-text (atom "")
          stripper (proxy [PDFTextStripper] []
                     (writeString [text text-pos-list]
                       ;; Store all character positions for later processing
                       (doseq [^TextPosition tp text-pos-list]
                         (let [char (.getUnicode tp)
                               char-index (count @full-text)]
                           (swap! text-positions conj {:char char
                                                       :x (.getX tp)
                                                       :y (.getY tp)
                                                       :width (.getWidth tp)
                                                       :height (.getHeight tp)
                                                       :index char-index})
                           (swap! full-text str char)))
                       
                       ;; Call parent method to continue normal processing
                       (proxy-super writeString text text-pos-list)))]
      
      (.setStartPage stripper page-number)
      (.setEndPage stripper page-number)
      (when debug? (println "DEBUG: About to extract text from page" page-number))
      (.getText stripper doc)
      (when debug? (println "DEBUG: Text extraction completed"))
      
      ;; After all text is processed, find position codes and their coordinates
      (let [final-text @full-text
            all-positions @text-positions]
        
        (when debug?
          (println "Final text length:" (count final-text))
          (println "Total character positions:" (count all-positions)))
        
        ;; Find all position code matches in the final text
        (let [matches (re-seq position-code-regex final-text)]
          (when debug?
            (println "Found position code matches:" matches)
            (println "First 500 chars of extracted text:")
            (println (subs final-text 0 (min 500 (count final-text)))))
          
          ;; For each match, find its location and extract coordinates
          (doseq [raw-match matches]
            (let [clean-match (clean-position-code raw-match)
                  match-pattern (re-pattern (java.util.regex.Pattern/quote raw-match))
                  matcher (re-matcher match-pattern final-text)]
              
              ;; Find all occurrences of this position code
              (loop [found-positions []]
                (if (.find matcher)
                  (let [start-idx (.start matcher)
                        end-idx (.end matcher)
                        relevant-positions (filter #(and (>= (:index %) start-idx)
                                                         (< (:index %) end-idx))
                                                  all-positions)]
                    
                    (when debug?
                      (println "Match" raw-match "-> clean:" clean-match "at indices" start-idx "to" end-idx
                              "with" (count relevant-positions) "character positions"))
                    
                    (when (= (count relevant-positions) (count raw-match))
                      ;; Calculate bounding box - use clean match for the result
                      (let [xs (map :x relevant-positions)
                            ys (map :y relevant-positions)
                            widths (map :width relevant-positions)
                            heights (map :height relevant-positions)
                            min-x (apply min xs)
                            max-x (apply max (map + xs widths))
                            min-y (apply min ys)
                            max-y (apply max (map + ys heights))
                            coord-data {:text (or clean-match raw-match)  ; Prefer clean match
                                       :x min-x
                                       :y min-y
                                       :width (- max-x min-x)
                                       :height (- max-y min-y)}]
                        
                        (swap! position-codes conj coord-data)
                        (when debug?
                          (println "Added coordinates for" (or clean-match raw-match) ":" coord-data))))
                    
                    (recur (conj found-positions [start-idx end-idx])))
                  
                  ;; No more matches found
                  nil)))))
      
      (when debug?
        (let [simple-matches (re-seq position-code-regex @full-text)
              coord-matches @position-codes]
          (println "\n=== COORDINATE EXTRACTION DEBUG ===")
          (println "Page" page-number "final text length:" (count @full-text))
          (println "Simple regex found" (count simple-matches) "matches:" simple-matches)
          (println "Coordinate extraction found" (count coord-matches) "matches:" (map :text coord-matches))))
      
      ;; Return all position code instances, but deduplicate identical coordinates
      ;; If text, x, y, width, and height are all identical, it's the same instance
      (let [all-results @position-codes
            unique-by-coords (->> all-results
                                 (group-by (fn [entry] 
                                            ;; Create a unique key from all coordinate values
                                            [(:text entry) (:x entry) (:y entry) 
                                             (:width entry) (:height entry)]))
                                 (map (fn [[_coords entries]] (first entries))) ; Take first of each group
                                 vec)]
        (when debug?
          (println "Before deduplication:" (count all-results) "entries")
          (println "After deduplication:" (count unique-by-coords) "entries"))
        unique-by-coords)))))


;; Debug helpers ------------------------------------------------------------

(def unicode-dash-chars #{\u2010 \u2011 \u2012 \u2013 \u2014 \u2212})

(defn dash-frequency
  "Return a map of any non-ASCII dash characters found in s to their counts."
  [s]
  (->> s (filter unicode-dash-chars) frequencies))

(defn debug-page
  "Return basic debug info for a 1-based page number: length, first 300 chars, dash freq."
  [pdf-path page-number]
  (let [pages (extract-pages-pdfbox pdf-path)
        idx (dec page-number)
        total (count pages)
        text (nth pages idx "")]
    {:requested-page page-number
     :total-pages total
     :page-empty? (str/blank? text)
     :length (count text)
     :dash-frequency (dash-frequency text)
     :snippet text}))

;; forward declare since raw-candidates-on-page references positions-on-page
(declare positions-on-page)

(defn raw-candidates-on-page
  "Looser pattern scan to see what near-position-like strings exist on the page.
  Returns up to first 30 candidates using a permissive pattern where any single
  char separates the digit groups. Useful when strict regex yields no results."
  [pdf-path page-number]
  (let [pages (extract-pages-pdfbox pdf-path)
        idx (dec page-number)
        text (nth pages idx "")
        permissive #"\d{3}.\d{3}.\d{4}.\d{3}" ;; any single char between groups
        cands (->> (re-seq permissive text) (distinct) (take 30))]
    {:page page-number
     :strict-matches nil ;; filled after positions-on-page is defined
     :permissive-candidates cands
     :dash-frequency (dash-frequency text)}))

(defn positions-on-page
  "Return all position codes (format ###-###-####-###) found on a 1-based page number.
  Options: :unique? (default true) -> remove duplicates. Returns [] if out of range."
  [pdf-path page-number & {:keys [unique?] :or {unique? true}}]
  (let [pages (extract-pages-pdfbox pdf-path)
        idx (dec page-number)]
    (if (and (>= idx 0) (< idx (count pages)))
      (let [text (nth pages idx)
            matches (re-seq position-code-regex (or text ""))]
        (if unique? (vec (distinct matches)) (vec matches)))
      [])))

(alter-var-root #'raw-candidates-on-page
                (fn [_]
                  (fn [pdf-path page-number]
                    (let [pages (extract-pages-pdfbox pdf-path)
                          idx (dec page-number)
                          text (nth pages idx "")
                          permissive #"\d{3}.\d{3}.\d{4}.\d{3}"
                          cands (->> (re-seq permissive text) distinct (take 30))]
                      {:page page-number
                       :strict-matches (positions-on-page pdf-path page-number :unique? true)
                       :permissive-candidates cands
                       :dash-frequency (dash-frequency text)}))))

(defn find-positions-in-pdf
  "Scan the whole PDF and return a map of position-code -> sorted vector of page numbers it appears on.
  If `position-codes` (seq) is provided, limit detection to those codes; otherwise discover all codes."
  ([pdf-path] (find-positions-in-pdf pdf-path nil))
  ([pdf-path position-codes]
   (let [pages (extract-pages-pdfbox pdf-path)
         restrict? (seq position-codes)
         code-set (when restrict? (set position-codes))]
     (->> pages
          (map-indexed (fn [i text]
                         (let [page (inc i)
                               found (if restrict?
                                       ;; Only test supplied codes
                                       (filter #(re-find (re-pattern (java.util.regex.Pattern/quote %)) (or text "")) code-set)
                                       ;; Discover all codes on page
                                       (re-seq position-code-regex (or text "")))]
                           [page (distinct found)])))
          (reduce (fn [m [page codes]]
                    (reduce (fn [m2 code]
                              (update m2 code (fnil (fn [v] (vec (distinct (concat v [page])))) [])))
                            m
                            codes))
                  {})
          (update-vals (fn [pgs] (->> pgs (sort) vec)))))))


(comment
  ;; Example usage:
  (count ((comp vec extract-pages-pdfbox) "resources/Southern Region Org Charts 01.01.25.pdf"))
  (tap> (positions-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3))
  (tap> (extract-pages-pdfbox "resources/Southern Region Org Charts 01.01.25.pdf"))
  ;; (find-positions-in-pdf "Southern Region Org Charts 01.01.25.pdf")
  ;; (find-positions-in-pdf "Southern Region Org Charts 01.01.25.pdf" ["542-434-1083-901" "541-314-1402-601"]) 
  
  ;; Test coordinate extraction - updated regex with alphanumeric extensions
  (tap> (re-seq position-code-regex "Test string 542-061-1039-904-005 and 541-028-4800-016 and 542-061-1039-904-XXX"))
  (tap> (debug-page "resources/Sac HQ Org Charts 01.01.25.pdf" 60))
  (tap> (positions-on-page "resources/Sac HQ Org Charts 01.01.25.pdf" 60))
  (tap> (positions-with-coordinates-on-page "resources/Sac HQ Org Charts 01.01.25.pdf" 60 :debug? true)) 
  (tap> (positions-with-coordinates-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3))
  
  ;; Compare with simple text extraction
  (let [simple (positions-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3)
        with-coords (positions-with-coordinates-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3)]
     (tap> with-coords)
    (println "Simple extraction found:" (count simple) "position codes")
    (println "Coordinate extraction found:" (count with-coords) "position codes")
    (println "All coordinate instances:" (map :text with-coords)))
  
  (tap> (debug-page "resources/Southern Region Org Charts 01.01.25.pdf" 2))
  :rcf)
