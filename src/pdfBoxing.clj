(ns pdfBoxing
  (:import [org.apache.pdfbox.pdmodel PDDocument]
           [org.apache.pdfbox.text PDFTextStripper])
  (:require [clojure.java.io :as io]))

;; Regex for position codes of form ###-###-####-### (3-3-4-3 digits with hyphens)
(def ^:private position-code-regex #"\b\d{3}-\d{3}-\d{4}-\d{3}\b")


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


(defn- extract-pages
  "Return a vector of page texts from a PDF path using pdfboxing.
  IMPORTANT: Preserve blank pages so that page indices match the original PDF.
  pdfboxing.text/extract returns one big string separated by form-feed (\f).
  We split with limit -1 to retain trailing empty segment if present."
  [pdf-path]
  (let [raw (pdf/extract pdf-path)]
    (cond
      (nil? raw) []
      (string? raw) (->> (str/split raw #"\f" -1)
                         ;; do not drop blanks; just trim CR characters
                         (map #(str/replace % #"\r" ""))
                         vec)
      :else (vec raw))))

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
     :snippet (subs text 0 (min 300 (count text)))}))

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

;; Example usage:
(count ((comp vec extract-pages-pdfbox) "resources/Southern Region Org Charts 01.01.25.pdf"))
(positions-on-page "resources/Southern Region Org Charts 01.01.25.pdf" 3)
;; (find-positions-in-pdf "Southern Region Org Charts 01.01.25.pdf")
;; (find-positions-in-pdf "Southern Region Org Charts 01.01.25.pdf" ["542-434-1083-901" "541-314-1402-601"]) 