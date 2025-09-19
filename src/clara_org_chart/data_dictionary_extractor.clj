(ns clara-org-chart.data-dictionary-extractor
  (:require [clojure.string :as str]))

(defrecord ClassificationCode
  [class-code
   class-title])

(defn extract-classification-codes
  "Extract classification codes from the 'ClassCodes' sheet in the parsed xlsx data."
  [xlsx-data]
  (let [sheet (->> (:sheets xlsx-data)
                   (filter #(= (:sheet-name %) "ClassCodes"))
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
         (map (fn [[_ row-cells]]
                ;; Create a complete row mapping with all columns 0 to max-col
                (let [row-cell-map (into {} (map (fn [cell] [(:col cell) (:value cell)]) row-cells))
                      complete-row-data (mapv #(get row-cell-map % nil) (range 0 (inc max-col)))
                      row-map (zipmap headers complete-row-data)
                      class-code (get row-map "ClassCode")
                      class-code-str (if (integer? class-code)
                                       (str class-code)
                                       (str/replace (str class-code) #"\.0$" ""))]
                  (map->ClassificationCode
                   {:class-code class-code-str
                    :class-title (get row-map "ClassTitle")})))))))