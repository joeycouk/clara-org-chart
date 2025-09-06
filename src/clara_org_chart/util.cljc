(ns clara-org-chart.util)


;; Utility function that calculates all neighboring cells 
(defn calculateAllRows
  "Returns a set of all neighboring cell coordinates for a given cell.
  The input is a row and column index, and it returns a set of [row col] pairs
  representing all immediate neighbors (8-directional adjacency). Remove any calculations with negative cells or rows"

  [row col]
  (filter #(and (>= (first %) 0) (>= (second %) 0))
          (set (for [row-offset [-1 0 1]
                     col-offset [-1 0 1]
                     :when (not (and (zero? row-offset) (zero? col-offset)))]
                 [(+ row row-offset) (+ col col-offset)]))))

(comment
;;   some tests of the calculateAllRows function
    (calculateAllRows 0 4)
  :rcf)


;; Utility functions for cell neighbor detection
(defn neighbors?
  "Returns true if two cells are immediate neighbors (8-directional adjacency).
  Neighbors are cells that differ by at most 1 in both row and column,
  but are not the same cell."
  [row1 col1 row2 col2]
  (let [row-diff #?(:clj  (Math/abs (- row1 row2))
                    :cljs (js/Math.abs (- row1 row2)))
        col-diff #?(:clj  (Math/abs (- col1 col2))
                    :cljs (js/Math.abs (- col1 col2)))]
    (and (not (and (= row1 row2) (= col1 col2)))  ; not the same cell
         (<= row-diff 1)                          ; within 1 row
         (<= col-diff 1))))                       ; within 1 column


(defn extract-Eid-from-entities
  "Extracts the EID from a collection of entities"
  ([entities]
   ;;  (println "entities" entities)
   (set (map #(get % :eav/eid) entities)))
  ([entities attribute value] ;; where attribute is a key on the map 
   ;;  filter the entities by the attribute and value
   ;;  (println "entities" entities)
   ;;  (println "attribute" attribute)
   ;;  (println "value" value)
   (set (map #(get % :eav/eid) (filter #(= (get % attribute) value) entities)))))

(comment
  (neighbors? 1 4 1 5)
    (neighbors? 0 2 0 4)
  :rcf)

(defn neighbor-cell?
  "Returns true if two cells are neighbors in the same column (vertically adjacent).
  This means they have the same column but differ by exactly 1 in row."
  [row1 col1 row2 col2]
  (and (= col1 col2)                    ; same column
       (let [row-diff #?(:clj  (Math/abs (- row1 row2))
                         :cljs (js/Math.abs (- row1 row2)))]
         (= row-diff 1))))               ; exactly 1 row apart


(defn neighbor-row?
  "Returns true if two cells are neighbors in the same row (horizontally adjacent).
  This means they have the same row but differ by exactly 1 in column."
  [row1 col1 row2 col2]
  (and (= row1 row2)                    ; same row
       (let [col-diff #?(:clj  (Math/abs (- col1 col2))
                         :cljs (js/Math.abs (- col1 col2)))]
         (= col-diff 1))))               ; exactly 1 column apart


(defn diagonal-neighbor?
  "Returns true if two cells are diagonal neighbors.
  This means they differ by exactly 1 in both row and column."
  [row1 col1 row2 col2]
  (let [row-diff #?(:clj  (Math/abs (- row1 row2))
                    :cljs (js/Math.abs (- row1 row2)))
        col-diff #?(:clj  (Math/abs (- col1 col2))
                    :cljs (js/Math.abs (- col1 col2)))]
    (and (= row-diff 1)                 ; exactly 1 row apart
         (= col-diff 1))))              ; exactly 1 column apart


;; Helper functions using the granular neighbor detection functions
(defn neighbors-alt?
  "Alternative implementation of neighbors? using the more granular functions.
  Returns true if two cells are immediate neighbors (8-directional adjacency)."
  [row1 col1 row2 col2]
  (or (neighbor-cell? row1 col1 row2 col2)     ; vertical neighbors
      (neighbor-row? row1 col1 row2 col2)      ; horizontal neighbors  
      (diagonal-neighbor? row1 col1 row2 col2))) ; diagonal neighbors


(defn neighbor-type
  "Returns the type of neighbor relationship between two cells.
  Returns one of: :same, :vertical, :horizontal, :diagonal, :not-neighbors"
  [row1 col1 row2 col2]
  (cond
    (and (= row1 row2) (= col1 col2))           :same
    (neighbor-cell? row1 col1 row2 col2)        :vertical
    (neighbor-row? row1 col1 row2 col2)         :horizontal
    (diagonal-neighbor? row1 col1 row2 col2)    :diagonal
    :else                                       :not-neighbors))


(defn adjacent-cell-coords
  "Returns all possible adjacent cell coordinates for a given cell.
  Returns a sequence of [row col] vectors for all 8 neighboring positions."
  [row col]
  (for [row-offset [-1 0 1]
        col-offset [-1 0 1]
        :when (not (and (zero? row-offset) (zero? col-offset)))]
    [(+ row row-offset) (+ col col-offset)]))

