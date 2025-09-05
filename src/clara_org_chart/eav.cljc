(ns clara-xlsx-diff.eav
  "EAV (Entity-Attribute-Value) transformation utilities for XLSX data"
  (:require [clara-eav.eav :as eav]))

(defn cell->eav
  "Transform a single cell into EAV triples.
  
  Each cell becomes an entity with attributes for its properties."
  [sheet-name cell version]
  ;; (println "testing whats inside the cell" cell)
  (let [entity-id (str version ":" sheet-name ":" (:cell-ref cell))]
    [;; Core cell data as EAV records
     (eav/->EAV entity-id :cell/sheet sheet-name)
     (eav/->EAV entity-id :cell/ref (:cell-ref cell))
     (eav/->EAV entity-id :cell/row (:row cell))
     (eav/->EAV entity-id :cell/col (:col cell))
     (eav/->EAV entity-id :cell/value (:value cell))
     (eav/->EAV entity-id :cell/type (:type cell))
     (eav/->EAV entity-id :cell/version version)]))

(defn sheet->eav
  "Transform a single sheet into EAV triples"
  [sheet version]
  (let [sheet-name (:sheet-name sheet)]
    (mapcat #(cell->eav sheet-name % version) (:cells sheet))))


(defn extract-facts
  "Extracts the :fact key from each map in a sequence."
  [sequence-of-matches]
  (map :fact sequence-of-matches))


(defn xlsx->eav
  "Transform XLSX data structure into EAV triples.
  
  Each cell becomes an entity with attributes describing its properties,
  making it easy to query and compare using Clara Rules."
  [xlsx-data & {:keys [version] :or {version :v1}}]
  (mapcat #(sheet->eav % version) (:sheets xlsx-data)))

(defn eav->cell-map
  "Convert EAV triples back to a cell-centric map for easier processing"
  [eav-triples]
  (reduce (fn [acc {:keys [e a v]}]
            (assoc-in acc [e a] v))
          {}
          eav-triples))

(defn get-cells-by-sheet
  "Get all cell entities grouped by sheet name from EAV triples"
  [eav-triples]
  (let [cell-map (eav->cell-map eav-triples)]
    (group-by #(get-in cell-map [% :cell/sheet]) (keys cell-map))))

(comment
  ;; REPL experiments
  (require '[clara-xlsx-diff.xlsx :as xlsx])
  
  ;; Transform sample data
  (def sample-data (xlsx/extract-data "test/sample_data.xlsx"))
  (def eav-triples (xlsx->eav sample-data :version :v1))

  (def eav-v2 (xlsx->eav (xlsx/extract-data "test/sample_data.xlsx") :version :v2))
  ;; Inspect EAV structure
  (take 10 eav-triples)
  
  ;; Convert back to map for inspection
  (def cell-map (eav->cell-map eav-triples))
  (keys cell-map)
  
  ;; Group by sheet
  (get-cells-by-sheet eav-triples)
  )
