(ns clara-org-chart.xlsx-jvm
  
  "XLSX file parsing and data extraction utilities"
   (:require [clojure.java.io :as io])
  (:import [org.apache.poi.ss.usermodel WorkbookFactory Sheet Row Cell CellType CellStyle Font FillPatternType BorderStyle DataFormatter]
        [org.apache.poi.ss.util CellReference]
        [org.apache.poi.ss.formula.function FunctionMetadataRegistry]
        [org.apache.poi.xssf.usermodel XSSFColor]
        [java.io FileInputStream]))
  

  (defn cell-value
    "Extract value from a POI Cell object, evaluating formulas. Returns nil if formula evaluation fails. Logs formula evaluation for debugging."
    [^Cell cell]
    (when cell
      (let [cell-type (.getCellType cell)
            cell-ref (try (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell))) (catch Exception _ "?"))]
        (cond
          (= cell-type CellType/STRING)
          (let [v (.getStringCellValue cell)]
            (println "Cell" cell-ref "type: STRING, value:" v)
            v)
          (= cell-type CellType/NUMERIC)
          (let [v (.getNumericCellValue cell)]
            (println "Cell" cell-ref "type: NUMERIC, value:" v)
            v)
          (= cell-type CellType/BOOLEAN)
          (let [v (.getBooleanCellValue cell)]
            (println "Cell" cell-ref "type: BOOLEAN, value:" v)
            v)
          (= cell-type CellType/FORMULA)
          (try
            (let [workbook (.getWorkbook (.getSheet cell))
                  evaluator (.createFormulaEvaluator (.getCreationHelper workbook))
                  evaluated (.evaluate evaluator cell)
                  cell-type-evaluated (.getCellType evaluated)
                  result (cond
                           (.equals cell-type-evaluated CellType/STRING) (.getStringValue evaluated)
                           (.equals cell-type-evaluated CellType/NUMERIC) (.getNumberValue evaluated)
                           (.equals cell-type-evaluated CellType/BOOLEAN) (.getBooleanValue evaluated)
                           (.equals cell-type-evaluated CellType/ERROR) (.getErrorValue evaluated)
                           (.equals cell-type-evaluated CellType/BLANK) nil
                           :else nil)]
              (println "Formula cell:" cell-ref "formula:" (.getCellFormula cell) "Evaluated value:" result)
              result)
            (catch Exception e
              (println "Formula evaluation error for cell" cell-ref "formula:" (.getCellFormula cell) ":" (.getMessage e))
              nil))
          (= cell-type CellType/BLANK)
          (do (println "Cell" cell-ref "type: BLANK") nil)
          (= cell-type CellType/_NONE)
          (do (println "Cell" cell-ref "type: _NONE") nil)
          :else
          (do (println "Cell" cell-ref "type: UNKNOWN" cell-type) nil)))))
  
  ;; (defn extract-font-info
  ;;   "Extract font information from a cell"
  ;;   [^Cell cell]
  ;;   (when cell
  ;;     (let [workbook (.getWorkbook (.getSheet cell))
  ;;           cell-style (.getCellStyle cell)
  ;;           font (.getFontAt workbook (.getFontIndexAsInt cell-style))]
  ;;       {:font-name (.getFontName font)
  ;;        :font-size (.getFontHeightInPoints font)
  ;;        :bold (.getBold font)
  ;;        :italic (.getItalic font)
  ;;        :underline (not= (.getUnderline font) 0)
  ;;        :strikeout (.getStrikeout font)
  ;;        :color (when-let [color (.getColor font)]
  ;;                 (str color))})))
  
  ;; (defn extract-cell-style-info
  ;;   "Extract comprehensive style information from a cell"
  ;;   [^Cell cell]
  ;;   (when cell
  ;;     (let [cell-style (.getCellStyle cell)]
  ;;       {:alignment (str (.getAlignment cell-style))
  ;;        :vertical-alignment (str (.getVerticalAlignment cell-style))
  ;;        :wrap-text (.getWrapText cell-style)
  ;;        :indention (.getIndention cell-style)
  ;;        :rotation (.getRotation cell-style)
  ;;        :fill-pattern (str (.getFillPattern cell-style))
  ;;        :fill-foreground-color (when-let [color (.getFillForegroundColorColor cell-style)]
  ;;                                 (str color))
  ;;        :fill-background-color (when-let [color (.getFillBackgroundColorColor cell-style)]
  ;;                                 (str color))
  ;;        :border-top (str (.getBorderTop cell-style))
  ;;        :border-right (str (.getBorderRight cell-style))
  ;;        :border-bottom (str (.getBorderBottom cell-style))
  ;;        :border-left (str (.getBorderLeft cell-style))
  ;;        :border-top-color (when-let [color (.getTopBorderColor cell-style)]
  ;;                            (str color))
  ;;        :border-right-color (when-let [color (.getRightBorderColor cell-style)]
  ;;                              (str color))
  ;;        :border-bottom-color (when-let [color (.getBottomBorderColor cell-style)]
  ;;                               (str color))
  ;;        :border-left-color (when-let [color (.getLeftBorderColor cell-style)]
  ;;                             (str color))
  ;;        :data-format (.getDataFormat cell-style)
  ;;        :data-format-string (.getDataFormatString cell-style)
  ;;        :hidden (.getHidden cell-style)
  ;;        :locked (.getLocked cell-style)})))
  
  (defn- cell-reference
    "Get Excel-style cell reference (e.g., 'A1', 'B5')"
    [^Cell cell]
    (when cell
      (.formatAsString (CellReference. (.getRowIndex cell) (.getColumnIndex cell)))))
  

(defn- extract-sheet-data
  "Extract all data from a single sheet. Always include formula cells, even if value is nil."
  [^Sheet sheet]
  (let [sheet-name (.getSheetName sheet)]
    {:sheet-name sheet-name
     :cells (for [^Row row sheet
                  ^Cell cell row
                  :let [value (cell-value cell)
                        ref (cell-reference cell)]
                  :when (or (some? value)
                            (= (.getCellType cell) CellType/FORMULA))]
              {:cell-ref ref
               :row (.getRowIndex cell)
               :col (.getColumnIndex cell)
               :value value
               :type (str (.getCellType cell))
               :formula (when (= (.getCellType cell) CellType/FORMULA)
                          (.getCellFormula cell))
               :formatted-text (try
                                 (when-let [data-formatter (.getDataFormatter (.getWorkbook (.getSheet cell)))]
                                   (.formatCellValue data-formatter cell))
                                 (catch Exception _ nil))
               :style-index (.getIndex (.getCellStyle cell))
               :number-format (.getDataFormatString (.getCellStyle cell))
               :hyperlink (when-let [hyperlink (.getHyperlink cell)]
                            {:address (.getAddress hyperlink)
                             :label (.getLabel hyperlink)
                             :type (str (.getType hyperlink))})
               :comment (when-let [comment (.getCellComment cell)]
                          {:text (.getString comment)
                           :author (.getAuthor comment)
                           :visible (.isVisible comment)})})}))
  
  (defn extract-data
    "Extract structured data from XLSX file.
    
    Options:
    - :sheets - vector of sheet names to extract (default: all sheets)
    
    Returns a map with sheet data and metadata"
    [file-buffer & {:keys [sheets]}]
    (with-open [workbook (WorkbookFactory/create file-buffer)]
      (let [all-sheets (for [i (range (.getNumberOfSheets workbook))]
                         (.getSheetAt workbook i))
            target-sheets (if sheets
                            (filter #(contains? (set sheets) (.getSheetName %)) all-sheets)
                            all-sheets)
            sheet-data (mapv extract-sheet-data target-sheets)]
        {:file-path "buffer-data"
         :sheets sheet-data
         :sheet-count (count sheet-data)
         :total-cells (reduce + (map #(count (:cells %)) sheet-data))})))
  
  
  (defn create-file-buffer
    [file-path]
    (FileInputStream. (io/file file-path)))
  
  (defn extract-cell-comment
    "Extract comment information from a cell"
    [^Cell cell]
    (when cell
      (when-let [comment (.getCellComment cell)]
        {:comment-text (.getString comment)
         :comment-author (.getAuthor comment)
         :comment-visible (.isVisible comment)})))
  
  (defn extract-hyperlink-info
    "Extract hyperlink information from a cell"
    [^Cell cell]
    (when cell
      (when-let [hyperlink (.getHyperlink cell)]
        {:hyperlink-address (.getAddress hyperlink)
         :hyperlink-label (.getLabel hyperlink)
         :hyperlink-type (str (.getType hyperlink))})))
  
  
  (comment
    ;; --- Apache POI supported formula functions (POI 5.x) ---
    ;; To get the list of supported formula functions in POI 5.x, run this in your REPL:
    ;;
    (require '[clojure.pprint :refer [pprint]])
    (import 'org.apache.poi.ss.formula.function.FunctionMetadataRegistry)
    ;; ;; Get function metadata by name
    (FunctionMetadataRegistry/getFunctionByName "LEFT")
    ;; ;; Get function index by name
    (FunctionMetadataRegistry/lookupIndexByName "LEFT") ; returns -1 if not supported
    ;;`
    ;; If you get ClassNotFoundException, check your POI version and classpath.
    )