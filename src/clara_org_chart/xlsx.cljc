(ns clara-org-chart.xlsx
    "XLSX file parsing and data extraction utilities - Cross-platform version"
   #?(:cljs (:require [clara-org-chart.cljs.xlsx :as xlsx-cljs])
      :clj  (:require [clara-org-chart.xlsx-jvm :as xlsx-jvm])))
  
  #?(:cljs
     (defn extract-data
       "Extract data from XLSX file - ClojureScript version using SheetJS"
       [file-path]
       (xlsx-cljs/extract-data (xlsx-cljs/create-file-buffer file-path)))
  
     :clj
     (defn extract-data
       "Extract data from XLSX file - Clojure version using Apache POI with transducer optimizations"
       [file-path & {:keys [minimal streaming batch-size] :as opts}]
       (->
        (xlsx-jvm/create-file-buffer file-path)
        (xlsx-jvm/extract-data opts))))
  
  
  
  (comment
  
    ;;   (xlsx-jvm/extract-data (xlsx-jvm/create-file-buffer "test/sample_data.xlsx"))
  
  
    (extract-data "test/sample_data.xlsx")
  
    :rcf)